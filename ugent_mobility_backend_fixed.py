# -*- coding: utf-8 -*-
"""
UGent Mobility Backend

PIPELINE OVERVIEW:
This module processes UGent employee mobility data across 6-7 Excel sheets (MOB0-MOB6),
analyzing commute modes, distances, and patterns during a specified period.

The processing pipeline follows these stages:

1. DATA LOADING: Read Excel files (main data + distance reference files).
2. ID STANDARDIZATION: Clean and normalize Pers.nr. and UGent ID fields to ensure
   consistency, removing NaN/None/"nan" entries and handling float IDs (e.g., "123.0" → "123").
3. ID LINKING: Fill missing UGent IDs using Pers.nr. lookups from MOB0 (primary ID map).
   This ensures all rows have a valid UGent ID for downstream merging.
4. DEDUPLICATION: Remove duplicate records per sheet, with special handling:
   - MOB6: dedup by Pers.nr. (has no UGent ID)
   - MOB3: aggregate hours per person, keep highest Uur/wk row first
   - Others: keep first occurrence per UGent ID
5. BUILD MOB4bis: Filter public transit subscriptions to those valid for the entire analysis window.
6. BUILD COMB: Merge transit subscriptions (MOB4) and absence types (MOB5), prioritizing Trein > bus/tram > fiets.
7. BIKE DAYS: Count days per employee marked as "Gefietste dag" in MOB2.
8. BUILD CONCLUSIE: Synthesize final output combining all analyses:
   - Bike commuting frequency (threshold-based)
   - Transit subsidies (OV) and types
   - Transit trajectory participation
   - Distance-based mode inference (auto vs. public transit vs. bike vs. foot)
   - Distance categorization (1km, 5km, 10km, etc.)
   - Email and home address enrichment
9. VALIDATION: Detect anomalies (e.g., distance mismatches with declared home location).

The WHY:
- Standardization ensures consistent ID matching across sheets.
- Deduplication handles multiple contracts/roles per employee.
- Distance-based logic infers commute mode when explicit data missing.
- Comprehensive enrichment enables HR policy reporting on mobility incentives.
"""

import pandas as pd
import numpy as np

REQUIRED_FIELDS = {
    "MOB0": ["Pers.nr.", "UGent ID", "E-mail"],
    "MOB1": ["Pers.nr.", "DSTN1"],
    "MOB1B": ["Pers.nr.", "MOBM1"],
    "MOB2": ["Pers.nr.", "Looncomponent"],
    "MOB3": ["Pers.nr.", "UGent ID", "site", "Uur/wk", "WeekDg", "Werkplek postcode"],
    "MOB4": ["Pers.nr.", "UGent ID", "Toewijzingsnummer", "Contractnr", "Looncomponent"],
    "MOB5": ["Pers.nr.", "UGent ID", "Soort afwezigheid"],
    "MOB6": ["Pers.nr.", "Plaats (laatste dom.)"],
}

AUTO_CATCHALL = "auto, moto, carpool of te voet"

OV_PRIORITY = {
    "Trein":       4,
    "bus/tram":    3,
    "fiets":       2,
    AUTO_CATCHALL: 0,
    "":            0,
}

OV_CODE_MAP = {
    "FTS":  "fiets",
    "EFTS": "fiets",
    "SFTS": "fiets",
    "PEFI": "fiets",
    "GNFI": "fiets",

    "TRN1":                      "Trein",
    "TRN2":                      "Trein",
    "NMBS":                      "Trein",
    "Abonnement NMBS":           "Trein",
    "Abonn. NMBS":               "Trein",
    "Abonn. Andere":             "Trein",
    "Abonnement NMBS Mobib":     "Trein",
    "Abon.flex NMBS 3e bet.":    "Trein",
    "Abon.NMBS+De Lijn 3e bet.": "Trein",
    "Abon. NMBS 3e bet.":        "Trein",
    "Abonn. NMBS+De Lijn":       "Trein",
    "Abonnement NMBS+De Lijn":   "Trein",
    "NMBS+De Lijn":              "Trein",

    "Abonn. MIVB":               "bus/tram",
    "MIVB":                      "bus/tram",
    "Abonnement De Lijn":        "bus/tram",
    "Abonn. De Lijn":            "bus/tram",
    "LIJN":                      "bus/tram",
    "Abonnement DL Mobib":       "bus/tram",
    "Abon. De Lijn 3e bet.":     "bus/tram",
    "Abon.flex De Lijn 3e bet.": "bus/tram",
    "De Lijn":                   "bus/tram",
    "De Ljjn":                   "bus/tram",

    "GFTS": AUTO_CATCHALL,
    "MOTO": AUTO_CATCHALL,
    "TVT":  AUTO_CATCHALL,
    "AUTO": AUTO_CATCHALL,
}


def apply_mapping(data, mapping):
    """Apply column mapping from UI, function takes input from user file."""
    mapped = {}
    for sheet, df in data.items():
        if sheet in mapping:
            ren = {v: k for k, v in mapping[sheet].items() if v}
            df = df.rename(columns=ren)
        for req in REQUIRED_FIELDS.get(sheet, []):
            if req not in df.columns:
                df[req] = ""
        mapped[sheet] = df
    return mapped


class MobilityEngine:
    def __init__(self, main_file, afstanden_file, start, end, bike_threshold=4):
        self.main_file      = main_file
        self.afstanden_file = afstanden_file
        self.start          = pd.to_datetime(start)
        self.end            = pd.to_datetime(end)
        self.bike_threshold = int(bike_threshold)

        self.data            = {}
        self.afst1           = pd.DataFrame()
        self.afst2           = pd.DataFrame()
        self.mob4bis         = pd.DataFrame()
        self.comb            = pd.DataFrame()
        self.conclusie       = pd.DataFrame()
        self.afstand_summary = pd.DataFrame()

    # -----------------------------------------------------------------------
    # LOAD
    # -----------------------------------------------------------------------
    def load_data(self):
        """first step, saves the user entry in data dataframe"""
        xls = pd.ExcelFile(self.main_file)
        self.data = {}
        for s in xls.sheet_names:
            df = pd.read_excel(xls, s)
            df.columns = df.columns.astype(str).str.strip()
            self.data[s] = df

    def load_afstanden_with_mapping(self, map1=None, map2=None):
        """loads the distances excel file and map it"""
        xls    = pd.ExcelFile(self.afstanden_file)
        sheets = xls.sheet_names

        def rename(df, m):
            if not m:
                return df
            r = {}
            if m.get("ID"):
                r[m["ID"]] = "ID"
            if m.get("Distance"):
                r[m["Distance"]] = "Distance"
            return df.rename(columns=r)

        self.afst1 = rename(pd.read_excel(xls, sheets[0]), map1)
        if len(sheets) > 1:
            self.afst2 = rename(pd.read_excel(xls, sheets[1]), map2)

        self.afst1.columns = self.afst1.columns.astype(str).str.strip()
        if not self.afst2.empty:
            self.afst2.columns = self.afst2.columns.astype(str).str.strip()

    # -----------------------------------------------------------------------
    # ID UTILITIES
    # -----------------------------------------------------------------------
    @staticmethod
    def safe_col(df, col):
        """
        Read the first column of dataframe, to deal with UGent ID column and apply cleaning
        
        """
        result = df[col]
        if isinstance(result, pd.DataFrame):
            result = result.iloc[:, 0]
        return result  
    @staticmethod
    def clean(x) -> str:
        """
        Normalise any ID value to a plain string.Also applied to UGent ID col.
        
        """
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return ""
        s = str(x).strip()
        if s.lower() in ("nan", "none", ""):
            return ""
        if s.endswith(".0"):
            try:
                return str(int(float(s)))
            except (ValueError, OverflowError):
                pass
        return s

    def standardize_ids(self):
        """Apply clean() to every Pers.nr. and UGent ID column."""
        for k, df in self.data.items():
            for col in ["Pers.nr.", "UGent ID"]:
                if col in df.columns:
                    series = self.safe_col(df, col)
                    df[col] = series.apply(self.clean)
            self.data[k] = df

    def deduplicate_sheets(self):
        """
        Remove duplicates from every sheet BEFORE any processing.

        MOB6: deduplicate by Pers.nr. (it has NO UGent ID column).
        MOB3: keep highest Uur/wk row, sum all hours per UGent ID.
        MOB2/MOB4/MOB5: skip  multiple valid rows per person.
        All others: dedup by UGent ID.
        """

        #  MOB6 has no UGent ID , deduplicate by Pers.nr. instead
        if "MOB6" in self.data:
            before = len(self.data["MOB6"])
            self.data["MOB6"] = (
                self.data["MOB6"]
                .drop_duplicates(subset="Pers.nr.", keep="first")
                .reset_index(drop=True)
            )
            removed = before - len(self.data["MOB6"])
            if removed:
                print(f"[dedup] MOB6: removed {removed} duplicate Pers.nr. rows")

        #  MOB3 has duplicates of people with multiple contract: keep row with highest Uur/wk per UGent ID ─────────────────
        if "MOB3" in self.data:
            before = len(self.data["MOB3"])
            df3 = self.data["MOB3"].copy()

            uid_col = self.safe_col(df3, "UGent ID")
            df3["UGent ID"] = uid_col.apply(self.clean)
            df3 = df3[df3["UGent ID"] != ""]

            df3["Uur/wk"] = pd.to_numeric(
                self.safe_col(df3, "Uur/wk"), errors="coerce"
            ).fillna(0)

            #keeps site and week days info for the first mention
            # new name before it was assigned.
            if "WeekDg" in df3.columns:
                df3["_weekdg_num"] = pd.to_numeric(
                    self.safe_col(df3, "WeekDg"),   # <-- was "_weekdg_num" before
                    errors="coerce"
                ).fillna(0)
            else:
                df3["_weekdg_num"] = 0

            df3 = df3.sort_values(
                ["UGent ID", "Uur/wk", "_weekdg_num", "site"],
                ascending=[True, False, False, True]
            )

            main = df3.drop_duplicates(subset="UGent ID", keep="first")
            #hours per week are summed for people with double contracts
            total_hours = (
                df3.groupby("UGent ID", as_index=False)["Uur/wk"]
                .sum()
                .rename(columns={"Uur/wk": "_total_hours"})
            )

            df3_agg = main.merge(total_hours, on="UGent ID", how="left")
            df3_agg["Uur/wk"] = df3_agg["_total_hours"]
            df3_agg = df3_agg.drop(columns=["_total_hours", "_weekdg_num"], errors="ignore")

            dupes = df3_agg["UGent ID"].duplicated().sum()
            if dupes > 0:
                print(f"[dedup] MOB3 still has {dupes} duplicate UGent IDs after aggregation")

            self.data["MOB3"] = df3_agg.reset_index(drop=True)
            after = len(self.data["MOB3"])
            print(f"[dedup] MOB3: {before} -> {after} unique persons")

        # ── All other sheets: dedup by UGent ID ──────────────────────────────
        skip = {"MOB2", "MOB3", "MOB4", "MOB5", "MOB6"}
        for sheet, df in self.data.items():
            if sheet in skip:
                continue
            if "UGent ID" not in df.columns:
                continue
            before = len(df)
            self.data[sheet] = (
                df.drop_duplicates(subset="UGent ID", keep="first")
                .reset_index(drop=True)
            )
            removed = before - len(self.data[sheet])
            if removed:
                print(f"[dedup] {sheet}: removed {removed} duplicate UGent ID rows")

    def attach_ids(self):
        """
        Link Pers.nr. → UGent ID.

        similar to look up in excel, attaching ugent ID based on pers.nr.
        """
        id_map = {}

        # MOB0 is the primary Pers.nr. → UGent ID 
        mob0 = self.data.get("MOB0", pd.DataFrame())
        if not mob0.empty and "Pers.nr." in mob0.columns and "UGent ID" in mob0.columns:
            for _, row in mob0.iterrows():
                pid = self.clean(row["Pers.nr."])
                uid = self.clean(row["UGent ID"])
                if pid and uid:
                    id_map[pid] = uid

        # Fill UGent ID where missing using Pers.nr. lookup for the sheets
        for k, df in self.data.items():
            if "Pers.nr." not in df.columns:
                self.data[k] = df
                continue
            linked = df["Pers.nr."].apply(self.clean).map(id_map).fillna("")
            if "UGent ID" in df.columns:
                uid = self.safe_col(df, "UGent ID").apply(self.clean)
                uid = uid.where(uid != "", linked)
                df["UGent ID"] = uid
            else:
                df["UGent ID"] = linked
            self.data[k] = df

    # -----------------------------------------------------------------------
    # MOB4bis
    # -----------------------------------------------------------------------
    def build_mob4bis(self):
        """
        Filter MOB4 to subscriptions valid for the entire analysis window.
        FIX 7: warns about unparseable Toewijzingsnummer values.
        """
        df = self.data["MOB4"].copy()

        def parse(x):
            try:
                return pd.Timestamp(
                    year  = 2000 + int(x[6:8]),
                    month = int(x[3:5]),
                    day   = int(x[0:2])
                )
            except Exception:
                return pd.NaT

        toewijzing       = df["Toewijzingsnummer"].astype(str)
        df["Start_date"] = toewijzing.str.slice(0, 8).apply(parse)
        df["End_date"]   = toewijzing.str.slice(11, 19).apply(parse)

        unparseable = df["Start_date"].isna() | df["End_date"].isna()
        if unparseable.any():
            n = int(unparseable.sum())
            print(f"[build_mob4bis] WARNING: {n} rows in MOB4 have unparseable "
                  f"Toewijzingsnummer and will be excluded.")

        df["VALID"] = (
            df["Start_date"].notna() &
            df["End_date"].notna()   &
            (df["Start_date"] <= self.start) &
            (df["End_date"]   >= self.end)
        )
        self.mob4bis = df[df["VALID"]].copy()

    # -----------------------------------------------------------------------
    # COMB
    # -----------------------------------------------------------------------
    def build_comb(self):
        rows = []

        if not self.mob4bis.empty:
            tmp = self.mob4bis[["UGent ID", "Looncomponent"]].copy()
            tmp = tmp.rename(columns={"Looncomponent": "OV_type"})
            rows.append(tmp)

        mob5 = self.data.get("MOB5", pd.DataFrame())
        if not mob5.empty and "Soort afwezigheid" in mob5.columns:
            tmp = mob5[["UGent ID", "Soort afwezigheid"]].copy()
            tmp = tmp.rename(columns={"Soort afwezigheid": "OV_type"})
            rows.append(tmp)

        if not rows:
            self.comb = pd.DataFrame(columns=["UGent ID", "OV_type"])
            return

        combined = pd.concat(rows, ignore_index=True)
        combined = combined.loc[:, ~combined.columns.duplicated()]
        combined["UGent ID"] = combined["UGent ID"].apply(self.clean)
        combined = combined[combined["UGent ID"] != ""]
        combined["OV_type"] = combined["OV_type"].fillna("").astype(str).str.strip()

        combined["_std_mode"] = combined["OV_type"].map(OV_CODE_MAP).fillna(combined["OV_type"])
        combined["_priority"] = combined["_std_mode"].map(OV_PRIORITY).fillna(1)

        combined = (combined
                    .sort_values("_priority", ascending=False)
                    .drop_duplicates(subset="UGent ID", keep="first")
                    .drop(columns=["_std_mode", "_priority"]))

        dupes = combined["UGent ID"].duplicated().sum()
        if dupes > 0:
            print(f"[build_comb] {dupes} duplicate UGent IDs remain — investigate")

        self.comb = combined.reset_index(drop=True)

    # -----------------------------------------------------------------------
    # BIKE
    # -----------------------------------------------------------------------
    def bike_days(self):
        df = self.data["MOB2"].copy()

        # FIX A applied: safe_col now reliably returns a Series
        df["UGent ID"] = self.safe_col(df, "UGent ID").apply(self.clean)
        df = df[df["UGent ID"] != ""]
        df["Looncomponent"] = df["Looncomponent"].astype(str).str.strip()
        df = df[df["Looncomponent"] == "Gefietste dag"]

        self.bike = df.groupby("UGent ID").size().reset_index(name="bike_days")

        dupes = self.bike["UGent ID"].duplicated().sum()
        if dupes > 0:
            print(f"[bike_days] {dupes} duplicate UGent IDs — investigate")

    # -----------------------------------------------------------------------
    # DISTANCE
    # -----------------------------------------------------------------------
    def add_afstand(self, df):
        def make_dict(d):
            if d.empty:
                return {}
            cols_lower = {c: c.lower().strip() for c in d.columns}

            id_col = next(
                (c for c, cl in cols_lower.items()
                 if cl in ("id", "ugent id", "pers.nr.") or "ugent" in cl),
                None
            )
            dist_col = next(
                (c for c, cl in cols_lower.items()
                 if "dist" in cl or "afstand" in cl or cl == "km"),
                None
            )
            if not id_col or not dist_col:
                return {}

            tmp = d[[id_col, dist_col]].copy()
            tmp["_uid"] = tmp[id_col].apply(self.clean)
            tmp[dist_col] = pd.to_numeric(tmp[dist_col], errors="coerce")
            tmp = tmp.dropna(subset=[dist_col])
            tmp = tmp[tmp["_uid"] != ""]
            tmp = tmp.drop_duplicates(subset="_uid", keep="first")
            return dict(zip(tmp["_uid"], tmp[dist_col]))

        d1 = make_dict(self.afst1)
        d2 = make_dict(self.afst2)
        combined = {**d2, **d1}

        df["afstand"] = df["UGent ID"].apply(
            lambda x: combined.get(self.clean(x), None)
        )
        return df

    @staticmethod
    def add_afstand_cat(df):
        def cat(v):
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return "N/A"
            v = float(v)
            if v <  1:  return "<1km"
            if v <  5:  return "<5km"
            if v < 10:  return "<10km"
            if v < 15:  return "<15km"
            if v < 20:  return "<20km"
            if v < 50:  return "<50km"
            return ">>>50km"
        df["afstandscategorie"] = df["afstand"].apply(cat)
        return df

    # -----------------------------------------------------------------------
    # FINAL CONCLUSIE
    # -----------------------------------------------------------------------
    def build_conclusie(self):
        df = self.data["MOB3"].copy()
        dupes = df["UGent ID"].duplicated().sum()
        if dupes > 0:
            print(f"[build_conclusie] {dupes} duplicate UGent IDs in MOB3 — investigate")

        # --- Bike ---
        df = df.merge(self.bike, on="UGent ID", how="left").fillna({"bike_days": 0})
        df["aantal_fietsdagen"] = df["bike_days"].astype(int)
        df["Fietser"] = np.where(df["bike_days"] >= self.bike_threshold, "ja", "nee")

        # --- OV from COMB ---
        df["OV_betaald"] = np.where(df["UGent ID"].isin(self.comb["UGent ID"]), "ja", "nee")
        ov_map       = dict(zip(self.comb["UGent ID"], self.comb["OV_type"]))
        df["welke_OV"] = df["UGent ID"].map(ov_map).fillna("")

        # --- Trajecten (MOB1, MOB1B) ---
        # FIX D consequence: MOB1 and MOB1B had no UGent ID originally.
        # attach_ids() fills it from MOB0 via Pers.nr. Check if present.
        mob1  = self.data.get("MOB1", pd.DataFrame()).copy()
        mob1b = self.data.get("MOB1B", pd.DataFrame()).copy()

        if not mob1.empty:
            mob1["UGent ID"] = self.safe_col(mob1, "UGent ID").apply(self.clean) \
                if "UGent ID" in mob1.columns \
                else mob1["Pers.nr."].apply(self.clean)

        if not mob1b.empty:
            mob1b["UGent ID"] = self.safe_col(mob1b, "UGent ID").apply(self.clean) \
                if "UGent ID" in mob1b.columns \
                else mob1b["Pers.nr."].apply(self.clean)

        # FIX 5: any row with a non-null DSTN1 flags as "ja"
        if not mob1.empty and "DSTN1" in mob1.columns:
            traj_any = (mob1.groupby("UGent ID")["DSTN1"]
                            .apply(lambda s: s.notna().any()))
            df["OV_in_traject"] = (df["UGent ID"].map(traj_any)
                                   .map({True: "ja", False: "nee"})
                                   .fillna("nee"))
        else:
            df["OV_in_traject"] = "nee"

        # FIX 6: first non-empty MOBM1 per person
        if not mob1b.empty and "MOBM1" in mob1b.columns:
            mob1b_clean = (
                mob1b[mob1b["MOBM1"].notna() &
                      (mob1b["MOBM1"].astype(str).str.strip() != "")]
                .drop_duplicates(subset="UGent ID", keep="first")
            )
            hoofd_map = dict(zip(mob1b_clean["UGent ID"], mob1b_clean["MOBM1"]))
        else:
            hoofd_map = {}
        df["soort_hoofdtraject"] = df["UGent ID"].map(hoofd_map).fillna("")

        # --- Flags ---
        df["OV_any"]      = df["OV_betaald"]
        df["Fiets_en_OV"] = np.where(
            (df["Fietser"] == "ja") & (df["OV_any"] == "ja"), "ja", "nee")
        df["enkel_fiets"] = np.where(
            (df["Fiets_en_OV"] == "nee") & (df["Fietser"] == "ja"), "ja", "nee")
        df["ov_type_comb"] = np.where(df["OV_betaald"] == "ja", df["welke_OV"], "nee")

        # --- ts1 ---
        def ts1(r):
            if r["enkel_fiets"] == "ja":
                return "fiets"
            if r["OV_betaald"] == "ja":
                return r["welke_OV"] if r["welke_OV"] != "" else "OV unknown"
            return AUTO_CATCHALL

        df["ts1"] = df.apply(ts1, axis=1)

        # --- ts2 ---
        df["ts2"] = df["ts1"]

        # --- ts3 ---
        def map_code(code):
            code = str(code).strip()
            if code in OV_CODE_MAP:
                return OV_CODE_MAP[code]
            if code in ("fiets", "Trein", "bus/tram", AUTO_CATCHALL, "OV unknown"):
                return code
            cl = code.lower()
            if "nmbs" in cl or "trein" in cl:
                return "Trein"
            if "lijn" in cl or "tram" in cl or "bus" in cl or "mivb" in cl:
                return "bus/tram"
            if "fiets" in cl or "fts" in cl:
                return "fiets"
            if "auto" in cl or "moto" in cl or "voet" in cl:
                return AUTO_CATCHALL
            return code

        df["ts3"] = df["ts2"].apply(map_code)

        # --- tmp_mode ---
        df["tmp_mode"] = np.where(
            df["ts3"].isin(["Abonn. NMBS+De Lijn", "Abonnement NMBS+De Lijn", "NMBS+De Lijn"]),
            "Trein",
            df["ts3"]
        )

        # --- Distance ---
        df = self.add_afstand(df)
        df = self.add_afstand_cat(df)

        # --- Final mode ---
        def final_mode(r):
            val = str(r["tmp_mode"]).strip()
            if val != AUTO_CATCHALL:
                return val
            dist = r["afstand"]
            if dist is None or (isinstance(dist, float) and np.isnan(dist)):
                return "auto (geen afstand)"
            return "te voet" if float(dist) < 1.0 else "auto"

        df["vervoerswijze"] = df.apply(final_mode, axis=1)

        # --- FOD ---
        def fod(v):
            try:
                v = int(float(str(v).strip()))
                return v if 1000 <= v < 10000 else 9876
            except Exception:
                return 9876

        df["FOD"] = df["Werkplek postcode"].apply(fod)

        # --- Email + woonplaats ---
        mob0 = self.data.get("MOB0", pd.DataFrame())
        mob6 = self.data.get("MOB6", pd.DataFrame())

        emap = (dict(zip(mob0["UGent ID"].apply(self.clean), mob0["E-mail"]))
                if not mob0.empty and "UGent ID" in mob0.columns else {})

        # FIX D: MOB6 has no UGent ID — build wmap via Pers.nr. → UGent ID
        # lookup using the same id_map used in attach_ids.
        if not mob6.empty and "Pers.nr." in mob6.columns and "Plaats (laatste dom.)" in mob6.columns:
            mob0_map = {}
            if not mob0.empty and "Pers.nr." in mob0.columns and "UGent ID" in mob0.columns:
                mob0_map = dict(zip(
                    mob0["Pers.nr."].apply(self.clean),
                    mob0["UGent ID"].apply(self.clean)
                ))
            mob6_with_uid = mob6.copy()
            mob6_with_uid["_uid"] = mob6_with_uid["Pers.nr."].apply(self.clean).map(mob0_map).fillna("")
            mob6_with_uid = mob6_with_uid[mob6_with_uid["_uid"] != ""]
            wmap = dict(zip(mob6_with_uid["_uid"], mob6_with_uid["Plaats (laatste dom.)"]))
        else:
            wmap = {}

        df["Email"]      = df["UGent ID"].map(emap).fillna("")
        df["woonplaats"] = df["UGent ID"].map(wmap).fillna("")

        keep = [
            "UGent ID", "Email", "woonplaats", "site", "Uur/wk", "WeekDg",
            "Werkplek postcode", "aantal_fietsdagen", "Fietser",
            "OV_betaald", "welke_OV", "OV_in_traject", "soort_hoofdtraject",
            "OV_any", "Fiets_en_OV", "enkel_fiets", "ov_type_comb",
            "ts1", "ts2", "ts3", "tmp_mode",
            "vervoerswijze", "FOD", "afstand", "afstandscategorie",
        ]
        self.conclusie = df[[c for c in keep if c in df.columns]].copy()

        # --- Distance summary ---
        total = float(len(self.conclusie))
        if total > 0:
            vc = self.conclusie["afstandscategorie"].value_counts().reset_index()
            vc.columns = ["afstandscategorie", "Aantal"]
            vc["Percentage"] = (vc["Aantal"].astype(float) / total * 100).round(1)
            self.afstand_summary = vc

    # -----------------------------------------------------------------------
    # DISTANCE VALIDATION
    # -----------------------------------------------------------------------
    def validate_distances(self) -> pd.DataFrame:
        if self.conclusie.empty:
            return pd.DataFrame()

        df     = self.conclusie.copy()
        issues = []

        def _flag(mask, label):
            subset = df[mask].copy()
            if not subset.empty:
                subset["issue"] = label
                issues.append(subset[[
                    "UGent ID", "woonplaats", "afstand",
                    "afstandscategorie", "vervoerswijze", "issue"
                ]])

        def _dist_ok(s):
            return (s["afstand"] is not None and
                    not (isinstance(s["afstand"], float) and np.isnan(s["afstand"])))

        _flag(df["vervoerswijze"] == "auto (geen afstand)",
              "Geen afstand in afstandenbestand")
        _flag(df.apply(lambda r: (
                "gent" in str(r["woonplaats"]).lower() and
                _dist_ok(r) and float(r["afstand"]) > 20), axis=1),
              "Woont in Gent maar afstand >20 km — check afstandenbestand")
        _flag(df.apply(lambda r: (
                "drongen" in str(r["woonplaats"]).lower() and
                _dist_ok(r) and
                (float(r["afstand"]) > 15 or float(r["afstand"]) < 2)), axis=1),
              "Woont in Drongen maar afstand buiten verwacht bereik 2–15 km")
        _flag(df.apply(lambda r: _dist_ok(r) and float(r["afstand"]) > 100, axis=1),
              "Afstand >100 km — waarschijnlijk datafout")
        _flag(df.apply(lambda r: _dist_ok(r) and 0 < float(r["afstand"]) < 0.1, axis=1),
              "Afstand <0.1 km — mogelijk nulwaarde placeholder")

        if not issues:
            return pd.DataFrame(columns=[
                "UGent ID", "woonplaats", "afstand",
                "afstandscategorie", "vervoerswijze", "issue"
            ])

        return (pd.concat(issues, ignore_index=True)
                  .drop_duplicates(subset=["UGent ID", "issue"])
                  .sort_values(["issue", "woonplaats"])
                  .reset_index(drop=True))

    # -----------------------------------------------------------------------
    # PIVOT + RUN
    # -----------------------------------------------------------------------
    def get_pivot_table(self):
        if self.conclusie.empty:
            return pd.DataFrame()
        return pd.crosstab(
            self.conclusie["site"],
            self.conclusie["vervoerswijze"],
            margins=True, margins_name="Total"
        ).fillna(0).astype(int)

    def run(self, map1=None, map2=None):
        self.load_data()
        self.load_afstanden_with_mapping(map1, map2)

        self.standardize_ids()   # step 1: clean existing IDs
        self.attach_ids()        # step 2: fill missing IDs from MOB0
        self.standardize_ids()   # step 3: clean again after filling
        self.deduplicate_sheets()# step 4: dedup now that IDs are stable

        self.build_mob4bis()
        self.build_comb()
        self.bike_days()
        self.build_conclusie()

        return self.conclusie

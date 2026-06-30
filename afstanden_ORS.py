import pandas as pd
import openrouteservice
import time
import os

"""
afstanden_ORS.py

Purpose
-------
Compute driving distances (in kilometers) between pairs of coordinates using
the OpenRouteService (ORS) directions API for rows in an Excel input file,
and save the results back to an Excel output file.

High-level behavior
-------------------
- Reads an Excel file containing coordinate columns: lat_1, long_1, lat_2, long_2
  (and a string ID column 'UGent ID').
- For each row where 'travel_km' is missing, calls the ORS directions API to
  compute the driving-car route distance between the two points.
- Caches results in memory to avoid duplicate API calls for the same coordinate
  pair during a single run.
- Respects a simple rate-control delay between calls and handles basic 429
  rate-limit responses by sleeping.
- Periodically saves progress to the output Excel file so work is not lost if
  the script is interrupted.

Usage
-----
- Replace the API_KEY placeholder with your ORS API key.
- Update INPUT_FILE and OUTPUT_FILE paths to the correct files on your system.
- Required packages: pandas, openrouteservice
- Run as a simple script: python afstanden_ORS.py

Important notes
---------------
- Keep your API key secret. Do not commit real keys to public repos.
- The cache is in-memory only and is not persisted between runs.
- The script expects columns lat_1, long_1, lat_2, long_2, and 'UGent ID' in the
  input Excel file. It adds/uses a column named 'travel_km' for results.
"""

# 🔑 YOUR API KEY HERE — replace with your ORS key (keep it secret)
API_KEY = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6Ijk5MDIzOGZlMDcwYjQ2YzA4ZTY4ZjBiMGQ0MGJkNTg1IiwiaCI6Im11cm11cjY0In0="  # <-- REPLACE THIS

client = openrouteservice.Client(key=API_KEY)

# Edit these paths to point to your files
INPUT_FILE = "D:/data_folders/work/Mobility/Finalafstandeninput_mei2026.xlsx"
OUTPUT_FILE = "D:/data_folders/work/Mobility/AftsandenresultsFinal4.xlsx"

# Load file
df1 = pd.read_excel(INPUT_FILE, dtype={"UGent ID": str})
df1.columns = df1.columns.str.strip()

# If rerunning, keep existing results
if os.path.exists(OUTPUT_FILE):
    df_existing = pd.read_excel(OUTPUT_FILE)
    if 'travel_km' in df_existing.columns:
        # If the output exists and has a travel_km column, reuse it to avoid re-requesting
        df1['travel_km'] = df_existing['travel_km']
    else:
        df1['travel_km'] = None
else:
    df1['travel_km'] = None

# Cache to avoid duplicate API calls during this run.
# Key: (lat_1, long_1, lat_2, long_2) -> distance in km
cache = {}

# Simple rate control: timestamp of last API call (seconds since epoch)
LAST_CALL = 0

def wait_if_needed():
    """
    Enforce a minimum delay between API calls.

    This function computes the elapsed time since the last recorded API call
    and sleeps for the remaining time if the minimum interval is not yet
    satisfied. It then updates LAST_CALL to the current time.

    The current minimum interval is 0.5 seconds (configurable by editing the
    constant in the function). This is a simple approach to reduce the risk
    of hitting rate limits; ORS may have other quotas or burst rules.
    """
    global LAST_CALL
    now = time.time()
    elapsed = now - LAST_CALL

    # Minimum interval between calls in seconds
    min_interval = 0.5

    if elapsed < min_interval:
        time.sleep(min_interval - elapsed)

    LAST_CALL = time.time()

def get_route_distance(row):
    """
    Given a DataFrame row with coordinate fields, return the driving distance (km)
    between the two points using the OpenRouteService directions API.

    Parameters
    ----------
    row : pandas.Series
        Expected fields: 'lat_1', 'long_1', 'lat_2', 'long_2', and optionally
        'UGent ID' for logging.

    Returns
    -------
    float or None
        Distance in kilometers if successful, otherwise None (on error or missing coords).

    Behavior
    --------
    - If any coordinate is missing (NaN), returns None.
    - Uses an in-memory cache to avoid duplicate API calls for identical coordinate pairs.
    - Retries up to 3 times on general exceptions. If the exception message contains
      '429', it sleeps for 10 seconds before retrying (to handle rate limiting).
    """
    global cache

    # Skip if coordinates missing
    if pd.isna(row.get('lat_1')) or pd.isna(row.get('long_1')) or pd.isna(row.get('lat_2')) or pd.isna(row.get('long_2')):
        return None

    # Use tuple of coordinates as cache key. Keep the order consistent.
    key = (row['lat_1'], row['long_1'], row['lat_2'], row['long_2'])

    if key in cache:
        return cache[key]

    # ORS expects coords as (lon, lat)
    coords = [
        (row['long_1'], row['lat_1']),
        (row['long_2'], row['lat_2'])
    ]

    for attempt in range(3):
        try:
            wait_if_needed()
            route = client.directions(coords, profile='driving-car')
            # route['routes'][0]['summary']['distance'] is in meters
            dist = route['routes'][0]['summary']['distance'] / 1000.0
            cache[key] = dist
            return dist

        except Exception as e:
            # 429 responses (rate limiting) are handled by sleeping and retrying
            if "429" in str(e):
                print("⏳ Rate limit hit → sleeping 10 seconds...")
                time.sleep(10)
            else:
                # For other exceptions, log and return None for this row
                print(f"❌ Error for ID {row.get('UGent ID', 'unknown')}: {e}")
                return None

    # If all retries failed
    return None

# Main processing loop: process rows missing travel_km and save progress periodically.
for i, row in df1.iterrows():

    # Skip rows that already have a travel_km value
    if pd.notna(df1.at[i, 'travel_km']):
        continue  # already done

    print(f"Processing row {i+1}/{len(df1)}")

    df1.at[i, 'travel_km'] = get_route_distance(row)

    # Save progress to avoid losing work. Adjust save frequency as needed.
    if (i + 1) % 100 == 0:
        df1.to_excel(OUTPUT_FILE, index=False)
        print("💾 Progress saved")

# Final save
df1.to_excel(OUTPUT_FILE, index=False)
print("✅ Done")

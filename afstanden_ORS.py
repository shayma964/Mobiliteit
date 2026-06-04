import pandas as pd
import openrouteservice
import time
import os

# 🔑 YOUR API KEY HEre
API_KEY = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6Ijk5MDIzOGZlMDcwYjQ2YzA4ZTY4ZjBiMGQ0MGJkNTg1IiwiaCI6Im11cm11cjY0In0="  # <-- REPLACE THIS

client = openrouteservice.Client(key=API_KEY)

INPUT_FILE = "D:/data_folders/work/Mobility/Finalafstandeninput_mei2026.xlsx"
OUTPUT_FILE = "D:/data_folders/work/Mobility/AftsandenresultsFinal4.xlsx"

# Load file
df1 = pd.read_excel(INPUT_FILE, dtype={"UGent ID": str})
df1.columns = df1.columns.str.strip()

# If rerunning, keep existing results
if os.path.exists(OUTPUT_FILE):
    df_existing = pd.read_excel(OUTPUT_FILE)
    if 'travel_km' in df_existing.columns:
        df1['travel_km'] = df_existing['travel_km']
    else:
        df1['travel_km'] = None
else:
    df1['travel_km'] = None

# Cache to avoid duplicate API calls
cache = {}

# Rate control
LAST_CALL = 0

def wait_if_needed():
    global LAST_CALL
    now = time.time()
    elapsed = now - LAST_CALL
    
    
    if elapsed < 0.5:   
        time.sleep(0.5 - elapsed)
    
    LAST_CALL = time.time()

def get_route_distance(row):
    global cache
    
    if pd.isna(row['lat_1']) or pd.isna(row['long_1']) or pd.isna(row['lat_2']) or pd.isna(row['long_2']):
        return None

    key = (row['lat_1'], row['long_1'], row['lat_2'], row['long_2'])

    if key in cache:
        return cache[key]

    coords = [
        (row['long_1'], row['lat_1']),
        (row['long_2'], row['lat_2'])
    ]

    for attempt in range(3):
        try:
            wait_if_needed()
            route = client.directions(coords, profile='driving-car')
            dist = route['routes'][0]['summary']['distance'] / 1000
            cache[key] = dist
            return dist

        except Exception as e:
            if "429" in str(e):
                print("⏳ Rate limit hit → sleeping 10 seconds...")
                time.sleep(10)
            else:
                print(f"❌ Error for ID {row['UGent ID']}: {e}")
                return None

    return None

# Process only missing rows
for i, row in df1.iterrows():
    
    if pd.notna(df1.at[i, 'travel_km']):
        continue  # already done

    print(f"Processing row {i+1}/{len(df1)}")

    df1.at[i, 'travel_km'] = get_route_distance(row)

    # Save every 100 rows
    if (i + 1) % 100 == 0:
        df1.to_excel(OUTPUT_FILE, index=False)
        print("💾 Progress saved")

# Final save
df1.to_excel(OUTPUT_FILE, index=False)
print("✅ Done")

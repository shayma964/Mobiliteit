import pandas as pd
from geopy.distance import geodesic
import numpy as np

df1 = pd.read_excel("D:/data_folders/work/Mobility/Finalafstandeninput_mei2026.xlsx", 
                   dtype={"UGent ID": str})

df1.columns = df1.columns.str.strip()

def calc_distance(row):
    # Check if ID is missing
    if pd.isna(row['UGent ID']):
        print("Warning: Row with missing ID found, skipping...")
        return None
    
    # Check if coordinates are missing
    try:
        # Check for NaN values in coordinates
        if pd.isna(row['lat_1']) or pd.isna(row['long_1']) or pd.isna(row['lat_2']) or pd.isna(row['long_2']):
            print(f"Warning: Missing coordinates for ID {row['UGent ID']}")
            return None
            
        point1 = (row['lat_1'], row['long_1'])
        point2 = (row['lat_2'], row['long_2'])
        return geodesic(point1, point2).km
        
    except Exception as e:
        print(f"Error calculating distance for ID {row.get('UGent ID', 'Unknown')}: {e}")
        return None

# Apply with better error handling
df1['afstand_km'] = df1.apply(calc_distance, axis=1)

# Check how many missing values you have
print(f"Total rows: {len(df1)}")
print(f"Rows with missing ID: {df1['UGent ID'].isna().sum()}")
print(f"Rows with missing distance: {df1['afstand_km'].isna().sum()}")

# Save results
df1.to_excel("D:/data_folders/work/Mobility/AftsandenresultsFinal.xlsx", index=False)

"""
Distance calculation module for analyzing travel distances using coordinates.

This module reads mobility data from an Excel file, calculates straight-line
(as-the-crow-flies) distances between two coordinate pairs using the geodesic
method, and saves the results to a new Excel file. It includes error handling
for missing or invalid coordinate data.

Example:
    The script expects an Excel file with the following columns:
        - UGent ID: Unique identifier for the employee
        - lat_1, long_1: Starting point latitude and longitude
        - lat_2, long_2: Ending point latitude and longitude

    Results are saved with an additional 'afstand_km' column containing
    the calculated distances in kilometers.
"""

import pandas as pd
from geopy.distance import geodesic
import numpy as np

df1 = pd.read_excel("D:/data_folders/work/Mobility/Finalafstandeninput_mei2026.xlsx", 
                   dtype={"UGent ID": str})

df1.columns = df1.columns.str.strip()

def calc_distance(row):
    """
    Calculate the geodesic distance between two geographic coordinates.

    This function computes the straight-line distance (as-the-crow-flies)
    between two points specified by their latitude and longitude coordinates.
    It includes validation checks for missing data and handles errors gracefully.

    Parameters
    ----------
    row : pd.Series
        A pandas Series containing the following keys:
            - UGent ID : str
                Unique identifier for the record (used for error logging)
            - lat_1 : float
                Latitude of the first point
            - long_1 : float
                Longitude of the first point
            - lat_2 : float
                Latitude of the second point
            - long_2 : float
                Longitude of the second point

    Returns
    -------
    float or None
        The distance in kilometers between the two points, or None if:
        - The UGent ID is missing
        - Any of the coordinate values are NaN
        - An exception occurs during calculation

    Notes
    -----
    - Coordinates should be in decimal degrees (latitude: -90 to 90, 
      longitude: -180 to 180)
    - Uses the Vincenty distance formula via geopy for accuracy
    - Missing data is logged as warnings to aid in data quality assessment

    Examples
    --------
    >>> row = pd.Series({
    ...     'UGent ID': 'ID001',
    ...     'lat_1': 51.2194,
    ...     'long_1': 4.4024,
    ...     'lat_2': 50.8503,
    ...     'long_2': 4.3517
    ... })
    >>> calc_distance(row)
    40.2
    """
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

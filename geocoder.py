from geopy.geocoders import Nominatim
import pandas as pd
import time
import numpy as np
import os
import pickle
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
import socket

# INPUT ADRESSES FILE
df = pd.read_excel("D:/data_folders/work/Mobility/Newwoonadressen_uncoded.xlsx") ##change to location in your computer
OUTPUT_FILE = "D:/data_folders/work/Mobility/newadressewoon2_coded.xlsx"

USER_AGENTS = [
    "my_geocoder_app1",
    "my_geocoder_app2", 
    "my_geocoder_app3",
    "my_geocoder_app4",
    "my_geocoder_app5"
]

# Cache file
CACHE_FILE = "D:/data_folders/work/Mobility/geocode_cache.pkl"
CHECKPOINT_FILE = "D:/data_folders/work/Mobility/checkpoint.pkl"


# Load existing cache if available
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, 'rb') as f:
        cache = pickle.load(f)
    print(f"Loaded {len(cache)} cached addresses")
else:
    cache = {}
    print("No existing cache found")

# Load checkpoint to see where we left off
start_index = 0
if os.path.exists(CHECKPOINT_FILE):
    with open(CHECKPOINT_FILE, 'rb') as f:
        checkpoint = pickle.load(f)
        start_index = checkpoint.get('last_index', 0)
        # Merge any cached results from checkpoint
        cache.update(checkpoint.get('cache', {}))
    print(f"Resuming from row {start_index}")

# Function to create geocoder with different user agent
def get_geocoder(agent_index):
    return Nominatim(user_agent=USER_AGENTS[agent_index % len(USER_AGENTS)])

def geocode_address(row, agent_index):
    try:
        # Convert to string and handle potential NaN/None values
        address = str(row['address']).strip() if pd.notna(row['address']) else ""
        city = str(row['city']).strip() if pd.notna(row['city']) else ""
        country = str(row['country']).strip() if pd.notna(row['country']) else ""
        
        # Create unique key for this address
        address_key = f"{address}_{city}_{country}"
        
        # Skip if all fields are empty
        if not address and not city and not country:
            print("Skipping empty address row")
            return pd.Series([None, None])
        
        # Check cache first
        if address_key in cache:
            print(f"Cache hit: {address_key}")
            return pd.Series(cache[address_key])
        
        # Build address string for geocoding
        address_parts = []
        if address:
            address_parts.append(address)
        if city:
            address_parts.append(city)
        if country:
            address_parts.append(country)
        
        full_address = ", ".join(address_parts)
        print(f"[Agent {agent_index}] Geocoding: {full_address}")
        
        # Get geocoder with specific user agent
        geolocator = get_geocoder(agent_index)
        
        # Try geocoding with retry logic
        max_retries = 3
        for attempt in range(max_retries):
            try:
                location = geolocator.geocode(full_address)
                time.sleep(1.1)  # Slightly more than 1 second to be safe
                
                if location:
                    result = (location.latitude, location.longitude)
                    cache[address_key] = result
                    return pd.Series(result)
                else:
                    # No result found, cache the failure too
                    cache[address_key] = (None, None)
                    return pd.Series([None, None])
                    
            except (GeocoderTimedOut, GeocoderServiceError, socket.gaierror) as e:
                print(f"Attempt {attempt + 1} failed: {e}")
                if attempt < max_retries - 1:
                    wait_time = (2 ** attempt) * 5  # Exponential backoff
                    print(f"Waiting {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    print(f"All attempts failed for {full_address}")
                    cache[address_key] = (None, None)
                    return pd.Series([None, None])
                    
    except Exception as e:
        print(f"Error processing row: {e}")
        return pd.Series([None, None])
    
    return pd.Series([None, None])

# DIVIDE THE DATA INTO CHUNKS for different user agents
print(f"Total rows: {len(df)}")

# Split data into chunks (one per user agent)
n_agents = len(USER_AGENTS)
chunk_size = len(df) // n_agents
chunks = []
start_idx = 0

for i in range(n_agents):
    if i < n_agents - 1:
        chunk = df.iloc[start_idx:start_idx + chunk_size]
    else:
        # Last chunk gets remaining rows
        chunk = df.iloc[start_idx:]
    
    chunks.append(chunk)
    print(f"Agent {i} chunk size: {len(chunk)} rows")
    start_idx += chunk_size

# Process chunks with different user agents
all_results = []

for agent_idx, chunk in enumerate(chunks):
    print(f"\n{'='*50}")
    print(f"Processing chunk {agent_idx + 1}/{n_agents} with agent {USER_AGENTS[agent_idx]}")
    print(f"{'='*50}")
    
    # Initialize lat/lon columns if they don't exist
    if 'lat1' not in chunk.columns:
        chunk['lat1'] = None
    if 'lon1' not in chunk.columns:
        chunk['lon1'] = None
    
    # Process each row in the chunk
    for idx, row in chunk.iterrows():
        try:
            # Skip if already processed (has coordinates)
            if pd.notna(row.get('lat1')) and pd.notna(row.get('lon1')):
                print(f"Row {idx} already processed, skipping")
                continue
            
            # Geocode the address
            result = geocode_address(row, agent_idx)
            chunk.at[idx, 'lat1'] = result[0]
            chunk.at[idx, 'lon1'] = result[1]
            
            # Save checkpoint every 10 rows
            if idx % 10 == 0:
                checkpoint = {
                    'last_index': idx,
                    'cache': cache
                }
                with open(CHECKPOINT_FILE, 'wb') as f:
                    pickle.dump(checkpoint, f)
                
                # Also save partial results
                temp_df = pd.concat([chunk for chunk in chunks if chunk is not None] + [chunk])
                temp_df.to_excel(OUTPUT_FILE.replace('.xlsx', '_partial.xlsx'), index=False)
                print(f"Checkpoint saved at row {idx}")
                
        except KeyboardInterrupt:
            print(f"\nInterrupted at row {idx}")
            # Save progress
            checkpoint = {
                'last_index': idx,
                'cache': cache
            }
            with open(CHECKPOINT_FILE, 'wb') as f:
                pickle.dump(checkpoint, f)
            
            # Save partial results
            temp_df = pd.concat([chunk for chunk in chunks if chunk is not None])
            temp_df.to_excel(OUTPUT_FILE.replace('.xlsx', '_interrupted.xlsx'), index=False)
            print("Progress saved. You can resume later.")
            break
    
    # Add small delay between chunks to be extra safe
    if agent_idx < len(chunks) - 1:
        print(f"\nWaiting 60 seconds before next chunk...")
        time.sleep(60)

# Combine all chunks back into one dataframe
final_df = pd.concat(chunks, axis=0)

# Sort by original index to restore order
final_df = final_df.sort_index()

# Save final results
final_df.to_excel(OUTPUT_FILE, index=False)
print(f"\nFinal results saved to: {OUTPUT_FILE}")

# Save cache for future runs
with open(CACHE_FILE, 'wb') as f:
    pickle.dump(cache, f)
print(f"Cache saved to: {CACHE_FILE}")

# Print summary
successful = final_df['lat1'].notna().sum()
print(f"\n{'='*50}")
print(f"SUMMARY:")
print(f"{'='*50}")
print(f"Total rows processed: {len(final_df)}")
print(f"Successfully geocoded: {successful}")
print(f"Failed/Empty: {len(final_df) - successful}")
print(f"Unique addresses cached: {len(cache)}")

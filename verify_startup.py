import streamlit as st
import sys
import os

print("Starting verification (Deep)...")

# 1. Check Secrets
try:
    print("Checking secrets...")
    secrets = st.secrets
    if "GEMINI_API_KEY" in secrets:
        print(f"GEMINI_API_KEY found: {len(secrets['GEMINI_API_KEY'])} chars")
    
    if "supabase" in secrets:
        print("supabase section found")
        # Ensure 'url' and 'key' are present
        if 'url' in secrets['supabase'] and 'key' in secrets['supabase']:
             print("supabase url/key present")
        else:
             print("supabase url/key MISSING in section")
    else:
        print("supabase section NOT found")

except Exception as e:
    print(f"Error loading secrets: {e}")
    sys.exit(1)

# 2. Check Database Connection
try:
    print("Checking database connection...")
    from modules.database import fetch_table
    from modules.constants import MAP_PERSONS
    
    print("Fetching 'persons' table...")
    df = fetch_table("persons", MAP_PERSONS)
    print(f"Fetch successful. Rows: {len(df)}")

except Exception as e:
    print(f"Error interacting with database: {e}")
    # Print full traceback
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("Verification complete. DB seems accessible.")

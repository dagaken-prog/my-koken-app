import streamlit as st
import pandas as pd
from modules.database import fetch_table, insert_data, delete_data
from modules.constants import MAP_ACTIVITIES
import time

# Mock secrets if needed? No, running in same env should have access to secrets.toml
# But I am running via run_command, which might not load streamlit secrets automatically unless run with `streamlit run`.
# I should run this using `streamlit run debug_db.py`.

st.set_page_config(layout="wide")
st.title("Debug DB")

if st.button("Test Insert Cycle"):
    # 1. Fetch initial count
    df_before = fetch_table("activities", MAP_ACTIVITIES)
    count_before = len(df_before)
    st.write(f"Count before: {count_before}")

    # 2. Insert dummy data
    # We need a valid person_id. Let's pick one from existing activities or persons if possible.
    # For now, let's try a dummy ID if foreign key constraint allows, or pick top 1 person_id.
    
    # Need to fetch persons to get a valid ID
    from modules.database import fetch_table
    from modules.constants import MAP_PERSONS
    df_p = fetch_table("persons", MAP_PERSONS)
    if df_p.empty:
        st.error("No persons found")
        st.stop()
    
    valid_pid = df_p.iloc[0]['person_id']
    st.write(f"Using Person ID: {valid_pid}")

    new_data = {
        'person_id': valid_pid,
        '記録日': '2025-01-01',
        '活動': 'その他',
        '場所': 'Debug Test',
        '所要時間': 10,
        '交通費・立替金': 0,
        '重要': False,
        '要点': 'Debug Entry'
    }

    insert_data("activities", new_data, MAP_ACTIVITIES)
    st.write("Inserted.")
    
    # insert_data clears cache. 
    # But streamlit script execution model means we usually verify on NEXT run.
    # However, within the same run, if we call fetch_table again, it should use the cache UNLESS it was cleared.
    # Since insert_data calls st.cache_data.clear(), the next fetch_table call SHOULD hit the DB.

    df_after = fetch_table("activities", MAP_ACTIVITIES)
    count_after = len(df_after)
    st.write(f"Count after: {count_after}")

    if count_after > count_before:
        st.success("Test passed: Count increased.")
        # Cleanup
        # We need the ID of the inserted row. 
        # Since we don't get ID back from insert_data (it just executes), we have to find it.
        # It should be the latest one for this person.
        df_after['safe_pid'] = df_after['person_id'].astype(str) # simplified safe id
        # ... logic to delete ...
        st.write("Please delete the test entry manually or implement cleanup.")
    else:
        st.error("Test failed: Count did not increase.")


import streamlit as st
import geopandas as gpd
import pandas as pd

st.set_page_config(page_title="Data Inspector Tool", layout="wide")

st.title("County Data Inspector 🕵️‍♂️")
st.markdown("Upload a zipped shapefile or CSV to reveal its column names and preview the data. Use this to build your `COUNTY_CONFIGS` dictionary.")

# File uploader
uploaded_file = st.file_uploader("Upload County Zip or CSV", type=["zip", "csv"])

if uploaded_file:
    with st.spinner("Reading file..."):
        try:
            # Determine file type and read
            if uploaded_file.name.endswith('.zip'):
                df = gpd.read_file(uploaded_file)
            elif uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            
            st.success("File successfully loaded!")
            
            # 1. Print all column names clearly
            st.subheader("📋 Column Names (Copy these for your dictionary)")
            columns_list = df.columns.tolist()
            st.write(columns_list)
            
            # 2. Show a preview of the data so the user can see what's inside the columns
            st.subheader("👀 Data Preview (First 5 Rows)")
            st.dataframe(df.head())
            
            # 3. Provide a checklist for the user
            st.subheader("✅ Your Mission Checklist")
            st.markdown("""
            Look at the columns and data above. Find the exact matching column name for:
            1. **Unique Parcel ID** (e.g., TAXKEY, PARCEL_ID, PIN)
            2. **House Number** (e.g., HouseNo, HS_NUM)
            3. **House Number Suffix** (e.g., HouseSx, fraction, letter) *(If none, note that)*
            4. **Street Direction** (e.g., Dir, N, S, E, W)
            5. **Street Name** (e.g., Street, STR_NAME)
            6. **Street Type** (e.g., StType, AVE, BLVD)
            7. **Municipality / City** (e.g., Muni, CITY)
            8. **Zip Code** (e.g., Zip_Code, ZIP)
            9. **Unit / Apartment Number** (e.g., Unit, APT)
            10. **Property Status / Class** (e.g., Addr_Statu, CLASS, LAND_USE)
            """)

        except Exception as e:
            st.error(f"Error reading file: {e}")

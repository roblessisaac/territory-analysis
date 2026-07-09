import streamlit as st
import geopandas as gpd
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import fiona
import io
import datetime

# Enable KML support in GeoPandas
fiona.drvsupport.supported_drivers['KML'] = 'rw'
fiona.drvsupport.supported_drivers['LIBKML'] = 'rw'

# --- 1. CONFIGURATION & UI SETUP ---
st.set_page_config(page_title="Territory Audit Engine", layout="wide")

st.title("Congregation Territory Address Analyzer")
st.markdown("Upload your KML map to generate a complete, filtered letter-writing database.")

st.sidebar.header("Step 1: Configuration")
congregation_name = st.sidebar.text_input("Congregation Name (No Spaces)", "ExampleCongregation")
selected_county = st.sidebar.selectbox("Select County Data", ["Milwaukee"]) 
goal_range = st.sidebar.selectbox("Goal # of Addresses Per Territory", 
                                  ["25-50", "50-75", "75-100", "100-125", "125-150", "150-175"])

st.header("Step 2: Upload Territory Map")
uploaded_kml = st.file_uploader("Upload Territory KML File", type=["kml"])

# Parse Goal Range
MIN_GOAL, MAX_GOAL = [int(x) for x in goal_range.split("-")]

# --- 2. DATA LOADING & CACHING ---
@st.cache_data
def load_county_data(county_name):
    if county_name == "Milwaukee":
        file_path = "zip://data/Milwaukee_Datapoints07072026.zip"
        try:
            gdf = gpd.read_file(file_path)
            invalid_statuses = [
                'Undeveloped', 'Parking Lot', 'ROW', 'Park or Recreational Facility',
                'Undeveloped Outlot', 'Sliver or Remnant', 'Non Addressable Assoc with Adj Parcel'
            ]
            gdf = gdf[~gdf['Addr_Statu'].isin(invalid_statuses)]
            return gdf
        except Exception as e:
            st.error(f"Error loading county shapefile. Check that the zip is in the /data/ folder. Error: {e}")
            return None
    return None

# --- 3. EXCEL GENERATION ENGINE ---
# Notice we added kml_gdf here so Tab 5 can see the borders!
def generate_excel_report(joined_gdf, kml_gdf, min_goal, max_goal, cong_name):
    output = io.BytesIO()
    counts_df = joined_gdf.groupby('Territory_Name').size().reset_index(name='Total_Addresses')
    
    def get_category(count):
        if count < min_goal: return "Undersized"
        elif min_goal <= count <= max_goal: return "Ideal"
        else: return "Oversized"
        
    counts_df['Category'] = counts_df['Total_Addresses'].apply(get_category)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- TAB 1: DASHBOARD ---
        total_territories = len(counts_df)
        total_addresses = counts_df['Total_Addresses'].sum()
        avg_addresses = total_addresses / total_territories if total_territories > 0 else 0
        largest_terr = counts_df.loc[counts_df['Total_Addresses'].idxmax()] if total_territories > 0 else None
        smallest_terr = counts_df.loc[counts_df['Total_Addresses'].idxmin()] if total_territories > 0 else None
        ideal_pct = (len(counts_df[counts_df['Category'] == 'Ideal']) / total_territories) * 100 if total_territories > 0 else 0
        
        dashboard_data = [
            [f"Congregation Name: {cong_name}"],
            [f"Analysis generated: {datetime.datetime.now().strftime('%B %Y')}"],
            [""],
            [f"Total Territories: {total_territories}"],
            [f"Total Addresses: {total_addresses}"],
            [f"Average Addresses per Territory: {int(avg_addresses)}"],
            [f"Largest Territory: {largest_terr['Territory_Name']} ({largest_terr['Total_Addresses']} addresses)" if largest_terr is not None else ""],
            [f"Smallest Territory: {smallest_terr['Territory_Name']} ({smallest_terr['Total_Addresses']} addresses)" if smallest_terr is not None else ""],
            [""],
            [f"Goal Range: {min_goal}-{max_goal}"],
            [f"Percentage of Ideal Territories: {ideal_pct:.1f}%"]
        ]
        
        pd.DataFrame(dashboard_data).to_excel(writer, sheet_name="Dashboard", index=False, header=False)
        
        ranges = ["25-50", "50-75", "75-100", "100-125", "125-150", "150-175"]
        distribution = []
        for r in ranges:
            rmin, rmax = [int(x) for x in r.split("-")]
            count = len(counts_df[(counts_df['Total_Addresses'] >= rmin) & (counts_df['Total_Addresses'] <= rmax)])
            cat = "Ideal" if rmin == min_goal else ("Undersized" if rmax <= min_goal else "Oversized")
            distribution.append([cat, r, count])
            
        pd.DataFrame(distribution, columns=["Category", "Range", "Count"]).to_excel(writer, sheet_name="Dashboard", startrow=13, index=False)

        # --- TAB 2: COUNT PER TERRITORY ---
        counts_df_sorted = counts_df.sort_values(by='Territory_Name')
        counts_df_sorted.to_excel(writer, sheet_name="Counts", index=False)
        worksheet = writer.sheets['Counts']
        for row in range(2, len(counts_df_sorted) + 2):
            cell = worksheet[f'C{row}']
            if cell.value == 'Ideal':
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif cell.value == 'Undersized':
                cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            elif cell.value == 'Oversized':
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        # --- TAB 3: ADDRESS LIST ---
        def build_address(row):
            addr = str(row['FullAddr']) if pd.notna(row['FullAddr']) else ""
            unit = f" Apt {str(row['Unit'])}" if pd.notna(row['Unit']) and str(row['Unit']).strip() != "" else ""
            muni = str(row['Muni']) if pd.notna(row['Muni']) else ""
            zip_c = str(row['Zip_Code']) if pd.notna(row['Zip_Code']) else ""
            return f"{addr}{unit}, {muni}, WI {zip_c}".strip(" ,")

        joined_gdf['Mailable_Address'] = joined_gdf.apply(build_address, axis=1)
        joined_gdf['HouseNum_Sort'] = pd.to_numeric(joined_gdf['HouseNo'], errors='coerce').fillna(0)
        address_list_df = joined_gdf.sort_values(by=['Territory_Name', 'Zip_Code', 'Street', 'HouseNum_Sort'])
        
        export_df = address_list_df[['Territory_Name', 'Mailable_Address', 'HouseNo', 'Street', 'Unit', 'Zip_Code']]
        export_df.to_excel(writer, sheet_name="Address List", index=False)
        
        ws3 = writer.sheets['Address List']
        ws3.column_dimensions['C'].hidden = True
        ws3.column_dimensions['D'].hidden = True
        ws3.column_dimensions['E'].hidden = True
        ws3.column_dimensions['F'].hidden = True
        ws3.column_dimensions['B'].width = 40

        # --- TAB 4: APARTMENTS / POTENTIAL LETTER WRITING ---
        apt_groups = joined_gdf.groupby(['Territory_Name', 'FullAddr', 'Zip_Code']).size().reset_index(name='Total_Units')
        apt_groups = apt_groups[apt_groups['Total_Units'] >= 5]
        cat_mapping = counts_df.set_index('Territory_Name')['Category'].to_dict()
        apt_groups['Status'] = apt_groups['Territory_Name'].map(cat_mapping)
        
        def format_complex(row):
            return f"{row['Territory_Name']} - [{row['Status']}] - {row['FullAddr']} - {row['Total_Units']} Units"
            
        apt_groups['Complex_Title'] = apt_groups.apply(format_complex, axis=1)
        apt_export = apt_groups[['Complex_Title', 'Territory_Name', 'FullAddr', 'Total_Units', 'Status']]
        apt_export.to_excel(writer, sheet_name="Apartments", index=False)
        writer.sheets['Apartments'].column_dimensions['A'].width = 60

        # --- TAB 5: BORDER REWRITES ---
        oversized = counts_df[counts_df['Category'] == 'Oversized']['Territory_Name'].tolist()
        undersized = counts_df[counts_df['Category'] == 'Undersized']['Territory_Name'].tolist()
        
        # FIX: Look at the raw kml_gdf for borders, not the joined_gdf!
        terr_geoms = kml_gdf.drop_duplicates('Territory_Name').set_index('Territory_Name')
        suggestions = []
        
        for over_name in oversized:
            over_geom = terr_geoms.loc[over_name, 'geometry_terr']
            over_count = counts_df[counts_df['Territory_Name'] == over_name]['Total_Addresses'].values[0]
            
            for under_name in undersized:
                under_geom = terr_geoms.loc[under_name, 'geometry_terr']
                if over_geom.touches(under_geom) or over_geom.intersects(under_geom):
                    under_count = counts_df[counts_df['Territory_Name'] == under_name]['Total_Addresses'].values[0]
                    rec = f"{over_name} ({over_count} addrs) borders {under_name} ({under_count} addrs). Shift border."
                    suggestions.append([over_name, over_count, under_name, under_count, rec])
                    
        pd.DataFrame(suggestions, columns=["Oversized Territory", "Current Count", "Adjacent Undersized", "Current Count", "Recommendation"]).to_excel(writer, sheet_name="Border Rewrites", index=False)
        writer.sheets['Border Rewrites'].column_dimensions['E'].width = 80

        # Tab colorization
        writer.sheets['Dashboard'].sheet_properties.tabColor = "1E90FF"
        writer.sheets['Counts'].sheet_properties.tabColor = "32CD32"
        writer.sheets['Address List'].sheet_properties.tabColor = "32CD32"
        writer.sheets['Apartments'].sheet_properties.tabColor = "FF8C00"
        writer.sheets['Border Rewrites'].sheet_properties.tabColor = "FF0000"

    output.seek(0)
    return output

# --- 4. EXECUTION FLOW ---
if 'last_uploaded_kml' not in st.session_state:
    st.session_state['last_uploaded_kml'] = None

if uploaded_kml != st.session_state['last_uploaded_kml']:
    if 'excel_data' in st.session_state:
        del st.session_state['excel_data']
    st.session_state['last_uploaded_kml'] = uploaded_kml

if uploaded_kml:
    if st.button("Generate Territory Analysis"):
        with st.spinner(f"Loading Master {selected_county} County Data..."):
            parcel_gdf = load_county_data(selected_county)
            
        if parcel_gdf is not None:
            with st.spinner("Parsing KML Territories & Executing Spatial Join..."):
                try:
                    kml_gdf = gpd.read_file(uploaded_kml, driver="KML")
                    
                    # Auto-repair invalid polygons
                    kml_gdf['geometry'] = kml_gdf['geometry'].make_valid()
                    
                    # Dynamic KML Name parsing (Pandas 3.0 Safe)
                    fallback_names = "Territory_" + kml_gdf.index.to_series().astype(str)
                    
                    if 'Name' in kml_gdf.columns:
                        kml_gdf['Territory_Name'] = kml_gdf['Name'].fillna(fallback_names)
                    elif 'Description' in kml_gdf.columns:
                        kml_gdf['Territory_Name'] = kml_gdf['Description'].fillna(fallback_names)
                    else:
                        kml_gdf['Territory_Name'] = fallback_names
                    
                    bounding_box = kml_gdf.unary_union.envelope
                    parcel_gdf = gpd.clip(parcel_gdf, bounding_box)
                    
                    if parcel_gdf.crs != kml_gdf.crs:
                        parcel_gdf = parcel_gdf.to_crs(kml_gdf.crs)
                    
                    kml_gdf = kml_gdf.rename(columns={'geometry': 'geometry_terr'})
                    kml_gdf = kml_gdf.set_geometry('geometry_terr')
                    
                    joined_gdf = gpd.sjoin(parcel_gdf, kml_gdf, how="inner", predicate="within")
                    joined_gdf = joined_gdf.dropna(subset=['Territory_Name'])
                    
                    with st.spinner("Generating Excel Report..."):
                        # Passed kml_gdf into the function here!
                        excel_file = generate_excel_report(joined_gdf, kml_gdf, MIN_GOAL, MAX_GOAL, congregation_name.replace(" ", ""))
                        filename = f"{congregation_name.replace(' ', '')}_{datetime.datetime.now().strftime('%B%Y')}_TerritoryAnalysis.xlsx"
                        
                        st.session_state['excel_data'] = excel_file.getvalue()
                        st.session_state['excel_filename'] = filename
                        
                        st.success("Analysis Complete!")
                        
                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")

    if 'excel_data' in st.session_state:
        st.info("Analysis results ready for download.")
        st.download_button(
            label="⬇️ Download Excel Analysis",
            data=st.session_state['excel_data'],
            file_name=st.session_state['excel_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
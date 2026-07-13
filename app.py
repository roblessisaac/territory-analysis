import streamlit as st
import geopandas as gpd
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
import fiona
import io
import datetime
import re

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
            # Load ALL data without filtering. We need everything for the spatial join first!
            gdf = gpd.read_file(file_path)
            return gdf
        except Exception as e:
            st.error(f"Error loading county shapefile. Check that the zip is in the /data/ folder. Error: {e}")
            return None
    return None

# --- NATURAL SORTING HELPER ---
def natural_keys(text):
    """Splits text into letters and numbers so Python sorts '2' before '10'."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

# --- ADDRESS BUILDER ---
def build_addresses(row):
    # Cleanly extract parts and safely ignore "nan" strings or nulls
    house = str(row['HouseNo']).replace('.0', '').strip() if pd.notna(row['HouseNo']) and str(row['HouseNo']).lower() != "nan" else ""
    house_sx = str(row['HouseSx']).strip() if pd.notna(row['HouseSx']) and str(row['HouseSx']).lower() != "nan" else ""
    direction = str(row['Dir']).strip() if pd.notna(row['Dir']) and str(row['Dir']).lower() != "nan" else ""
    street = str(row['Street']).strip() if pd.notna(row['Street']) and str(row['Street']).lower() != "nan" else ""
    st_type = str(row['StType']).strip() if pd.notna(row['StType']) and str(row['StType']).lower() != "nan" else ""
    muni = str(row['Muni']).strip() if pd.notna(row['Muni']) and str(row['Muni']).lower() != "nan" else ""
    zip_c = str(row['Zip_Code']).strip() if pd.notna(row['Zip_Code']) and str(row['Zip_Code']).lower() != "nan" else ""
    
    # Unit logic: only add "Apt " if unit actually exists
    unit_val = str(row['Unit']).strip() if pd.notna(row['Unit']) and str(row['Unit']).lower() != "nan" else ""
    unit_str = f" Apt {unit_val}" if unit_val else ""

    # Connect HouseNo and HouseSx without a space (e.g. 1452B)
    full_house_num = f"{house}{house_sx}"
    
    # Combine street direction, name, and type
    street_parts = [direction, street, st_type]
    full_street = " ".join([p for p in street_parts if p])

    # Base Address (Strictly NO unit)
    base_addr_line = f"{full_house_num} {full_street}".strip()
    base_addr = f"{base_addr_line}, {muni}, WI {zip_c}".replace(" ,", ",").strip(" ,")

    # Mailable Address (Includes unit)
    mailable_addr_line = f"{base_addr_line}{unit_str}".strip()
    mailable_addr = f"{mailable_addr_line}, {muni}, WI {zip_c}".replace(" ,", ",").strip(" ,")

    return pd.Series([base_addr, mailable_addr])


# --- 3. EXCEL GENERATION ENGINE ---
def generate_excel_report(joined_gdf, kml_gdf, min_goal, max_goal, cong_name):
    output = io.BytesIO()
    
    # Slice Zip Codes down to 5 digits
    joined_gdf['Zip_Code'] = joined_gdf['Zip_Code'].astype(str).str[:5]
    
    # Apply Address Builder to the entire dataset
    joined_gdf[['Base_Address', 'Mailable_Address']] = joined_gdf.apply(build_addresses, axis=1)
    
    # Order of Operations: Split the data AFTER the spatial join
    invalid_statuses = [
        'Undeveloped', 'Parking Lot', 'ROW', 'Park or Recreational Facility',
        'Undeveloped Outlot', 'Sliver or Remnant', 'Non Addressable Assoc with Adj Parcel'
    ]
    excluded_gdf = joined_gdf[joined_gdf['Addr_Statu'].isin(invalid_statuses)].copy()
    valid_gdf = joined_gdf[~joined_gdf['Addr_Statu'].isin(invalid_statuses)].copy()

    # Apply Natural Sorting globally to the Territory_Name column
    unique_territories = valid_gdf['Territory_Name'].unique().tolist()
    unique_territories.sort(key=natural_keys)
    valid_gdf['Territory_Name'] = pd.Categorical(valid_gdf['Territory_Name'], categories=unique_territories, ordered=True)
    
    if not excluded_gdf.empty:
        excluded_unique = excluded_gdf['Territory_Name'].unique().tolist()
        excluded_unique.sort(key=natural_keys)
        excluded_gdf['Territory_Name'] = pd.Categorical(excluded_gdf['Territory_Name'], categories=excluded_unique, ordered=True)

    # Base counts off the VALID data only
    counts_df = valid_gdf.groupby('Territory_Name', observed=True).size().reset_index(name='Total_Addresses')
    counts_df = counts_df[counts_df['Total_Addresses'] > 0]
    
    def get_category(count):
        if count < min_goal: return "Undersized"
        elif min_goal <= count <= max_goal: return "Ideal"
        else: return "Oversized"
        
    counts_df['Category'] = counts_df['Total_Addresses'].apply(get_category)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- TAB 1: DASHBOARD ---
        total_territories = len(counts_df)
        total_addresses = counts_df['Total_Addresses'].sum()
        largest_terr = counts_df.loc[counts_df['Total_Addresses'].idxmax()] if total_territories > 0 else None
        smallest_terr = counts_df.loc[counts_df['Total_Addresses'].idxmin()] if total_territories > 0 else None
        ideal_pct = (len(counts_df[counts_df['Category'] == 'Ideal']) / total_territories) * 100 if total_territories > 0 else 0
        
        largest_name = largest_terr['Territory_Name'] if largest_terr is not None else ""
        largest_count = largest_terr['Total_Addresses'] if largest_terr is not None else 0
        smallest_name = smallest_terr['Territory_Name'] if smallest_terr is not None else ""
        smallest_count = smallest_terr['Total_Addresses'] if smallest_terr is not None else 0

        # Build Dashboard Rows 1-11
        dashboard_top = [
            [f"Territory Analysis: {cong_name}"],
            [f"Generated {datetime.datetime.now().strftime('%B %Y')} by Territory Analysis Engine."],
            [""],
            [f"Total Territories: {total_territories}"],
            [f"Total Valid Addresses: {total_addresses}"],
            [f"Excluded Addresses (See Tab 6): {len(excluded_gdf)}"],
            [f"The largest territory has {largest_count} addresses in it ({largest_name})."],
            [f"The smallest territory has {smallest_count} addresses in it ({smallest_name})."],
            [""],
            [f"Goal Range: {min_goal}-{max_goal}"],
            [f"About {ideal_pct:.1f}% of territories fall within this range."]
        ]
        pd.DataFrame(dashboard_top).to_excel(writer, sheet_name="Dashboard", index=False, header=False)
        
        # Build Dashboard Data Grid (Rows 12-18)
        ranges = ["25-50", "50-75", "75-100", "100-125", "125-150", "150-175"]
        distribution = []
        for r in ranges:
            rmin, rmax = [int(x) for x in r.split("-")]
            count = len(counts_df[(counts_df['Total_Addresses'] >= rmin) & (counts_df['Total_Addresses'] <= rmax)])
            cat = "Ideal" if rmin == min_goal else ("Undersized" if rmax <= min_goal else "Oversized")
            distribution.append([cat, r, count])
            
        pd.DataFrame(distribution, columns=["Category", "Range", "Count"]).to_excel(writer, sheet_name="Dashboard", startrow=11, index=False)

        # Build Dashboard Bottom Narratives (Rows 19-22)
        dashboard_bottom = [
            [""],
            ["As a part of this analysis, every address point within your territory was collected & identified. If you’d like to incorporate these specific addresses into your territory management system, please see http://www.territoryanalysis.com/ to see if your system is supported."],
            ["Note: It is suggested that you export this into a program you can easily edit to expand cells to read easier, create filters to see specific data, and customize to make them more legible."],
            ["If you have questions on what any data in the spreadsheet means, please see http://www.territoryanalysis.com/ for explanation"]
        ]
        pd.DataFrame(dashboard_bottom).to_excel(writer, sheet_name="Dashboard", startrow=18, index=False, header=False)

        # Apply OpenPyXL formatting to Dashboard
        ws1 = writer.sheets['Dashboard']
        ws1.column_dimensions['A'].width = 110
        
        # Formatting Row 1 (Title)
        ws1['A1'].font = Font(size=20, bold=True)
        
        # Formatting Row 2 (Hyperlink)
        ws1['A2'].hyperlink = "http://www.territoryanalysis.com/"
        ws1['A2'].font = Font(color="0563C1", underline="single")
        
        # Text Wrapping ONLY the bottom rows (Rows 20, 21, 22 in Excel)
        for r in [20, 21, 22]:
            ws1.cell(row=r, column=1).alignment = Alignment(wrap_text=True)

        # --- TAB 2: COUNT PER TERRITORY ---
        counts_df_sorted = counts_df.sort_values(by='Territory_Name')
        counts_df_sorted.to_excel(writer, sheet_name="Counts", index=False)
        worksheet2 = writer.sheets['Counts']
        for row in range(2, len(counts_df_sorted) + 2):
            cell = worksheet2[f'C{row}']
            if cell.value == 'Ideal':
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif cell.value == 'Undersized':
                cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            elif cell.value == 'Oversized':
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        # --- TAB 3: ADDRESS LIST ---
        valid_gdf['HouseNum_Sort'] = pd.to_numeric(valid_gdf['HouseNo'], errors='coerce').fillna(0)
        address_list_df = valid_gdf.sort_values(by=['Territory_Name', 'Street', 'HouseNum_Sort', 'Unit'])
        
        export_df = address_list_df[['Territory_Name', 'Mailable_Address', 'HouseNo', 'Street', 'Unit', 'Zip_Code']]
        export_df.to_excel(writer, sheet_name="Address List", index=False)
        
        ws3 = writer.sheets['Address List']
        ws3.column_dimensions['C'].hidden = True
        ws3.column_dimensions['D'].hidden = True
        ws3.column_dimensions['E'].hidden = True
        ws3.column_dimensions['F'].hidden = True
        ws3.column_dimensions['B'].width = 45

        # --- TAB 4: APARTMENTS / POTENTIAL LETTER WRITING ---
        apt_groups = valid_gdf.groupby(['Territory_Name', 'Base_Address'], observed=True).size().reset_index(name='Total Units')
        apt_groups = apt_groups[apt_groups['Total Units'] >= 5]
        
        if not counts_df.empty:
            cat_mapping = counts_df.set_index('Territory_Name')['Category'].to_dict()
            apt_groups['Status'] = apt_groups['Territory_Name'].map(cat_mapping)
        else:
            apt_groups['Status'] = "Unknown"
        
        def format_terr_name(row):
            return f"{row['Territory_Name']} [{row['Status']}]"
            
        if not apt_groups.empty:
            apt_groups['Territory Name'] = apt_groups.apply(format_terr_name, axis=1)
            apt_groups.rename(columns={'Base_Address': 'Base Address'}, inplace=True)
            apt_export = apt_groups[['Territory Name', 'Base Address', 'Total Units']]
        else:
            apt_export = pd.DataFrame(columns=['Territory Name', 'Base Address', 'Total Units'])
            
        apt_export.to_excel(writer, sheet_name="Apartments", index=False)
        writer.sheets['Apartments'].column_dimensions['A'].width = 30
        writer.sheets['Apartments'].column_dimensions['B'].width = 40

        # --- TAB 5: BORDER REWRITES ---
        oversized = counts_df[counts_df['Category'] == 'Oversized']['Territory_Name'].tolist() if not counts_df.empty else []
        undersized = counts_df[counts_df['Category'] == 'Undersized']['Territory_Name'].tolist() if not counts_df.empty else []
        
        terr_geoms = kml_gdf.drop_duplicates('Territory_Name').set_index('Territory_Name')
        suggestions = []
        
        for over_name in oversized:
            if over_name in terr_geoms.index:
                over_geom = terr_geoms.loc[over_name, 'geometry_terr']
                over_count = counts_df[counts_df['Territory_Name'] == over_name]['Total_Addresses'].values[0]
                
                for under_name in undersized:
                    if under_name in terr_geoms.index:
                        under_geom = terr_geoms.loc[under_name, 'geometry_terr']
                        if over_geom.touches(under_geom) or over_geom.intersects(under_geom):
                            under_count = counts_df[counts_df['Territory_Name'] == under_name]['Total_Addresses'].values[0]
                            diff = abs(over_count - under_count)
                            rec = f"That is a {diff} address difference. Shrink {over_name} & Expand {under_name}."
                            suggestions.append([over_name, over_count, under_name, under_count, rec])
                        
        pd.DataFrame(suggestions, columns=["Oversized Territory", "Current Count", "Adjacent Undersized", "Current Count", "Recommendation"]).to_excel(writer, sheet_name="Border Rewrites", index=False)
        writer.sheets['Border Rewrites'].column_dimensions['E'].width = 85

        # --- TAB 6: EXCLUDED AUDIT ---
        if not excluded_gdf.empty:
            excluded_gdf['HouseNum_Sort'] = pd.to_numeric(excluded_gdf['HouseNo'], errors='coerce').fillna(0)
            excluded_list_df = excluded_gdf.sort_values(by=['Territory_Name', 'Street', 'HouseNum_Sort', 'Unit'])
            
            export_ex_df = excluded_list_df[['Territory_Name', 'Mailable_Address', 'Addr_Statu', 'HouseNo', 'Street', 'Unit', 'Zip_Code']]
            export_ex_df.to_excel(writer, sheet_name="Excluded Audit", index=False)
            
            ws6 = writer.sheets['Excluded Audit']
            ws6.column_dimensions['D'].hidden = True
            ws6.column_dimensions['E'].hidden = True
            ws6.column_dimensions['F'].hidden = True
            ws6.column_dimensions['G'].hidden = True
            ws6.column_dimensions['B'].width = 45
            ws6.column_dimensions['C'].width = 30
        else:
            pd.DataFrame(columns=["Notice"]).to_excel(writer, sheet_name="Excluded Audit", index=False)
            writer.sheets['Excluded Audit'].cell(row=2, column=1, value="No addresses were excluded in this map area.")

        # --- EXCEL UX POLISH (FREEZE PANES & BOLD HEADERS) ---
        tabs_to_format = ["Counts", "Address List", "Apartments", "Border Rewrites", "Excluded Audit"]
        for tab_name in tabs_to_format:
            ws = writer.sheets[tab_name]
            ws.freeze_panes = 'A2' # Freezes Row 1
            for cell in ws[1]:     # Bolds every cell in Row 1
                cell.font = Font(bold=True)

        # Tab colorization
        writer.sheets['Dashboard'].sheet_properties.tabColor = "1E90FF"
        writer.sheets['Counts'].sheet_properties.tabColor = "32CD32"
        writer.sheets['Address List'].sheet_properties.tabColor = "32CD32"
        writer.sheets['Apartments'].sheet_properties.tabColor = "FF8C00"
        writer.sheets['Border Rewrites'].sheet_properties.tabColor = "FF0000"
        writer.sheets['Excluded Audit'].sheet_properties.tabColor = "808080" 

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
                    kml_gdf['geometry'] = kml_gdf['geometry'].make_valid()
                    
                    fallback_names = "Territory_" + kml_gdf.index.to_series().astype(str)
                    if 'Name' in kml_gdf.columns:
                        kml_gdf['Territory_Name'] = kml_gdf['Name'].fillna(fallback_names)
                    elif 'Description' in kml_gdf.columns:
                        kml_gdf['Territory_Name'] = kml_gdf['Description'].fillna(fallback_names)
                    else:
                        kml_gdf['Territory_Name'] = fallback_names
                    
                    if parcel_gdf.crs != kml_gdf.crs:
                        parcel_gdf = parcel_gdf.to_crs(kml_gdf.crs)
                        
                    bounding_box = kml_gdf.unary_union.envelope
                    parcel_gdf = gpd.clip(parcel_gdf, bounding_box)
                    
                    kml_gdf = kml_gdf.rename(columns={'geometry': 'geometry_terr'})
                    kml_gdf = kml_gdf.set_geometry('geometry_terr')
                    
                    joined_gdf = gpd.sjoin(parcel_gdf, kml_gdf, how="inner", predicate="within")
                    joined_gdf = joined_gdf.dropna(subset=['Territory_Name'])
                    
                    with st.spinner("Generating Excel Report..."):
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
            label="⬇️ Download Excel Analysis File",
            data=st.session_state['excel_data'],
            file_name=st.session_state['excel_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

import streamlit as st
import geopandas as gpd
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.table import Table, TableStyleInfo
import fiona
import io
import datetime
import re
from fractions import Fraction

# Enable KML support in GeoPandas
fiona.drvsupport.supported_drivers['KML'] = 'rw'
fiona.drvsupport.supported_drivers['LIBKML'] = 'rw'

# --- 1. CONFIGURATION & UI SETUP ---
st.set_page_config(page_title="Territory Audit Engine", layout="wide")

st.title("Congregation Territory Analysis Engine")
st.markdown("Upload your territories KML map to generate a complete, filtered address database & analysis.")

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
            return gdf
        except Exception as e:
            st.error(f"Error loading county shapefile. Check that the zip is in the /data/ folder. Error: {e}")
            return None
    return None

# --- NATURAL SORTING HELPER ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

# --- ADDRESS BUILDER + NORMALIZATION HELPERS ---
def clean_field(value):
    if pd.isna(value): return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "<na>"}: return ""
    return text

def normalize_house_number(value):
    text = clean_field(value)
    if not text: return ""
    if re.fullmatch(r"[+-]?\d+\.0+", text): return text.split(".", 1)[0]
    return re.sub(r"\s+", " ", text)

def normalize_zip_code(value):
    text = clean_field(value)
    if not text: return ""
    text = re.sub(r"\.0+$", "", text)
    digits = re.sub(r"\D", "", text)
    
    if not digits: return ""
    if len(digits) == 4: digits = digits.zfill(5)
    elif len(digits) == 8: digits = digits.zfill(9)

    if len(digits) == 5: return digits
    if len(digits) == 9: return f"{digits[:5]}-{digits[5:]}"
    if len(digits) > 9:
        digits = digits[:9]
        return f"{digits[:5]}-{digits[5:]}"
    return digits

def normalize_unit(value):
    text = clean_field(value)
    if not text: return ""
    if re.fullmatch(r"\d+\.0+", text): text = text.split(".", 1)[0]
    text = re.sub(r"\s+", " ", text).strip()

    descriptive_pattern = re.compile(
        r"^(?:"
        r"apt(?:artment)?|"
        r"unit|"
        r"ste|suite|"
        r"upper|lower|"
        r"bsmt|basement|"
        r"rear|front|"
        r"floor|fl|"
        r"building|bldg|"
        r"room|rm"
        r")\b",
        flags=re.IGNORECASE,
    )

    if descriptive_pattern.search(text): return text
    return f"Apt {text}"

def house_number_sort_parts(value):
    text = normalize_house_number(value).upper()
    if not text: return pd.Series([float("inf"), 9, ""])
    compact = re.sub(r"\s+", " ", text).strip()

    mixed_fraction = re.match(r"^(\d+)\s+(\d+)\s*/\s*(\d+)(.*)$", compact)
    if mixed_fraction:
        whole, numerator, denominator, suffix = mixed_fraction.groups()
        try: numeric_value = int(whole) + float(Fraction(int(numerator), int(denominator)))
        except (ValueError, ZeroDivisionError): numeric_value = float(int(whole))
        return pd.Series([numeric_value, 1, suffix.strip()])

    simple_fraction = re.match(r"^(\d+)\s*/\s*(\d+)(.*)$", compact)
    if simple_fraction:
        numerator, denominator, suffix = simple_fraction.groups()
        try: numeric_value = float(Fraction(int(numerator), int(denominator)))
        except (ValueError, ZeroDivisionError): numeric_value = float("inf")
        return pd.Series([numeric_value, 1, suffix.strip()])

    numeric_prefix = re.match(r"^(\d+(?:\.\d+)?)(.*)$", compact)
    if numeric_prefix:
        number, suffix = numeric_prefix.groups()
        suffix = suffix.strip()
        suffix_rank = 0 if not suffix else 2
        return pd.Series([float(number), suffix_rank, suffix])

    return pd.Series([float("inf"), 8, compact])

def build_addresses(row):
    house = normalize_house_number(row.get("HouseNo"))
    house_sx = clean_field(row.get("HouseSx"))
    direction = clean_field(row.get("Dir"))
    street = clean_field(row.get("Street"))
    st_type = clean_field(row.get("StType"))
    muni = clean_field(row.get("Muni"))
    zip_c = normalize_zip_code(row.get("Zip_Code"))
    unit_str = normalize_unit(row.get("Unit"))

    full_house_num = f"{house}{house_sx}".strip()
    street_parts = [direction, street, st_type]
    full_street = " ".join(part for part in street_parts if part)
    base_addr_line = " ".join(part for part in [full_house_num, full_street] if part)

    locality_parts = [muni, "WI"]
    locality = ", ".join(part for part in locality_parts if part)
    if zip_c: locality = f"{locality} {zip_c}".strip()

    base_addr = ", ".join(part for part in [base_addr_line, locality] if part)
    mailable_addr_line = " ".join(part for part in [base_addr_line, unit_str] if part)
    mailable_addr = ", ".join(part for part in [mailable_addr_line, locality] if part)

    return pd.Series([base_addr, mailable_addr], index=["Base_Address", "Mailable_Address"])

def evaluate_data_quality(row):
    issues = []
    house = normalize_house_number(row.get("HouseNo"))
    street = clean_field(row.get("Street"))
    municipality = clean_field(row.get("Muni"))
    zip_code = normalize_zip_code(row.get("Zip_Code"))
    base_address = clean_field(row.get("Base_Address"))
    mailable_address = clean_field(row.get("Mailable_Address"))

    if not street: issues.append("Missing Street")
    if not municipality: issues.append("Missing Municipality")
    if not zip_code: issues.append("Missing ZIP")
    if house and re.fullmatch(r"[+-]?0+(?:\.0+)?", house): issues.append("Zero House Number")

    zip_is_valid = not zip_code or bool(re.fullmatch(r"\d{5}(?:-\d{4})?", zip_code))
    address_is_malformed = (
        not house or not base_address or not mailable_address
        or ",," in base_address or ",," in mailable_address or not zip_is_valid
    )

    if address_is_malformed: issues.append("Malformed Address")
    return " | ".join(issues)

# --- 3. EXCEL GENERATION ENGINE ---
def generate_excel_report(joined_gdf, kml_gdf, min_goal, max_goal, cong_name):
    output = io.BytesIO()
    run_timestamp = datetime.datetime.now()

    joined_gdf["Zip_Code"] = joined_gdf["Zip_Code"].map(normalize_zip_code)
    joined_gdf[["Base_Address", "Mailable_Address"]] = joined_gdf.apply(build_addresses, axis=1)
    joined_gdf["Data_Quality_Flag"] = joined_gdf.apply(evaluate_data_quality, axis=1)
    
    flagged_record_count = joined_gdf["Data_Quality_Flag"].ne("").sum()
    
    invalid_statuses = [
        'Undeveloped', 'Parking Lot', 'ROW', 'Park or Recreational Facility',
        'Undeveloped Outlot', 'Sliver or Remnant', 'Non Addressable Assoc with Adj Parcel'
    ]
    excluded_gdf = joined_gdf[joined_gdf['Addr_Statu'].isin(invalid_statuses)].copy()
    valid_gdf = joined_gdf[~joined_gdf['Addr_Statu'].isin(invalid_statuses)].copy()

    unique_territories = valid_gdf['Territory_Name'].unique().tolist()
    unique_territories.sort(key=natural_keys)
    valid_gdf['Territory_Name'] = pd.Categorical(valid_gdf['Territory_Name'], categories=unique_territories, ordered=True)
    
    if not excluded_gdf.empty:
        excluded_unique = excluded_gdf['Territory_Name'].unique().tolist()
        excluded_unique.sort(key=natural_keys)
        excluded_gdf['Territory_Name'] = pd.Categorical(excluded_gdf['Territory_Name'], categories=excluded_unique, ordered=True)

    counts_df = valid_gdf.groupby('Territory_Name', observed=True).size().reset_index(name='Total_Addresses')
    counts_df = counts_df[counts_df['Total_Addresses'] > 0].copy()
    
    def get_category(count):
        if count < min_goal: return "Undersized"
        elif min_goal <= count <= max_goal: return "Ideal"
        else: return "Oversized"
        
    counts_df['Category'] = counts_df['Total_Addresses'].apply(get_category)

    # NWS Explicit Extraction
    valid_gdf[['NWS_Category', 'NWS_Number']] = valid_gdf['Territory_Name'].str.extract(r'^([A-Za-z]+)[-\s]+(.*)$')
    valid_gdf['NWS_Category'] = valid_gdf['NWS_Category'].fillna("UNK")
    valid_gdf['NWS_Number'] = valid_gdf['NWS_Number'].fillna("0")
    
    if not excluded_gdf.empty:
        excluded_gdf[['NWS_Category', 'NWS_Number']] = excluded_gdf['Territory_Name'].str.extract(r'^([A-Za-z]+)[-\s]+(.*)$')
        excluded_gdf['NWS_Category'] = excluded_gdf['NWS_Category'].fillna("UNK")
        excluded_gdf['NWS_Number'] = excluded_gdf['NWS_Number'].fillna("0")

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

        dashboard_top = [
            [f"Territory Analysis: {cong_name}"],
            [f"Generated {run_timestamp.strftime('%B %Y')} by Territory Analysis Engine."],
            [""],
            [f"Total Territories: {total_territories}"],
            [f"Total Valid Addresses: {total_addresses}"],
            [f"Excluded Addresses (See Tab 6): {len(excluded_gdf)}"],
            [f"The largest territory has {largest_count} addresses in it ({largest_name})."],
            [f"The smallest territory has {smallest_count} addresses in it ({smallest_name})."],
            [""],
            [f"Goal Range: {min_goal}-{max_goal}"],
            [""] 
        ]
        pd.DataFrame(dashboard_top).to_excel(writer, sheet_name="Dashboard", index=False, header=False)
        
        distribution = []
        if min_goal > 0:
            under_count = len(counts_df[counts_df['Total_Addresses'] < min_goal])
            distribution.append(["Undersized", f"Under {min_goal}", under_count])
            
        ideal_count = len(counts_df[counts_df['Category'] == 'Ideal'])
        distribution.append(["Ideal", f"{min_goal}-{max_goal}", ideal_count])
        
        current_floor = max_goal + 1
        max_address_count = counts_df['Total_Addresses'].max() if total_territories > 0 else current_floor
        
        while current_floor <= max_address_count:
            current_ceil = current_floor + 24
            chunk_count = len(counts_df[(counts_df['Total_Addresses'] >= current_floor) & (counts_df['Total_Addresses'] <= current_ceil)])
            if chunk_count > 0 or current_floor < 175:
                distribution.append(["Oversized", f"{current_floor}-{current_ceil}", chunk_count])
            current_floor += 25
            
        if current_floor < 175 and max_address_count < 175:
            distribution.append(["Oversized", f"{current_floor} and over", 0])
        elif max_address_count >= current_floor:
            over_count = len(counts_df[counts_df['Total_Addresses'] >= current_floor])
            distribution.append(["Oversized", f"{current_floor} and over", over_count])

        pd.DataFrame(distribution, columns=["Category", "Range", "Count"]).to_excel(writer, sheet_name="Dashboard", startrow=11, index=False)

        ws1 = writer.sheets['Dashboard']
        ws1.column_dimensions['A'].width = 18
        ws1.column_dimensions['B'].width = 25
        ws1.column_dimensions['C'].width = 12
        
        ws1['A1'].font = Font(size=20, bold=True)
        ws1['A2'].hyperlink = "http://www.territoryanalysis.com/"
        ws1['A2'].font = Font(color="0563C1", underline="single")
        
        bold_inline = InlineFont(b=True)
        ws1['A11'].value = CellRichText([
            "About ",
            TextBlock(bold_inline, f"{ideal_pct:.1f}%"),
            " of territories fall within this range."
        ])

        header_fill = PatternFill(start_color="C7CDDB", end_color="C7CDDB", fill_type="solid")
        for col in range(1, 4):
            ws1.cell(row=12, column=col).fill = header_fill
            ws1.cell(row=12, column=col).font = Font(bold=True)

        dist_end_row = 12 + len(distribution)
        for r in range(13, dist_end_row + 1):
            if ws1.cell(row=r, column=1).value == "Ideal":
                for col in range(1, 4):
                    ws1.cell(row=r, column=col).font = Font(bold=True)

        instruct_start = max(20, dist_end_row + 2)
        instructions = [
            CellRichText(["As a part of this analysis, every ", TextBlock(bold_inline, "address point"), " within your territory was collected & identified."]),
            "These addresses, with a little reformatting, can be added to NWS or other programs (Please see http://www.territoryanalysis.com/ to see if your system is supported.)",
            "It's suggested to export this file into a program you can easily edit, like excel or google sheets.",
            "That will allow you to expand cells to read easier, create custom filters to see specific data, and customize the sheet to make it more legible.",
            "",
            CellRichText(["The ", TextBlock(bold_inline, "DASHBOARD"), " tab displays basic statistics about the territory that was analyzed"]),
            CellRichText(["The ", TextBlock(bold_inline, "COUNTS"), " tab organizes territories by size. This is done by 'counting' workable addresses, not geographical size."]),
            CellRichText(["The ", TextBlock(bold_inline, "ADDRESS LIST"), " tab displays every workable address in your territory."]),
            CellRichText(["The ", TextBlock(bold_inline, "APARTMENTS"), " tab displays every multifamily above 5 units in your territory. Large units can be explanations for inflated door-to-door territories."]),
            CellRichText(["The ", TextBlock(bold_inline, "BORDER REWRITES"), " tab displays borders within your territory that may benefit from being redrawn. The intent is to shrink oversized territories adjacent to undersized territories. These are just suggestions."]),
            CellRichText(["The ", TextBlock(bold_inline, "EXCLUDED AUDIT"), " tab displays addresses that are NOT counted towards your territory. These are usually addresses of highways, vacant lots, parks, etc. This is included for confidence."])
        ]
        
        for i, text in enumerate(instructions):
            cell = ws1.cell(row=instruct_start + i, column=1)
            cell.value = text
            if "http" in str(text):
                cell.hyperlink = "http://www.territoryanalysis.com/"
                cell.font = Font(color="0563C1", underline="single")

        tech_start = instruct_start + len(instructions) + 2
        ws1.cell(row=tech_start, column=1, value="Technical: Run Information").font = Font(bold=True, size=12)
        ws1.cell(row=tech_start, column=1).fill = header_fill
        ws1.cell(row=tech_start, column=2).fill = header_fill

        tech_info = [
            ("Run Timestamp", run_timestamp.strftime("%Y-%m-%d %I:%M:%S %p")),
            ("Goal Range Setting", f"{min_goal}-{max_goal} addresses"),
            ("Total Records Loaded", len(joined_gdf)),
            ("Valid Addresses Assigned", len(valid_gdf)),
            ("Excluded Address Count", len(excluded_gdf)),
            ("Records Flagged with Warnings", flagged_record_count)
        ]
        
        for i, (label, val) in enumerate(tech_info, start=1):
            ws1.cell(row=tech_start + i, column=1, value=label).font = Font(bold=True)
            ws1.cell(row=tech_start + i, column=2, value=val)

        def add_excel_table(worksheet, df, table_name):
            if df.empty: return
            max_row, max_col = df.shape
            col_letter = openpyxl.utils.get_column_letter(max_col)
            table_ref = f"A1:{col_letter}{max_row + 1}"
            tab = Table(displayName=table_name, ref=table_ref)
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            worksheet.add_table(tab)

        # --- TAB 2: COUNT PER TERRITORY ---
        counts_df_sorted = counts_df.sort_values(by='Territory_Name').rename(columns={
            'Territory_Name': 'Territory Name', 
            'Total_Addresses': '# of Addresses'
        })
        counts_df_sorted.to_excel(writer, sheet_name="Counts", index=False)
        ws2 = writer.sheets['Counts']
        ws2.column_dimensions['A'].width = 18
        ws2.column_dimensions['B'].width = 15
        ws2.column_dimensions['C'].width = 15
        add_excel_table(ws2, counts_df_sorted, "CountsTable")
        
        for row in range(2, len(counts_df_sorted) + 2):
            ws2[f'B{row}'].alignment = Alignment(horizontal='center')
            cell = ws2[f'C{row}']
            if cell.value == 'Ideal': cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif cell.value == 'Undersized': cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            elif cell.value == 'Oversized': cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')

        # --- TAB 3: ADDRESS LIST ---
        valid_gdf['Latitude'] = valid_gdf.geometry.y
        valid_gdf['Longitude'] = valid_gdf.geometry.x
        valid_gdf[["HouseNum_Sort", "HouseNum_Suffix_Rank", "HouseNum_Text_Sort"]] = valid_gdf["HouseNo"].apply(house_number_sort_parts)
        valid_gdf["Unit_Sort"] = valid_gdf["Unit"].map(clean_field).str.upper()

        address_list_df = valid_gdf.sort_values(by=["Territory_Name", "Street", "HouseNum_Sort", "HouseNum_Suffix_Rank", "HouseNum_Text_Sort", "Unit_Sort"], kind="stable")
        
        export_df = address_list_df[['Territory_Name', 'NWS_Category', 'NWS_Number', 'Mailable_Address', 'Data_Quality_Flag', 'HouseNo', 'HouseSx', 'Street', 'Unit', 'Muni', 'Zip_Code', 'Latitude', 'Longitude']].rename(columns={
            'Territory_Name': 'Territory Name', 
            'Mailable_Address': 'Mailable Address',
            'Data_Quality_Flag': 'Data Quality Flag'
        })
        export_df.to_excel(writer, sheet_name="Address List", index=False)
        
        ws3 = writer.sheets['Address List']
        add_excel_table(ws3, export_df, "AddressListTable")
        
        for col_letter in ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
            ws3.column_dimensions[col_letter].hidden = True
            
        ws3.column_dimensions['A'].width = 18
        ws3.column_dimensions['B'].width = 15
        ws3.column_dimensions['C'].width = 15
        ws3.column_dimensions['D'].width = 55
        ws3.column_dimensions['E'].width = 35

        # --- TAB 4: APARTMENTS / POTENTIAL LETTER WRITING ---
        apartment_source = valid_gdf[["Territory_Name", "Base_Address", "Unit"]].copy()
        apartment_source["_Unit_Normalized"] = apartment_source["Unit"].map(clean_field).str.upper().str.replace(r"\s+", " ", regex=True).str.strip()
        apartment_source = apartment_source[apartment_source["_Unit_Normalized"].ne("") & apartment_source["Base_Address"].map(clean_field).ne("")].copy()

        apt_groups = apartment_source.groupby(["Territory_Name", "Base_Address"], observed=True)["_Unit_Normalized"].nunique().reset_index(name="Total Units")
        apt_groups = apt_groups[apt_groups["Total Units"] >= 5].copy()
        
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
        ws4 = writer.sheets['Apartments']
        add_excel_table(ws4, apt_export, "ApartmentsTable")
        
        ws4.column_dimensions['A'].width = 30
        ws4.column_dimensions['B'].width = 40
        ws4.column_dimensions['C'].width = 15
        for row in range(2, len(apt_export) + 2): ws4[f'C{row}'].alignment = Alignment(horizontal='center')

        # --- TAB 5: BORDER REWRITES ---
        oversized = counts_df.loc[counts_df["Category"].eq("Oversized"), "Territory_Name"].tolist() if not counts_df.empty else []
        undersized = counts_df.loc[counts_df["Category"].eq("Undersized"), "Territory_Name"].tolist() if not counts_df.empty else []

        terr_geoms = kml_gdf[["Territory_Name", "geometry_terr"]].dropna(subset=["Territory_Name", "geometry_terr"]).set_geometry("geometry_terr").dissolve(by="Territory_Name")
        terr_geoms["geometry_terr"] = terr_geoms.geometry.make_valid()
        terr_geoms = terr_geoms[terr_geoms.geometry.notna() & ~terr_geoms.geometry.is_empty].copy()

        count_lookup = counts_df.drop_duplicates("Territory_Name").set_index("Territory_Name")["Total_Addresses"].to_dict()
        undersized_set = set(undersized)
        territory_sindex = terr_geoms.sindex
        suggestions = []

        for over_name in oversized:
            if over_name not in terr_geoms.index: continue
            over_geom = terr_geoms.at[over_name, "geometry_terr"]
            over_count = count_lookup.get(over_name, 0)
            candidate_positions = territory_sindex.query(over_geom, predicate="intersects")

            for candidate_position in candidate_positions:
                under_name = terr_geoms.index[candidate_position]
                if under_name == over_name or under_name not in undersized_set: continue
                under_geom = terr_geoms.iloc[candidate_position].geometry_terr
                if not over_geom.touches(under_geom): continue
                shared_boundary = over_geom.boundary.intersection(under_geom.boundary)
                if shared_boundary.is_empty or shared_boundary.length <= 0: continue
                under_count = count_lookup.get(under_name, 0)
                suggestions.append([over_name, over_count, under_name, under_count, ""])
                        
        sugg_df = pd.DataFrame(suggestions, columns=["Too Large", "Count", "Too Small", "Count ", "Recommendation"])
        sugg_df.to_excel(writer, sheet_name="Border Rewrites", index=False)
        ws5 = writer.sheets['Border Rewrites']
        add_excel_table(ws5, sugg_df, "BorderRewritesTable")
        
        ws5.column_dimensions['A'].width = 18
        ws5.column_dimensions['C'].width = 18
        ws5.column_dimensions['E'].width = 85

        for r in range(2, len(suggestions) + 2):
            diff = abs(suggestions[r-2][1] - suggestions[r-2][3])
            ws5.cell(row=r, column=5).value = CellRichText([
                "That is a ", TextBlock(bold_inline, f"{diff} address difference"),
                f". Shrink {suggestions[r-2][0]} & Expand {suggestions[r-2][2]}."
            ])

        # --- TAB 6: EXCLUDED AUDIT ---
        if not excluded_gdf.empty:
            excluded_gdf[["HouseNum_Sort", "HouseNum_Suffix_Rank", "HouseNum_Text_Sort"]] = excluded_gdf["HouseNo"].apply(house_number_sort_parts)
            excluded_gdf["Unit_Sort"] = excluded_gdf["Unit"].map(clean_field).str.upper()
            excluded_list_df = excluded_gdf.sort_values(by=["Territory_Name", "Street", "HouseNum_Sort", "HouseNum_Suffix_Rank", "HouseNum_Text_Sort", "Unit_Sort"], kind="stable")
            
            export_ex_df = excluded_list_df[['Territory_Name', 'NWS_Category', 'NWS_Number', 'Mailable_Address', 'Addr_Statu', 'Data_Quality_Flag', 'HouseNo', 'Street', 'Unit', 'Zip_Code']].rename(columns={
                'Territory_Name': 'Territory Name', 
                'Mailable_Address': 'Mailable Address',
                'Addr_Statu': 'Exclusion Reason',
                'Data_Quality_Flag': 'Data Quality Flag'
            })
            export_ex_df.to_excel(writer, sheet_name="Excluded Audit", index=False)
            
            ws6 = writer.sheets['Excluded Audit']
            add_excel_table(ws6, export_ex_df, "ExcludedAuditTable")
            
            for col_letter in ['G', 'H', 'I', 'J']:
                ws6.column_dimensions[col_letter].hidden = True
                
            ws6.column_dimensions['A'].width = 18
            ws6.column_dimensions['B'].width = 15
            ws6.column_dimensions['C'].width = 15
            ws6.column_dimensions['D'].width = 55
            ws6.column_dimensions['E'].width = 25
            ws6.column_dimensions['F'].width = 35
        else:
            pd.DataFrame(columns=["Notice"]).to_excel(writer, sheet_name="Excluded Audit", index=False)
            writer.sheets['Excluded Audit'].cell(row=2, column=1, value="No addresses were excluded in this map area.")

        # --- EXCEL UX POLISH ---
        for tab_name in ["Counts", "Address List", "Apartments", "Border Rewrites", "Excluded Audit"]:
            ws = writer.sheets[tab_name]
            ws.freeze_panes = 'A2'

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

                    if kml_gdf.crs is None:
                        kml_gdf = kml_gdf.set_crs("EPSG:4326", allow_override=True)

                    if parcel_gdf.crs is None:
                        raise ValueError(
                            "The parcel dataset has no CRS. Assign the correct source CRS before processing."
                        )

                    kml_gdf = kml_gdf.copy()
                    kml_gdf["geometry"] = kml_gdf.geometry.make_valid()

                    kml_gdf = kml_gdf[
                        kml_gdf.geometry.notna()
                        & ~kml_gdf.geometry.is_empty
                    ].copy()

                    fallback_names = "Territory_" + kml_gdf.index.to_series().astype(str)

                    if "Name" in kml_gdf.columns:
                        kml_gdf["Territory_Name"] = (
                            kml_gdf["Name"]
                            .replace(r"^\s*$", pd.NA, regex=True)
                            .fillna(fallback_names)
                        )
                    elif "Description" in kml_gdf.columns:
                        kml_gdf["Territory_Name"] = (
                            kml_gdf["Description"]
                            .replace(r"^\s*$", pd.NA, regex=True)
                            .fillna(fallback_names)
                        )
                    else:
                        kml_gdf["Territory_Name"] = fallback_names

                    if parcel_gdf.crs != kml_gdf.crs:
                        parcel_gdf = parcel_gdf.to_crs(kml_gdf.crs)

                    parcel_gdf = parcel_gdf.copy()
                    parcel_gdf["geometry"] = parcel_gdf.geometry.make_valid()

                    parcel_gdf = parcel_gdf[
                        parcel_gdf.geometry.notna()
                        & ~parcel_gdf.geometry.is_empty
                    ].copy()

                    territory_envelope = kml_gdf.geometry.union_all().envelope
                    parcel_gdf = parcel_gdf[
                        parcel_gdf.geometry.intersects(territory_envelope)
                    ].copy()

                    parcel_gdf["_join_point"] = parcel_gdf.geometry.representative_point()
                    parcel_join_gdf = parcel_gdf.set_geometry("_join_point")

                    kml_gdf = kml_gdf.rename(columns={"geometry": "geometry_terr"})
                    kml_gdf = kml_gdf.set_geometry("geometry_terr")

                    joined_gdf = gpd.sjoin(
                        parcel_join_gdf,
                        kml_gdf[["Territory_Name", "geometry_terr"]],
                        how="inner",
                        predicate="covered_by",
                    )

                    joined_gdf = joined_gdf.dropna(subset=["Territory_Name"])

                    joined_gdf = joined_gdf.set_geometry("geometry")
                    joined_gdf = joined_gdf.drop(columns=["_join_point"], errors="ignore")
                    
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
            label="⬇️ Download Excel Analysis",
            data=st.session_state['excel_data'],
            file_name=st.session_state['excel_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

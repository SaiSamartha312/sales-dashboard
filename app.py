import streamlit as st
import pandas as pd
import plotly.express as px

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Sales & Product Dashboard", layout="wide")

st.title("📊 Sales & Inventory Master Dashboard")
st.markdown("Upload your **Pricelist**, **Quotation**, and **OA** Excel files on the left to begin.")

# --- SIDEBAR: FILE UPLOADER ---
st.sidebar.header("1. Upload Data")

pricelist_file = st.sidebar.file_uploader("Upload PRICELIST (Excel)", type=['xlsx'])
quote_file = st.sidebar.file_uploader("Upload QUOTATION List (Excel)", type=['xlsx'])
oa_file = st.sidebar.file_uploader("Upload ORDER ACK (OA) (Excel)", type=['xlsx'])

# --- HELPER FUNCTIONS ---
def normalize_columns(df):
    """
    Cleans column names to make them standard.
    Removes spaces, makes lowercase, handles variations like 'Part nr' vs 'Part No'.
    """
    df.columns = df.columns.str.strip() # Remove empty spaces
    
    # Rename columns to a standard format if they exist
    rename_map = {
        'Part nr': 'Part_Number',
        'Part No': 'Part_Number',
        'Part No.': 'Part_Number',
        'Art. Nr.': 'Article_Code',
        'Art No.': 'Article_Code',
        'Client Name': 'Customer',
        'Client PO': 'PO_Number',
        'Unit wt': 'Unit_Weight',
        'UWT (kg)': 'Unit_Weight',
        'Total Amount': 'Amount',
        'Matl. Code': 'Material_Code',
        'Date': 'Date',
        'Model': 'Model',
        'Description': 'Description'
    }
    df = df.rename(columns=rename_map)
    return df

# --- MAIN LOGIC ---
if pricelist_file and quote_file and oa_file:
    try:
        # 1. LOAD DATA
        df_price = pd.read_excel(pricelist_file)
        df_quote = pd.read_excel(quote_file)
        df_oa = pd.read_excel(oa_file)

        # 2. CLEAN DATA
        df_price = normalize_columns(df_price)
        df_quote = normalize_columns(df_quote)
        df_oa = normalize_columns(df_oa)

        # Convert Dates
        if 'Date' in df_quote.columns:
            df_quote['Date'] = pd.to_datetime(df_quote['Date'], errors='coerce')
            df_quote['Year'] = df_quote['Date'].dt.year

        # 3. CREATE TABS
        tab1, tab2, tab3 = st.tabs(["📈 Sales Matrix", "🔍 Product Lookup", "📄 Raw Data"])

        # --- TAB 1: SALES MATRIX ---
        with tab1:
            st.header("Sales Data Analysis")
            
            # Filter Options
            if 'Year' in df_quote.columns:
                years = sorted(df_quote['Year'].dropna().unique())
                selected_year = st.selectbox("Select Year to Filter", options=["All"] + list(years))
                
                # Filter Data based on Year
                filtered_sales = df_quote.copy()
                if selected_year != "All":
                    filtered_sales = filtered_sales[filtered_sales['Year'] == selected_year]
            else:
                filtered_sales = df_quote.copy()
                st.warning("Could not find a 'Date' column to filter by Year.")

            # Pivot Table: Customer vs Model
            st.subheader(f"Sales Summary")
            
            if not filtered_sales.empty and 'Customer' in filtered_sales.columns and 'Model' in filtered_sales.columns:
                pivot = pd.pivot_table(
                    filtered_sales, 
                    values='Amount', 
                    index='Customer', 
                    columns='Model', 
                    aggfunc='sum', 
                    fill_value=0
                )
                
                # formatting numbers
                st.dataframe(pivot.style.format("{:,.2f}"))
                
                # Chart
                st.subheader("Sales by Customer")
                chart_data = filtered_sales.groupby('Customer')['Amount'].sum().reset_index()
                fig = px.bar(chart_data, x='Customer', y='Amount', title="Total Sales by Customer")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("No data available for the selected filters (or columns Customer/Model missing).")

        # --- TAB 2: PRODUCT LOOKUP ---
        with tab2:
            st.header("Search Database")
            
            search_term = st.text_input("Enter Part Number OR Article Code (e.g., 12345):")

            if search_term:
                # Search in Pricelist (Master Data)
                # We check if the search term is in Part Number OR Article Code
                # We convert to string first to avoid errors
                
                # Create mask safely
                p_mask = df_price['Part_Number'].astype(str).str.contains(search_term, case=False, na=False) if 'Part_Number' in df_price.columns else False
                a_mask = df_price['Article_Code'].astype(str).str.contains(search_term, case=False, na=False) if 'Article_Code' in df_price.columns else False
                
                results = df_price[p_mask | a_mask]

                if not results.empty:
                    st.success(f"Found {len(results)} items in Pricelist:")
                    
                    for index, row in results.iterrows():
                        part_num = row.get('Part_Number', 'N/A')
                        
                        with st.expander(f"Product: {part_num} (Click to Expand)"):
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown(f"**Description:** {row.get('Description', 'N/A')}")
                                st.markdown(f"**Article Code:** {row.get('Article_Code', 'N/A')}")
                                st.markdown(f"**Material Code:** {row.get('Material_Code', 'N/A')}")
                            with col2:
                                st.markdown(f"**Unit Weight:** {row.get('Unit_Weight', 'N/A')}")
                                st.markdown(f"**Model:** {row.get('Model', 'N/A')}")

                            # Find Sales History for this specific Part Number
                            st.markdown("---")
                            st.markdown("#### 📜 Sales History (Quotations)")
                            
                            if 'Part_Number' in df_quote.columns:
                                history = df_quote[df_quote['Part_Number'].astype(str) == str(part_num)]
                                
                                if not history.empty:
                                    cols_to_show = [c for c in ['Date', 'Customer', 'Amount', 'Status'] if c in df_quote.columns]
                                    st.dataframe(history[cols_to_show])
                                else:
                                    st.info("No quotation history found for this part.")

                else:
                    st.error("No product found with that Part Number or Article Code.")

        # --- TAB 3: RAW DATA ---
        with tab3:
            st.header("Raw Data Inspector")
            dataset = st.radio("Select Dataset", ["Pricelist", "Quotations", "Order Acknowledgements"])
            
            if dataset == "Pricelist":
                st.dataframe(df_price)
            elif dataset == "Quotations":
                st.dataframe(df_quote)
            else:
                st.dataframe(df_oa)

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.warning("Tip: Check your column names in Excel. The app looks for 'Part nr', 'Art. Nr.', 'Client Name', etc.")

else:
    st.info("👋 Please upload all 3 Excel files in the sidebar (left) to generate the dashboard.")

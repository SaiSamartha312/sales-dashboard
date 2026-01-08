You are using a much nicer, more advanced version of the code now! That is great.

However, **this new code has the exact same bug** regarding the Excel headers. If your Excel file has a Year (like `2024`) or a Number in the header row, this code will crash with `AttributeError: 'int' object has no attribute 'strip'`.

You need to change **one specific line** inside the `normalize_columns` function.

### The Fix (Short Version)

Look for this block of code (around line 86):

```python
    for standard_name, aliases in mapping.items():
        for col in df.columns:
            # THIS IS THE LINE THAT CAUSES THE ERROR:
            if col.strip().lower() in [alias.lower() for alias in aliases]:
                new_columns[col] = standard_name
                break
```

**Change it to this (add `str(...)`):**

```python
    for standard_name, aliases in mapping.items():
        for col in df.columns:
            # THIS IS THE FIXED LINE:
            if str(col).strip().lower() in [alias.lower() for alias in aliases]:
                new_columns[col] = standard_name
                break
```

---

### The Full Corrected Code (Copy and Paste)

To make it easy, here is the **entire** file with the fix applied. You can simply delete everything in your `app.py` and paste this fresh version.

```python
"""
Sales & Product Data Manager
A Streamlit application for managing sales and product data.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Sales & Product Data Manager",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #667eea;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        background-color: #f0f2f6;
        border-radius: 4px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #667eea;
        color: white;
    }
    div[data-testid="stMetricValue"] {
        font-size: 2rem;
    }
</style>
""", unsafe_allow_html=True)


# Column mappings for normalization
COLUMN_MAPPINGS = {
    'pricelist': {
        'art_nr': ['Art. Nr.', 'Art Nr', 'Article Number', 'Article Code', 'Art No.', 'Art No', 'ArtNr'],
        'description': ['Description', 'Desc', 'Product Description'],
        'part_nr': ['Part nr', 'Part No', 'Part Number', 'Part No.', 'PartNr', 'Part Nr'],
        'unit_wt': ['Unit wt', 'Unit Weight', 'Weight', 'Unit Wt.', 'UnitWt', 'UWT (kg)'],
        'matl_code': ['Matl. Code', 'Material Code', 'Matl Code', 'Material', 'MatlCode'],
        'make': ['Make', 'Brand', 'Manufacturer'],
        'model': ['Model', 'Product Model']
    },
    'quotation': {
        'date': ['Date', 'Quote Date', 'Quotation Date'],
        'client_name': ['Client Name', 'Customer', 'Client', 'Customer Name', 'ClientName'],
        'model': ['Model', 'Product Model'],
        'art_no': ['Art No.', 'Art Nr', 'Article Number', 'Article Code', 'Art No', 'ArtNo'],
        'part_no': ['Part No', 'Part Number', 'Part Nr', 'Part No.', 'PartNo'],
        'total_amount': ['Total Amount', 'Amount', 'Total', 'Value', 'TotalAmount'],
        'vertical': ['Vertical', 'Industry', 'Segment']
    },
    'oa': {
        'client_po': ['Client PO', 'PO Number', 'Purchase Order', 'PO', 'ClientPO'],
        'customer': ['Customer', 'Client Name', 'Client', 'Customer Name'],
        'model': ['Model', 'Product Model'],
        'art_no': ['Art No.', 'Art Nr', 'Article Number', 'Article Code', 'Art No', 'ArtNo'],
        'part_no': ['Part No', 'Part Number', 'Part Nr', 'Part No.', 'PartNo'],
        'total_amount': ['Total Amount', 'Amount', 'Total', 'Value', 'TotalAmount'],
        'status': ['Status', 'Order Status', 'State']
    }
}


def normalize_columns(df: pd.DataFrame, data_type: str) -> pd.DataFrame:
    """Normalize column names based on predefined mappings."""
    mapping = COLUMN_MAPPINGS.get(data_type, {})
    new_columns = {}
    
    for standard_name, aliases in mapping.items():
        for col in df.columns:
            # FIX APPLIED HERE: Added str(col) to prevent crashing on Number headers
            if str(col).strip().lower() in [alias.lower() for alias in aliases]:
                new_columns[col] = standard_name
                break
    
    df_normalized = df.rename(columns=new_columns)
    return df_normalized


def load_file(uploaded_file, data_type: str) -> pd.DataFrame:
    """Load CSV or Excel file and normalize columns."""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
        else:
            st.error(f"Unsupported file format: {uploaded_file.name}")
            return pd.DataFrame()
        
        df = normalize_columns(df, data_type)
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return pd.DataFrame()


def parse_amount(value) -> float:
    """Parse amount values, handling various formats."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    # Remove currency symbols and commas
    cleaned = str(value).replace('$', '').replace(',', '').replace(' ', '')
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def get_year_from_date(date_value) -> int:
    """Extract year from various date formats."""
    if pd.isna(date_value):
        return None
    try:
        if isinstance(date_value, datetime):
            return date_value.year
        date_parsed = pd.to_datetime(date_value)
        return date_parsed.year
    except:
        return None


def format_currency(value: float) -> str:
    """Format value as currency."""
    return f"${value:,.2f}"


# Initialize session state
if 'pricelist' not in st.session_state:
    st.session_state.pricelist = pd.DataFrame()
if 'quotation' not in st.session_state:
    st.session_state.quotation = pd.DataFrame()
if 'oa' not in st.session_state:
    st.session_state.oa = pd.DataFrame()


# Sidebar - File Uploaders
with st.sidebar:
    st.markdown("## 📁 Data Manager")
    st.markdown("---")
    
    # Pricelist Upload
    st.markdown("### 🏷️ Pricelist")
    pricelist_file = st.file_uploader(
        "Upload Pricelist",
        type=['csv', 'xlsx', 'xls'],
        key='pricelist_uploader',
        help="Upload your product pricelist (CSV or Excel)"
    )
    if pricelist_file:
        st.session_state.pricelist = load_file(pricelist_file, 'pricelist')
        if not st.session_state.pricelist.empty:
            st.success(f"✅ {len(st.session_state.pricelist)} products loaded")
    
    st.markdown("---")
    
    # Quotation Upload
    st.markdown("### 📋 Quotation List")
    quotation_file = st.file_uploader(
        "Upload Quotation List",
        type=['csv', 'xlsx', 'xls'],
        key='quotation_uploader',
        help="Upload your quotation list (CSV or Excel)"
    )
    if quotation_file:
        st.session_state.quotation = load_file(quotation_file, 'quotation')
        if not st.session_state.quotation.empty:
            st.success(f"✅ {len(st.session_state.quotation)} quotations loaded")
    
    st.markdown("---")
    
    # Order Acknowledgement Upload
    st.markdown("### 📦 Order Acknowledgement")
    oa_file = st.file_uploader(
        "Upload Order Acknowledgement",
        type=['csv', 'xlsx', 'xls'],
        key='oa_uploader',
        help="Upload your order acknowledgement (CSV or Excel)"
    )
    if oa_file:
        st.session_state.oa = load_file(oa_file, 'oa')
        if not st.session_state.oa.empty:
            st.success(f"✅ {len(st.session_state.oa)} orders loaded")
    
    st.markdown("---")
    
    # Data Summary
    st.markdown("### 📊 Data Summary")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Products", len(st.session_state.pricelist))
        st.metric("Orders", len(st.session_state.oa))
    with col2:
        st.metric("Quotes", len(st.session_state.quotation))
        if not st.session_state.oa.empty and 'total_amount' in st.session_state.oa.columns:
            total_rev = st.session_state.oa['total_amount'].apply(parse_amount).sum()
            st.metric("Revenue", f"${total_rev/1000:.1f}K")
    
    st.markdown("---")
    
    # Clear Data Button
    if st.button("🗑️ Clear All Data", type="secondary", use_container_width=True):
        st.session_state.pricelist = pd.DataFrame()
        st.session_state.quotation = pd.DataFrame()
        st.session_state.oa = pd.DataFrame()
        st.rerun()


# Main Content
st.markdown('<p class="main-header">📊 Sales & Product Dashboard</p>', unsafe_allow_html=True)

# Create tabs
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Dashboard",
    "🔍 Product Lookup",
    "📊 Sales Matrix",
    "📋 Data Viewer"
])


# ============= TAB 1: DASHBOARD =============
with tab1:
    st.markdown("### Key Performance Indicators")
    
    # KPI Cards
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="Total Products",
            value=len(st.session_state.pricelist),
            delta=None
        )
    
    with col2:
        st.metric(
            label="Total Quotations",
            value=len(st.session_state.quotation),
            delta=None
        )
    
    with col3:
        st.metric(
            label="Total Orders",
            value=len(st.session_state.oa),
            delta=None
        )
    
    with col4:
        total_revenue = 0
        if not st.session_state.oa.empty and 'total_amount' in st.session_state.oa.columns:
            total_revenue = st.session_state.oa['total_amount'].apply(parse_amount).sum()
        st.metric(
            label="Total Revenue",
            value=format_currency(total_revenue),
            delta=None
        )
    
    st.markdown("---")
    
    # Charts
    if not st.session_state.oa.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 Sales by Customer")
            if 'customer' in st.session_state.oa.columns and 'total_amount' in st.session_state.oa.columns:
                df_customer = st.session_state.oa.copy()
                df_customer['amount'] = df_customer['total_amount'].apply(parse_amount)
                customer_sales = df_customer.groupby('customer')['amount'].sum().sort_values(ascending=False).head(10)
                
                fig = px.bar(
                    x=customer_sales.index,
                    y=customer_sales.values,
                    labels={'x': 'Customer', 'y': 'Sales Amount'},
                    color=customer_sales.values,
                    color_continuous_scale='Viridis'
                )
                fig.update_layout(showlegend=False, coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Upload Order data with 'Customer' and 'Total Amount' columns")
        
        with col2:
            st.markdown("#### 🥧 Sales by Model")
            if 'model' in st.session_state.oa.columns and 'total_amount' in st.session_state.oa.columns:
                df_model = st.session_state.oa.copy()
                df_model['amount'] = df_model['total_amount'].apply(parse_amount)
                model_sales = df_model.groupby('model')['amount'].sum().sort_values(ascending=False).head(8)
                
                fig = px.pie(
                    names=model_sales.index,
                    values=model_sales.values,
                    hole=0.4
                )
                fig.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Upload Order data with 'Model' and 'Total Amount' columns")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📈 Sales Trend")
            # Combine quotations and orders for trend
            trend_data = []
            
            if not st.session_state.quotation.empty and 'date' in st.session_state.quotation.columns:
                df_q = st.session_state.quotation.copy()
                df_q['amount'] = df_q['total_amount'].apply(parse_amount) if 'total_amount' in df_q.columns else 0
                df_q['date_parsed'] = pd.to_datetime(df_q['date'], errors='coerce')
                df_q['month'] = df_q['date_parsed'].dt.to_period('M').astype(str)
                df_q['source'] = 'Quotation'
                trend_data.append(df_q[['month', 'amount', 'source']].dropna())
            
            if trend_data:
                df_trend = pd.concat(trend_data, ignore_index=True)
                monthly = df_trend.groupby(['month', 'source'])['amount'].sum().reset_index()
                
                fig = px.line(
                    monthly,
                    x='month',
                    y='amount',
                    color='source',
                    markers=True,
                    labels={'month': 'Month', 'amount': 'Amount'}
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Upload Quotation data with dates to see trends")
        
        with col2:
            st.markdown("#### 📋 Order Status")
            if 'status' in st.session_state.oa.columns:
                status_counts = st.session_state.oa['status'].value_counts()
                
                fig = px.bar(
                    x=status_counts.index,
                    y=status_counts.values,
                    labels={'x': 'Status', 'y': 'Count'},
                    color=status_counts.index,
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Upload Order data with 'Status' column")
    else:
        st.info("📁 Upload data files from the sidebar to view the dashboard")


# ============= TAB 2: PRODUCT LOOKUP =============
with tab2:
    st.markdown("### 🔍 Product Search")
    st.markdown("Search by Part Number or Article Code to find product details and history.")
    
    col1, col2 = st.columns([3, 1])
    with col1:
        search_query = st.text_input(
            "Enter Part Number or Article Code",
            placeholder="e.g., PN12345 or ART-001",
            key="product_search"
        )
    with col2:
        search_button = st.button("🔍 Search", type="primary", use_container_width=True)
    
    if search_query:
        query_lower = search_query.lower().strip()
        
        # Search in Pricelist
        matching_products = pd.DataFrame()
        if not st.session_state.pricelist.empty:
            df_pl = st.session_state.pricelist
            mask = pd.Series([False] * len(df_pl))
            
            if 'part_nr' in df_pl.columns:
                mask |= df_pl['part_nr'].astype(str).str.lower().str.contains(query_lower, na=False)
            if 'art_nr' in df_pl.columns:
                mask |= df_pl['art_nr'].astype(str).str.lower().str.contains(query_lower, na=False)
            
            matching_products = df_pl[mask]
        
        # Search in Quotations
        related_quotes = pd.DataFrame()
        if not st.session_state.quotation.empty:
            df_q = st.session_state.quotation
            mask = pd.Series([False] * len(df_q))
            
            if 'part_no' in df_q.columns:
                mask |= df_q['part_no'].astype(str).str.lower().str.contains(query_lower, na=False)
            if 'art_no' in df_q.columns:
                mask |= df_q['art_no'].astype(str).str.lower().str.contains(query_lower, na=False)
            
            related_quotes = df_q[mask]
        
        # Search in Orders
        related_orders = pd.DataFrame()
        if not st.session_state.oa.empty:
            df_oa = st.session_state.oa
            mask = pd.Series([False] * len(df_oa))
            
            if 'part_no' in df_oa.columns:
                mask |= df_oa['part_no'].astype(str).str.lower().str.contains(query_lower, na=False)
            if 'art_no' in df_oa.columns:
                mask |= df_oa['art_no'].astype(str).str.lower().str.contains(query_lower, na=False)
            
            related_orders = df_oa[mask]
        
        # Display Results
        if matching_products.empty and related_quotes.empty and related_orders.empty:
            st.warning(f"No results found for '{search_query}'")
        else:
            # Product Details
            if not matching_products.empty:
                st.markdown("#### 📦 Product Details")
                st.dataframe(
                    matching_products,
                    use_container_width=True,
                    hide_index=True
                )
                
                # Show Many-to-Many Relationships
                st.markdown("#### 🔗 Relationships (Many-to-Many)")
                col1, col2 = st.columns(2)
                
                with col1:
                    if 'part_nr' in matching_products.columns:
                        part_numbers = matching_products['part_nr'].dropna().unique()
                        st.markdown("**Related Part Numbers:**")
                        for pn in part_numbers:
                            st.markdown(f"- `{pn}`")
                
                with col2:
                    if 'art_nr' in matching_products.columns:
                        art_numbers = matching_products['art_nr'].dropna().unique()
                        st.markdown("**Related Article Codes:**")
                        for an in art_numbers:
                            st.markdown(f"- `{an}`")
            
            # Quotation History
            if not related_quotes.empty:
                st.markdown("#### 📋 Quotation History")
                
                # Merge with pricelist for additional details
                if not st.session_state.pricelist.empty and 'part_nr' in st.session_state.pricelist.columns:
                    pricelist_attrs = st.session_state.pricelist[['part_nr', 'description', 'unit_wt', 'matl_code']].drop_duplicates('part_nr')
                    if 'part_no' in related_quotes.columns:
                        related_quotes = related_quotes.merge(
                            pricelist_attrs,
                            left_on='part_no',
                            right_on='part_nr',
                            how='left'
                        )
                
                st.dataframe(related_quotes, use_container_width=True, hide_index=True)
                
                # Summary stats
                if 'total_amount' in related_quotes.columns:
                    total = related_quotes['total_amount'].apply(parse_amount).sum()
                    st.info(f"📊 Total Quoted Amount: {format_currency(total)} across {len(related_quotes)} quotations")
            
            # Order History
            if not related_orders.empty:
                st.markdown("#### 🛒 Order History")
                
                # Merge with pricelist for additional details
                if not st.session_state.pricelist.empty and 'part_nr' in st.session_state.pricelist.columns:
                    pricelist_attrs = st.session_state.pricelist[['part_nr', 'description', 'unit_wt', 'matl_code']].drop_duplicates('part_nr')
                    if 'part_no' in related_orders.columns:
                        related_orders = related_orders.merge(
                            pricelist_attrs,
                            left_on='part_no',
                            right_on='part_nr',
                            how='left'
                        )
                
                st.dataframe(related_orders, use_container_width=True, hide_index=True)
                
                # Summary stats
                if 'total_amount' in related_orders.columns:
                    total = related_orders['total_amount'].apply(parse_amount).sum()
                    st.success(f"💰 Total Order Amount: {format_currency(total)} across {len(related_orders)} orders")
    else:
        st.info("👆 Enter a Part Number or Article Code to search")


# ============= TAB 3: SALES MATRIX =============
with tab3:
    st.markdown("### 📊 Sales Pivot Table")
    
    # Filters
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        data_source = st.selectbox(
            "Data Source",
            options=['Orders (OA)', 'Quotations'],
            key='matrix_source'
        )
    
    source_df = st.session_state.oa if data_source == 'Orders (OA)' else st.session_state.quotation
    customer_col = 'customer' if data_source == 'Orders (OA)' else 'client_name'
    date_col = 'date' if data_source == 'Quotations' else None
    
    if not source_df.empty:
        # Get unique years
        years = ['All Years']
        if date_col and date_col in source_df.columns:
            source_df['_year'] = source_df[date_col].apply(get_year_from_date)
            unique_years = source_df['_year'].dropna().unique()
            years.extend(sorted([int(y) for y in unique_years], reverse=True))
        
        # Get unique customers
        customers = ['All Customers']
        if customer_col in source_df.columns:
            unique_customers = source_df[customer_col].dropna().unique()
            customers.extend(sorted(unique_customers))
        
        with col2:
            selected_year = st.selectbox("Year", options=years, key='matrix_year')
        
        with col3:
            selected_customer = st.selectbox("Customer", options=customers, key='matrix_customer')
        
        with col4:
            pivot_type = st.selectbox(
                "Pivot By",
                options=['Customer × Year', 'Customer × Model', 'Model × Year'],
                key='matrix_pivot'
            )
        
        # Filter data
        filtered_df = source_df.copy()
        
        if selected_year != 'All Years' and '_year' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['_year'] == selected_year]
        
        if selected_customer != 'All Customers' and customer_col in filtered_df.columns:
            filtered_df = filtered_df[filtered_df[customer_col] == selected_customer]
        
        # Parse amounts
        if 'total_amount' in filtered_df.columns:
            filtered_df['_amount'] = filtered_df['total_amount'].apply(parse_amount)
        else:
            filtered_df['_amount'] = 0
        
        # Create pivot table
        if not filtered_df.empty:
            st.markdown("---")
            
            try:
                if pivot_type == 'Customer × Year':
                    if customer_col in filtered_df.columns and '_year' in filtered_df.columns:
                        pivot = pd.pivot_table(
                            filtered_df,
                            values='_amount',
                            index=customer_col,
                            columns='_year',
                            aggfunc='sum',
                            fill_value=0,
                            margins=True,
                            margins_name='Total'
                        )
                    else:
                        st.warning("Required columns not available for this pivot type")
                        pivot = pd.DataFrame()
                
                elif pivot_type == 'Customer × Model':
                    if customer_col in filtered_df.columns and 'model' in filtered_df.columns:
                        pivot = pd.pivot_table(
                            filtered_df,
                            values='_amount',
                            index=customer_col,
                            columns='model',
                            aggfunc='sum',
                            fill_value=0,
                            margins=True,
                            margins_name='Total'
                        )
                    else:
                        st.warning("Required columns not available for this pivot type")
                        pivot = pd.DataFrame()
                
                else:  # Model × Year
                    if 'model' in filtered_df.columns and '_year' in filtered_df.columns:
                        pivot = pd.pivot_table(
                            filtered_df,
                            values='_amount',
                            index='model',
                            columns='_year',
                            aggfunc='sum',
                            fill_value=0,
                            margins=True,
                            margins_name='Total'
                        )
                    else:
                        st.warning("Required columns not available for this pivot type")
                        pivot = pd.DataFrame()
                
                if not pivot.empty:
                    # Format as currency
                    pivot_formatted = pivot.applymap(lambda x: format_currency(x) if isinstance(x, (int, float)) else x)
                    
                    st.markdown("#### 📋 Pivot Table")
                    st.dataframe(pivot_formatted, use_container_width=True)
                    
                    # Export button
                    csv = pivot.to_csv()
                    st.download_button(
                        label="📥 Download CSV",
                        data=csv,
                        file_name="sales_matrix.csv",
                        mime="text/csv"
                    )
                    
                    # Visualization
                    st.markdown("#### 📊 Visualization")
                    
                    # Remove Total row/column for chart
                    pivot_chart = pivot.drop('Total', axis=0, errors='ignore').drop('Total', axis=1, errors='ignore')
                    
                    if not pivot_chart.empty:
                        fig = px.bar(
                            pivot_chart.reset_index(),
                            x=pivot_chart.index.name or 'index',
                            y=pivot_chart.columns.tolist(),
                            barmode='group',
                            title=f"Sales by {pivot_type}"
                        )
                        fig.update_layout(
                            xaxis_title=pivot_chart.index.name or 'Category',
                            yaxis_title='Sales Amount',
                            legend_title='Category'
                        )
                        st.plotly_chart(fig, use_container_width=True)
            
            except Exception as e:
                st.error(f"Error creating pivot table: {str(e)}")
        else:
            st.warning("No data matching the selected filters")
    else:
        st.info("📁 Upload Order or Quotation data to view the sales matrix")


# ============= TAB 4: DATA VIEWER =============
with tab4:
    st.markdown("### 📋 Data Viewer")
    
    # Data source selector
    col1, col2 = st.columns([2, 3])
    
    with col1:
        view_source = st.selectbox(
            "Select Data Source",
            options=['Pricelist', 'Quotations', 'Orders (OA)', 'Merged View'],
            key='viewer_source'
        )
    
    with col2:
        filter_text = st.text_input(
            "Filter Data",
            placeholder="Type to filter across all columns...",
            key='viewer_filter'
        )
    
    # Get the appropriate dataframe
    if view_source == 'Pricelist':
        view_df = st.session_state.pricelist.copy()
    elif view_source == 'Quotations':
        view_df = st.session_state.quotation.copy()
    elif view_source == 'Orders (OA)':
        view_df = st.session_state.oa.copy()
    else:  # Merged View
        # Merge OA and Quotations with Pricelist attributes
        merged_dfs = []
        
        if not st.session_state.oa.empty:
            df_oa = st.session_state.oa.copy()
            df_oa['_source'] = 'Order'
            
            if not st.session_state.pricelist.empty and 'part_nr' in st.session_state.pricelist.columns:
                pricelist_attrs = st.session_state.pricelist[['part_nr', 'description', 'unit_wt', 'matl_code']].drop_duplicates('part_nr')
                if 'part_no' in df_oa.columns:
                    df_oa = df_oa.merge(pricelist_attrs, left_on='part_no', right_on='part_nr', how='left')
            
            merged_dfs.append(df_oa)
        
        if not st.session_state.quotation.empty:
            df_q = st.session_state.quotation.copy()
            df_q['_source'] = 'Quotation'
            
            if not st.session_state.pricelist.empty and 'part_nr' in st.session_state.pricelist.columns:
                pricelist_attrs = st.session_state.pricelist[['part_nr', 'description', 'unit_wt', 'matl_code']].drop_duplicates('part_nr')
                if 'part_no' in df_q.columns:
                    df_q = df_q.merge(pricelist_attrs, left_on='part_no', right_on='part_nr', how='left')
            
            merged_dfs.append(df_q)
        
        if merged_dfs:
            view_df = pd.concat(merged_dfs, ignore_index=True)
        else:
            view_df = pd.DataFrame()
    
    if not view_df.empty:
        # Apply filter
        if filter_text:
            mask = pd.Series([False] * len(view_df))
            for col in view_df.columns:
                mask |= view_df[col].astype(str).str.lower().str.contains(filter_text.lower(), na=False)
            view_df = view_df[mask]
        
        # Display stats
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Rows", len(view_df))
        with col2:
            st.metric("Columns", len(view_df.columns))
        with col3:
            if 'total_amount' in view_df.columns:
                total = view_df['total_amount'].apply(parse_amount).sum()
                st.metric("Total Amount", format_currency(total))
        
        st.markdown("---")
        
        # Display dataframe
        st.dataframe(
            view_df,
            use_container_width=True,
            hide_index=True,
            height=500
        )
        
        # Export options
        col1, col2 = st.columns([1, 4])
        with col1:
            csv = view_df.to_csv(index=False)
            st.download_button(
                label="📥 Export CSV",
                data=csv,
                file_name=f"{view_source.lower().replace(' ', '_')}_export.csv",
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info(f"📁 No data available for {view_source}. Upload files from the sidebar.")


# Footer
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #888; padding: 1rem;">
        <p>📊 Sales & Product Data Manager | Built with Streamlit</p>
    </div>
    """,
    unsafe_allow_html=True
)
```

import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# Initialize session state
if 'selected_years' not in st.session_state:
    st.session_state.selected_years = []
if 'locked_year' not in st.session_state:
    st.session_state.locked_year = None

st.set_page_config(layout="wide", page_title="Solar Generation Map")
st.title("Solar Generation by State")

# Load data
@st.cache_data
def load_and_process_data():
    years = list(range(2012, 2025))
    sheet_names = [f"{year}_Final" if year != 2024 else "2024_Preliminary" for year in years]
    all_years_data = []
    
    excel_file = pd.ExcelFile("generation_monthly.xlsx")
    
    for year, sheet in zip(years, sheet_names):
        df = pd.read_excel(excel_file, sheet_name=sheet, skiprows=4)
        
        # Get total energy
        total_energy = df[
            (df['TYPE OF PRODUCER'] == 'Total Electric Power Industry') & 
            (df['ENERGY SOURCE'] == 'Total')
        ].groupby('STATE')['generation'].sum().reset_index()
        
        # Get solar energy
        solar_energy = df[
            (df['TYPE OF PRODUCER'] == 'Total Electric Power Industry') & 
            (df['ENERGY SOURCE'] == 'Solar Thermal and Photovoltaic')
        ].groupby('STATE')['generation'].sum().reset_index()
        
        merged_data = pd.merge(total_energy, solar_energy, 
                             on='STATE', 
                             suffixes=('_total', '_solar'),
                             how='left')
        
        merged_data['generation_solar'] = merged_data['generation_solar'].fillna(0)
        merged_data['solar_percentage'] = (merged_data['generation_solar'] / 
                                         merged_data['generation_total'] * 100).round(2)
        merged_data['Year'] = year
        all_years_data.append(merged_data)
    
    final_df = pd.concat(all_years_data, ignore_index=True)
    return final_df[final_df['STATE'] != 'US-TOTAL']

# Load data once
final_df = load_and_process_data()

# Year selection
st.write("Select years:")

# Create button container
cols = st.columns([20, 1, 4])

with cols[0]:
    # Create year buttons in a single row
    year_cols = st.columns(len(years))
    for idx, year in enumerate(years):
        with year_cols[idx]:
            # Simple button with clear state management
            if st.button(
                str(year),
                key=f'year_{year}',
                type="primary" if year in st.session_state.get('selected_years', []) else "secondary",
                use_container_width=True
            ):
                # If no years are selected, select this year
                if not st.session_state.get('selected_years', []):
                    st.session_state.selected_years = [year]
                    st.rerun()
                
                # If this year is already selected
                elif year in st.session_state.selected_years:
                    # If it's the only year selected, prepare for comparison
                    if len(st.session_state.selected_years) == 1:
                        st.session_state.locked_year = year
                        st.rerun()
                    # If we're already comparing, reset to just this year
                    else:
                        st.session_state.selected_years = [year]
                        st.session_state.locked_year = None
                        st.rerun()
                
                # If another year is selected
                else:
                    # If we're in comparison mode, add this year
                    if st.session_state.get('locked_year', None) is not None:
                        st.session_state.selected_years = [st.session_state.locked_year, year]
                        st.rerun()
                    # Otherwise, switch to this year
                    else:
                        st.session_state.selected_years = [year]
                        st.rerun()

# Reset button
with cols[1]:
    if st.button('↺', type="secondary", help="Reset selection"):
        st.session_state.selected_years = []
        st.session_state.locked_year = None

# Info text
with cols[2]:
    if len(st.session_state.get('selected_years', [])) == 2:
        year1, year2 = min(st.session_state.selected_years), max(st.session_state.selected_years)
        st.write(f"Comparing {year1} → {year2}")
    elif st.session_state.get('locked_year', None) is not None:
        st.write(f"Select year to compare with {st.session_state.locked_year}")
    elif len(st.session_state.get('selected_years', [])) == 1:
        st.write(f"Viewing {st.session_state.selected_years[0]}")
    else:
        st.write("Select a year")

# Add CSS to make the buttons more compact
st.markdown("""
    <style>
    .stButton button {
        padding: 2px 8px;
        font-size: 12px;
        height: auto;
        min-height: 25px;
    }
    
    .stButton {
        margin: 0;
        padding: 0;
    }
    
    [data-testid="column"] {
        padding: 0 2px;
    }
    </style>
""", unsafe_allow_html=True)

# Create visualization based on selection
if len(st.session_state.get('selected_years', [])) == 1:
    # Single year view - handles both manual selection and animation
    current_year = st.session_state.get('selected_years', [2024])[0]
    year_data = final_df[final_df['Year'] == current_year]
    
    # Create the map
    fig = px.choropleth(year_data,
                        locations='STATE',
                        locationmode='USA-states',
                        color='solar_percentage',
                        scope="usa",
                        color_continuous_scale=[
                            [0, 'white'],      # 0% = white
                            [1, 'yellow']      # 100% = yellow
                        ],
                        title=f'Solar Generation Percentage by State ({current_year})',
                        labels={'solar_percentage': 'Solar Generation (%)'},
                        range_color=[-0.1, final_df['solar_percentage'].max()]
                        )

elif len(st.session_state.get('selected_years', [])) == 2:
    # Percent change view between two years
    year1, year2 = min(st.session_state.selected_years), max(st.session_state.selected_years)
    data1 = final_df[final_df['Year'] == year1].set_index('STATE')['solar_percentage']
    data2 = final_df[final_df['Year'] == year2].set_index('STATE')['solar_percentage']
    
    # Calculate difference in percentages
    change_data = pd.DataFrame({
        'STATE': data2.index,
        'initial_percentage': data1,
        'final_percentage': data2,
        'difference': (data2 - data1).round(2)  # Simple difference instead of percent change
    }).reset_index(drop=True)
    
    fig = px.choropleth(change_data,
                        locations='STATE',
                        locationmode='USA-states',
                        color='difference',
                        scope="usa",
                        color_continuous_scale=[
                            [0, 'red'],        # Lowest values
                            [0.5, 'yellow'],   # Middle values
                            [1, 'blue']        # Highest values
                        ],
                        title=f'Change in Solar Generation ({year1} to {year2})',
                        labels={'difference': 'Change in Percentage Points'}
                        )

else:
    # Default view (2024)
    year_data = final_df[final_df['Year'] == 2024]
    fig = px.choropleth(year_data,
                        locations='STATE',
                        locationmode='USA-states',
                        color='solar_percentage',
                        scope="usa",
                        color_continuous_scale=[
                            [0, 'white'],      # 0% = white
                            [1, 'yellow']      # 100% = yellow
                        ],
                        title='Solar Generation Percentage by State (2024)',
                        labels={'solar_percentage': 'Solar Generation (%)'},
                        range_color=[-0.1, final_df['solar_percentage'].max()]
                        )

# Update layout
fig.update_layout(
    geo=dict(
        scope='usa',
        projection=dict(type='albers usa'),
        showlakes=True,
        lakecolor='rgb(255, 255, 255)',
    ),
)

# Display the map
st.plotly_chart(fig, use_container_width=True)

# Now show the comparison table if in comparison mode
if len(st.session_state.get('selected_years', [])) == 2:
    # Update the data display
    st.subheader(f"Changes from {year1} to {year2}")
    display_data = change_data.sort_values('difference', ascending=False)
    
    # Format the table
    formatted_data = display_data[['STATE', 'initial_percentage', 'final_percentage', 'difference']].copy()
    formatted_data.columns = ['State', f'{year1} (%)', f'{year2} (%)', 'Change (pp)']
    
    st.dataframe(
        formatted_data,
        column_config={
            f"{year1} (%)": st.column_config.NumberColumn(format="%.2f%%"),
            f"{year2} (%)": st.column_config.NumberColumn(format="%.2f%%"),
            "Change (pp)": st.column_config.NumberColumn(format="%.2f")
        },
        hide_index=True
    )

# Now show rankings and statistics if it's a single year view
if len(st.session_state.get('selected_years', [])) == 1:
    st.subheader(f"State Rankings for {st.session_state.selected_years[0]}")
    
    # Create ranking table
    display_data = year_data.sort_values('solar_percentage', ascending=False)
    display_data['rank'] = range(1, len(display_data) + 1)
    display_data['solar_percentage'] = display_data['solar_percentage'].round(2)
    
    # Format the table
    formatted_data = display_data[['rank', 'STATE', 'solar_percentage']].copy()
    formatted_data.columns = ['Rank', 'State', 'Solar Generation (%)']
    
    # Display with custom formatting
    st.dataframe(
        formatted_data,
        column_config={
            "Rank": st.column_config.NumberColumn(format="%d"),
            "Solar Generation (%)": st.column_config.NumberColumn(format="%.2f%%")
        },
        hide_index=True
    )
    
    # Statistics for single year
    st.subheader("Summary Statistics")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Highest Solar %", 
                f"{year_data['solar_percentage'].max():.2f}%",
                f"{year_data.loc[year_data['solar_percentage'].idxmax(), 'STATE']}")
    with col2:
        st.metric("Average Solar %", 
                f"{year_data['solar_percentage'].mean():.2f}%")
    with col3:
        st.metric("Total States with Solar", 
                f"{(year_data['generation_solar'] > 0).sum()}")
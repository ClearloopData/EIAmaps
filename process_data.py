import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill, Font
import os
import sys

def create_clean_excel():
    print("Starting Excel file creation...")
    
    output_file = 'solar_generation_summary_EIC.xlsx'
    input_file = r"C:\Users\lofo6\Documents\Work Data\generation_monthly.xlsx"
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file not found at {input_file}")
        return
    
    print(f"Input file found at {input_file}")
    
    # Remove existing output file if it exists
    if os.path.exists(output_file):
        try:
            os.remove(output_file)
            print(f"Removed existing output file: {output_file}")
        except PermissionError:
            print(f"Error: Please close {output_file} if it's open and try again.")
            return
        except Exception as e:
            print(f"Error removing existing file: {str(e)}")
            return
    
    # Read the original Excel file
    years = list(range(2012, 2025))
    sheet_names = [f"{year}_Final" if year != 2024 else "2024_Preliminary" for year in years]
    
    try:
        print("Creating new Excel file...")
        with pd.ExcelWriter(
            output_file,
            engine='openpyxl',
            mode='w'
        ) as writer:
            for year, sheet in zip(years, sheet_names):
                print(f"Processing year {year}...")
                
                try:
                    # Read the sheet
                    df = pd.read_excel(input_file, 
                                     sheet_name=sheet, 
                                     skiprows=4)
                except Exception as e:
                    print(f"Error reading sheet {sheet}: {str(e)}")
                    raise
                
                # Calculate total energy by state
                total_energy = df[
                    (df['TYPE OF PRODUCER'] == 'Total Electric Power Industry') & 
                    (df['ENERGY SOURCE'] == 'Total')
                ].groupby('STATE')['generation'].sum().reset_index()
                
                # Calculate solar energy by state
                solar_energy = df[
                    (df['TYPE OF PRODUCER'] == 'Total Electric Power Industry') & 
                    (df['ENERGY SOURCE'] == 'Solar Thermal and Photovoltaic')
                ].groupby('STATE')['generation'].sum().reset_index()
                
                # Merge and calculate percentage
                merged_data = pd.merge(total_energy, solar_energy, 
                                     on='STATE', 
                                     suffixes=('_total', '_solar'),
                                     how='left')
                
                # Fill NaN values with 0 for states with no solar
                merged_data['generation_solar'] = merged_data['generation_solar'].fillna(0)
                
                # Calculate percentage
                merged_data['solar_percentage'] = (merged_data['generation_solar'] / 
                                                 merged_data['generation_total'] * 100).round(2)
                
                # Separate US-TOTAL and other states
                us_total = merged_data[merged_data['STATE'] == 'US-TOTAL']
                states_data = merged_data[merged_data['STATE'] != 'US-TOTAL']
                
                # Sort states alphabetically
                states_data = states_data.sort_values('STATE')
                
                # Combine back with US-TOTAL at the end
                merged_data = pd.concat([states_data, us_total])
                
                # Rename columns for clarity
                merged_data.columns = ['State', 'Total Generation (MWh)', 
                                     'Solar Generation (MWh)', 'Solar Percentage (%)']
                
                # Write to Excel
                merged_data.to_excel(writer, 
                                   sheet_name=str(year), 
                                   index=False)
                
                # Get the worksheet
                worksheet = writer.sheets[str(year)]
                
                # Define colors for headers
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                
                # Style the headers
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                
                # Adjust column widths
                worksheet.column_dimensions['A'].width = 20
                worksheet.column_dimensions['B'].width = 25
                worksheet.column_dimensions['C'].width = 25
                worksheet.column_dimensions['D'].width = 20
                
                # Freeze the header row
                worksheet.freeze_panes = 'A2'
                
                print(f"Completed processing for year {year}")
        
        # Verify file was created
        if os.path.exists(output_file):
            print(f"Successfully created {output_file}")
            print(f"File size: {os.path.getsize(output_file)} bytes")
        else:
            print("Error: File was not created despite no errors")
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        print(f"Error location: {sys.exc_info()[2].tb_frame.f_code.co_name}")
        if os.path.exists(output_file):
            os.remove(output_file)
            print("Cleaned up partial output file")

if __name__ == "__main__":
    create_clean_excel() 
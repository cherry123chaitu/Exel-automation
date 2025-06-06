import pandas as pd
import os
script_dir = os.path.dirname(os.path.abspath('StudentSort.py'))
mainfiles_path = os.path.join(script_dir, '..', 'MainFiles')
residuals_path=os.path.join(script_dir,'..','Residuals')
main_excel_file = os.path.join(mainfiles_path, 'ProjectMain.xlsx')
# Read the main Excel file into a DataFrame
 # Update with the path to your main Excel file
df = pd.read_excel(main_excel_file)

# Sort the data based on citizenship
df_sorted = df.sort_values(by='Citizenship')

# Group the data based on citizenship
grouped_citizenship = df_sorted.groupby('Citizenship')

# Create a new Excel file for each citizenship group
for citizenship, group_df in grouped_citizenship:
    # Define the output file name for each citizenship group
    output_excel_file = os.path.join(residuals_path, f'Projects for {citizenship} Citizens.xlsx')  
    
    # Save the group data to a new Excel file
    group_df.to_excel(output_excel_file, index=False)

    print(f"Data for citizenship '{citizenship}' has been exported to: {output_excel_file}")



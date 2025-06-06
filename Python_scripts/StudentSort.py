import pandas as pd
import os

# Read the main Excel file into a DataFrame
script_dir = os.path.dirname(os.path.abspath('StudentSort.py'))
mainfiles_path = os.path.join(script_dir, '..', 'MainFiles')
residuals_path=os.path.join(script_dir,'..','Residuals')
main_excel_file = os.path.join(mainfiles_path, 'StudentsMain.xlsx') # Update with the path to your main Excel file
df = pd.read_excel(main_excel_file)

# Sort the data based on citizenship
df_sorted = df.sort_values(by='CTIZ')

# Group the data based on citizenship
grouped_citizenship = df_sorted.groupby('CTIZ')

# Create a new Excel file for each citizenship group
for citizenship, group_df in grouped_citizenship:
    # Define the output file name for each citizenship group
    output_excel_file = os.path.join(residuals_path, f'Student {citizenship} Citizenship_.xlsx')  
    
    # Save the group data to a new Excel file
    group_df.to_excel(output_excel_file, index=False)

    print(f"Data for citizenship '{citizenship}' has been exported to: {output_excel_file}")

# Read and sort data for 'C' citizenship
output_excel_file_c = os.path.join(residuals_path, 'Student C Citizenship_.xlsx')
df_c = pd.read_excel(output_excel_file_c)
gpa_c_sort = df_c.sort_values(by='GPA', ascending=False)
gpa_c_sort.to_excel(output_excel_file_c, index=False)
print("Sorted C")

# Read and sort data for 'N' citizenship
output_excel_file_n = os.path.join(residuals_path, 'Student N Citizenship_.xlsx')
df_n = pd.read_excel(output_excel_file_n)
gpa_n_sort = df_n.sort_values(by='GPA', ascending=False)
gpa_n_sort.to_excel(output_excel_file_n, index=False)
print("Sorted N")

import pandas as pd
import random
import openpyxl
import os

script_dir=os.path.dirname(os.path.abspath('StudentSort.py'))
# Load the remaining students file
inputfile=os.path.join(script_dir,'..','Residuals')
remainingStudents=os.path.join(inputfile,'RemainingStudents.xlsx')
remaining_students_df = pd.read_excel(remainingStudents)
 # Update with the path to your main Excel file
# Load the assigned projects file
assignedStudents=os.path.join(inputfile,'assigned_projects.xlsx')
assigned_projects_df = pd.read_excel(assignedStudents)

# Loop through the rows in assigned projects
for idx, row in assigned_projects_df.iterrows():
    for i in range(1, 9):  # Assuming you have 8 student columns (Assigned Student 1 to Assigned Student 8)
        if pd.isna(row[f'Assigned Student {i}']):
            major = row[f'major_{i}']
            assigned_student_col = f'Assigned Student {i}'
            
            # Check if major is not null
            if pd.notna(major):
                # Search for students with the same major and GPA between 3.5 and 2.9
                matching_students = remaining_students_df[
                    (remaining_students_df['MJR'] == major) &
                    (remaining_students_df['GPA'] >= 2.9) &  # Minimum GPA
                    (remaining_students_df['GPA'] <= 3.5)     # Maximum GPA
                ]
                
                if not matching_students.empty:
                    random_student_idx = random.choice(matching_students.index)
                    random_student_info = matching_students.loc[random_student_idx, 'Student Info']
                    assigned_projects_df.at[idx, assigned_student_col] = random_student_info
                    remaining_students_df.drop(random_student_idx, inplace=True)
                elif not remaining_students_df.empty:
                    # If no matching student with the same major and GPA range, assign a random student
                    random_student_idx = random.choice(remaining_students_df.index)
                    random_student_info = remaining_students_df.loc[random_student_idx, 'Student Info']
                    assigned_projects_df.at[idx, assigned_student_col] = random_student_info
                    remaining_students_df.drop(random_student_idx, inplace=True)


output_folder = os.path.join(script_dir,'..','Output_Files')# Update with the path to your main Excel file

output_path=os.path.join(output_folder,'FinalOutput.xlsx')
existing_data_df=pd.read_excel(output_path)
final_data_df = pd.concat([existing_data_df, assigned_projects_df], ignore_index=True)
# Save the updated assigned projects DataFrame to the specified folder
final_data_df.to_excel(output_path, index=False)


with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    # Write the remaining students data to a new sheet
    remaining_students_df.to_excel(writer, sheet_name='RemainingStudents', index=False)



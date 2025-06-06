import pandas as pd
import random
import os

script_dir=os.path.dirname(os.path.abspath('StudentSort.py'))
inputpath=os.path.join(script_dir,'..','Residuals')
studentsN=os.path.join(inputpath,'Student N Citizenship_.xlsx')
student_n_citizenship_df = pd.read_excel(studentsN)
# Load the remaining students file
remainingStudentsC=os.path.join(inputpath,'RemainingStudentsC.xlsx')
remaining_students_df = pd.read_excel(remainingStudentsC)

# Load the assigned projects file
assignedp=os.path.join(inputpath,'assigned_projects.xlsx')
assigned_projects_df = pd.read_excel(assignedp)

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
                else:
                    # If no matching student with the same major and GPA range, assign a random student
                    random_student_idx = random.choice(remaining_students_df.index)
                    random_student_info = remaining_students_df.loc[random_student_idx, 'Student Info']
                    assigned_projects_df.at[idx, assigned_student_col] = random_student_info
                    remaining_students_df.drop(random_student_idx, inplace=True)

output_folder = os.path.join(script_dir,'..','Output_Files') # Update with the path to your main Excel file

combined_df = pd.concat([student_n_citizenship_df, remaining_students_df], ignore_index=True)
combined_df.to_excel(studentsN, index=False)
# Save the updated assigned projects DataFrame to the specified folder
output_file_path = os.path.join(output_folder,'FinalOutput.xlsx')
assigned_projects_df.to_excel(output_file_path, index=False)

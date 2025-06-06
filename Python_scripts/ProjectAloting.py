import pandas as pd
import os


script_dir=os.path.dirname(os.path.abspath('ProjectAloting.py'))
residual_folder= os.path.join(script_dir, '..', 'Residuals')
studentsCpath=os.path.join (residual_folder,'Student C Citizenship_.xlsx')
projectsCpath=os.path.join(residual_folder,'Projects for C Citizens.xlsx')
studentsNpath=os.path.join(residual_folder,'Student N Citizenship_.xlsx')
# Read the students and projects files into DataFrames
students_df =   pd.read_excel(studentsCpath)
projects_df =   pd.read_excel(projectsCpath)
Noncitizen_df = pd.read_excel(studentsNpath)

# Create a new DataFrame to store the assigned projects with student names
assigned_projects_df = projects_df.copy()

# Function to get student IDs with a specific project choice
def students_with_project_choice(project_id):
    # Search the project ID in the top 15 choices of students
    choices_columns = ['Choice 1', 'Choice 2', 'Choice 3', 'Choice 4', 'Choice 5',
                       'Choice 6', 'Choice 7', 'Choice 8', 'Choice 9', 'Choice 10',
                       'Choice 11', 'Choice 12', 'Choice 13', 'Choice 14', 'Choice 15']
    
    # Create a list to store student IDs with the project ID in their choices
    student_ids = []
    
    for col in choices_columns:
        students_with_choice = students_df[students_df[col] == project_id]
        student_ids.extend(students_with_choice['ID'].tolist())
    
    # Remove duplicates and return the list of student IDs
    return list(set(student_ids))

# Loop through each row of projects DataFrame
for idx, project_row in projects_df.iterrows():
    # Get the majors from the current row
    majors = [project_row[f'major_{i}'] for i in range(1, 9)]  # Assuming major columns are named 'major_1', 'major_2', ..., 'major_8'
    project_id = project_row['Project Title']

    # Create a list to store assigned student names for each major
    assigned_students = []

    num_non_empty_majors = sum([1 for major in majors if pd.notna(major)])
    
    if num_non_empty_majors == 1:
        major_1 = majors[0]  # Assuming major_1 is the first non-empty major
        print("Checking for Major:", major_1)  # Print the major you're checking
        eligible_students = students_df[students_df['MJR'] == major_1]
    
        # Filter students with GPA > 3.5 and having the project in their top 5 choices
        eligible_students_3_5_above = eligible_students[
        (eligible_students['GPA'] > 3.5) &
        (
            (eligible_students['Choice 1'] == project_id) |
            (eligible_students['Choice 2'] == project_id) |
            (eligible_students['Choice 3'] == project_id) |
            (eligible_students['Choice 4'] == project_id) |
            (eligible_students['Choice 5'] == project_id)
        )
    ]
    
        if not eligible_students_3_5_above.empty:
        # Assign from the top 5 choices
            student = eligible_students_3_5_above.iloc[0]
            assigned_students.append(student['Student Info'])
            students_df = students_df[students_df['ID'] != student['ID']]
        else:
        # If not found in top 5 choices, check next 5 choices
            eligible_students_3_5_above = eligible_students[
            (eligible_students['GPA'] > 3.5) &
            (
                (eligible_students['Choice 6'] == project_id) |
                (eligible_students['Choice 7'] == project_id) |
                (eligible_students['Choice 8'] == project_id) |
                (eligible_students['Choice 9'] == project_id) |
                (eligible_students['Choice 10'] == project_id)
            )
        ]
        
        if not eligible_students_3_5_above.empty:
            # Assign from the next 5 choices
            student = eligible_students_3_5_above.iloc[0]
            assigned_students.append(student['Student Info'])
            students_df = students_df[students_df['ID'] != student['ID']]
        else:
            # If not found in next 5 choices, check all 15 choices
            eligible_students_3_5_above = eligible_students[
                (eligible_students['GPA'] > 3.5) &
                (
                    (eligible_students['Choice 1'] == project_id) |
                    (eligible_students['Choice 2'] == project_id) |
                    (eligible_students['Choice 3'] == project_id) |
                    (eligible_students['Choice 4'] == project_id) |
                    (eligible_students['Choice 5'] == project_id) |
                    (eligible_students['Choice 6'] == project_id) |
                    (eligible_students['Choice 7'] == project_id) |
                    (eligible_students['Choice 8'] == project_id) |
                    (eligible_students['Choice 9'] == project_id) |
                    (eligible_students['Choice 10'] == project_id) |
                    (eligible_students['Choice 11'] == project_id) |
                    (eligible_students['Choice 12'] == project_id) |
                    (eligible_students['Choice 13'] == project_id) |
                    (eligible_students['Choice 14'] == project_id) |
                    (eligible_students['Choice 15'] == project_id)
                )
            ]
            
            if not eligible_students_3_5_above.empty:
                # Assign from all 15 choices
                student = eligible_students_3_5_above.iloc[0]
                assigned_students.append(student['Student Info'])
                students_df = students_df[students_df['ID'] != student['ID']]
            else:
                if len(assigned_students) == 0:
                    eligible_students_3_5_to_2_9 = eligible_students[
                    (3.5 > eligible_students['GPA']) & (eligible_students['GPA'] >= 2.9)
        ]
                    eligible_students_3_5_to_2_9 = eligible_students_3_5_to_2_9[
                    (
                                (eligible_students_3_5_to_2_9['Choice 1'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 2'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 3'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 4'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 5'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 6'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 7'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 8'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 9'] == project_id) |
                                (eligible_students_3_5_to_2_9['Choice 10'] == project_id)
                            )
                        ]
                if not eligible_students_3_5_to_2_9.empty:
                            # Assign from the top 5 choices
                            student = eligible_students_3_5_to_2_9.iloc[0]
                            assigned_students.append(student['Student Info'])
                            students_df = students_df[students_df['ID'] != student['ID']]
                else:
                            # If not found in top 5 choices, check next 5 choices
                            eligible_students_3_5_to_2_9 = eligible_students[
                                (3.5 > eligible_students['GPA']) & (eligible_students['GPA'] >= 2.9)
                            ]
                            eligible_students_3_5_to_2_9 = eligible_students_3_5_to_2_9[
                                    (
                                        (eligible_students_3_5_to_2_9['Choice 1'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 2'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 3'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 4'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 5'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 6'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 7'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 8'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 9'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 10'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 11'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 12'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 13'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 14'] == project_id) |
                                        (eligible_students_3_5_to_2_9['Choice 15'] == project_id)
                                    )
                                ]
                
                if not eligible_students_3_5_to_2_9.empty:
                            # Assign from the top 5 choices
                            student = eligible_students_3_5_to_2_9.iloc[0]
                            assigned_students.append(student['Student Info'])
                            students_df = students_df[students_df['ID'] != student['ID']]
        # Similar logic as above for this GPA category
        # Check top 10 choices, then all 15 choices if needed
        # ...
            
    # If not found in the GPA > 3.5 and 3.5 to 2.9 categories, check the below 2.9 GPA category
        if len(assigned_students) == 0:
            eligible_students_below_2_9 = eligible_students[
            eligible_students['GPA'] < 2.9]
            eligible_students_3_5_to_2_9 = eligible_students_3_5_to_2_9[
                                    (
                                        (eligible_students_below_2_9 ['Choice 1'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 2'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 3'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 4'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 5'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 6'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 7'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 8'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 9'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 10'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 11'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 12'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 13'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 14'] == project_id) |
                                        (eligible_students_below_2_9 ['Choice 15'] == project_id)
                                    )
                                ]
            if not eligible_students_below_2_9.empty:
                # Assign from all 15 choices
                student = eligible_students_below_2_9.iloc[0]
                assigned_students.append(student['Student Info'])
                students_df = students_df[students_df['ID'] != student['ID']]



# ############################################################################################################################








    elif num_non_empty_majors == 2:
        major_1 = majors[0]
        major_2 = majors[1]
        print("Checking for Majors:", major_1, major_2)
        
        eligible_students_major_1 = students_df[
            (students_df['MJR'] == major_1) & 
            (students_df['GPA'] > 3.5) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) 
            )
        ]
        
        eligible_students_major_2 = students_df[
            (students_df['MJR'] == major_2) & 
            (students_df['GPA'] <= 3.5) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) |
                (students_df['Choice 6']==project_id)|
                (students_df['Choice 7']==project_id)|
                (students_df['Choice 8']==project_id)|
                (students_df['Choice 9']==project_id)|
                (students_df['Choice 10']==project_id)
            )
        ]
        
        if not eligible_students_major_1.empty:
            student_major_1 = eligible_students_major_1.iloc[0]
            assigned_students.append(student_major_1['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_1['ID']]
            
        if not eligible_students_major_2.empty:
            student_major_2 = eligible_students_major_2.iloc[0]
            assigned_students.append(student_major_2['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_2['ID']]






            #######################################################################################################








    elif num_non_empty_majors == 3:
        major_1 = majors[0]
        major_2 = majors[1]
        major_3 = majors[2]
        print("Checking for Majors:", major_1, major_2, major_3)
        
        eligible_students_major_1 = students_df[
            (students_df['MJR'] == major_1) & 
            (students_df['GPA'] > 3.5) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) 
            )
        ]
        
        eligible_students_major_2 = students_df[
            (students_df['MJR'] == major_2) & 
            (3.5 >= students_df['GPA']) & 
            (students_df['GPA'] >= 2.9) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) |
                (students_df['Choice 6'] == project_id) |
                (students_df['Choice 7'] == project_id) |
                (students_df['Choice 8'] == project_id) |
                (students_df['Choice 9'] == project_id) |
                (students_df['Choice 10'] == project_id) 
            )
        ]
        
        eligible_students_major_3 = students_df[
            (students_df['MJR'] == major_3) & 
            (students_df['GPA'] < 2.9) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) |
                (students_df['Choice 6'] == project_id) |
                (students_df['Choice 7'] == project_id) |
                (students_df['Choice 8'] == project_id) |
                (students_df['Choice 9'] == project_id) |
                (students_df['Choice 10'] == project_id) |
                (students_df['Choice 11'] == project_id) |
                (students_df['Choice 12'] == project_id) |
                (students_df['Choice 13'] == project_id) |
                (students_df['Choice 14'] == project_id) |
                (students_df['Choice 15'] == project_id)
            )
        ]
        
        if not eligible_students_major_1.empty:
            student_major_1 = eligible_students_major_1.iloc[0]
            assigned_students.append(student_major_1['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_1['ID']]
        else:
            assigned_students.append("NA")
                
            
        if not eligible_students_major_2.empty:
            student_major_2 = eligible_students_major_2.iloc[0]
            assigned_students.append(student_major_2['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_2['ID']]
        else:
            assigned_students.append("NA")
            
        if not eligible_students_major_3.empty:
            student_major_3 = eligible_students_major_3.iloc[0]
            assigned_students.append(student_major_3['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_3['ID']]
        else:
            assigned_students.append("NA")



################################################################################################################################



 
 
 
 
 
 
    elif num_non_empty_majors == 4:
        major_1 = majors[0]
        major_2 = majors[1]
        major_3 = majors[2]
        major_4 = majors[3]
        print("Checking for Majors:", major_1, major_2, major_3, major_4)
        
        eligible_students_major_1 = students_df[
            (students_df['MJR'] == major_1) & 
            (students_df['GPA'] > 3.5) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) 
            )
        ]
        
        eligible_students_major_2 = students_df[
            (students_df['MJR'] == major_2) & 
            (3.5 >= students_df['GPA']) & 
            (students_df['GPA'] >= 2.9) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) |
                (students_df['Choice 6'] == project_id) |
                (students_df['Choice 7'] == project_id) |
                (students_df['Choice 8'] == project_id) |
                (students_df['Choice 9'] == project_id) |
                (students_df['Choice 10'] == project_id) 
            )
        ]
        if not eligible_students_major_2.empty:
            student_major_2 = eligible_students_major_2.iloc[0]
            assigned_students.append(student_major_2['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_2['ID']]
        else:
            assigned_students.append("NA")
        eligible_students_major_3 = students_df[
            (students_df['MJR'] == major_3) & 
            (3.5 >= students_df['GPA']) & 
            (students_df['GPA'] >= 2.9) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) |
                (students_df['Choice 6'] == project_id) |
                (students_df['Choice 7'] == project_id) |
                (students_df['Choice 8'] == project_id) |
                (students_df['Choice 9'] == project_id) |
                (students_df['Choice 10'] == project_id) 
            )
        ]
        
        eligible_students_major_4 = students_df[
            (students_df['MJR'] == major_4) & 
            (students_df['GPA'] < 2.9) & 
            (
                (students_df['Choice 1'] == project_id) |
                (students_df['Choice 2'] == project_id) |
                (students_df['Choice 3'] == project_id) |
                (students_df['Choice 4'] == project_id) |
                (students_df['Choice 5'] == project_id) |
                (students_df['Choice 6'] == project_id) |
                (students_df['Choice 7'] == project_id) |
                (students_df['Choice 8'] == project_id) |
                (students_df['Choice 9'] == project_id) |
                (students_df['Choice 10'] == project_id) |
                (students_df['Choice 11'] == project_id) |
                (students_df['Choice 12'] == project_id) |
                (students_df['Choice 13'] == project_id) |
                (students_df['Choice 14'] == project_id) |
                (students_df['Choice 15'] == project_id)
            )
        ]
        
        if not eligible_students_major_1.empty:
            student_major_1 = eligible_students_major_1.iloc[0]
            assigned_students.append(student_major_1['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_1['ID']]
        else:
            assigned_students.append("NA")
            
        
            
        if not eligible_students_major_3.empty:
            student_major_3 = eligible_students_major_3.iloc[0]
            assigned_students.append(student_major_3['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_3['ID']]
        else:
            assigned_students.append("NA")
            
        if not eligible_students_major_4.empty:
            student_major_4 = eligible_students_major_4.iloc[0]
            assigned_students.append(student_major_4['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_4['ID']]
        else:
            assigned_students.append("NA")



########################################################################################################






    elif num_non_empty_majors == 5:
        major_1 = majors[0]
        major_2 = majors[1]
        major_3 = majors[2]
        major_4 = majors[3]
        major_5 = majors[4]
        print("Checking for Majors:", major_1, major_2, major_3, major_4, major_5)
    # Logic for student 1 with major_1 and GPA > 3.5
        eligible_students_major_1 = students_df[
        (students_df['MJR'] == major_1) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id)  |
            (students_df['Choice 2'] == project_id)  |
            (students_df['Choice 3'] == project_id)  |
            (students_df['Choice 4'] == project_id)  |
            (students_df['Choice 5'] == project_id)  
        )
    ]

        if not eligible_students_major_1.empty:
            student_major_1 = eligible_students_major_1.iloc[0]
            assigned_students.append(student_major_1['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_1['ID']]
        else:
            assigned_students.append("NA")
         

    # Logic for student 2 with major_2 and GPA > 3.5
        eligible_students_major_2 = students_df[
        (students_df['MJR'] == major_2) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) 
        )
    ]

        if not eligible_students_major_2.empty:
            student_major_2 = eligible_students_major_2.iloc[0]
            assigned_students.append(student_major_2['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_2['ID']]

    # Logic for student 3 with major_3 and GPA between 3.5 and 2.9
        eligible_students_major_3 = students_df[
        (students_df['MJR'] == major_3) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_3.empty:
            student_major_3 = eligible_students_major_3.iloc[0]
            assigned_students.append(student_major_3['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_3['ID']]
        else:
            assigned_students.append("NA")

    # Logic for student 4 with major_4 and GPA below 2.9
        eligible_students_major_4 = students_df[
        (students_df['MJR'] == major_4) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_4.empty:
            student_major_4 = eligible_students_major_4.iloc[0]
            assigned_students.append(student_major_4['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_4['ID']]
        else:
            assigned_students.append("NA")

    # Logic for student 5 with major_5 and GPA below 2.9
        eligible_students_major_5 = students_df[
        (students_df['MJR'] == major_5) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_5.empty:
            student_major_5 = eligible_students_major_5.iloc[0]
            assigned_students.append(student_major_5['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_5['ID']]
        else:
            assigned_students.append("NA")

######################################################################################################################################
   
   
   
   
    elif num_non_empty_majors == 6:
        major_1 = majors[0]
        major_2 = majors[1]
        major_3 = majors[2]
        major_4 = majors[3]
        major_5 = majors[4]
        major_6=majors[5]
        print("Checking for Majors:", major_1, major_2, major_3, major_4, major_5, major_6)

    # Logic for student 1 with major_1 and GPA > 3.5
        eligible_students_major_1 = students_df[
        (students_df['MJR'] == major_1) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_1.empty:
            student_major_1 = eligible_students_major_1.iloc[0]
            assigned_students.append(student_major_1['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_1['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 2 with major_2 and GPA > 3.5
        eligible_students_major_2 = students_df[
        (students_df['MJR'] == major_2) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) 
        )
    ]

        if not eligible_students_major_2.empty:
            student_major_2 = eligible_students_major_2.iloc[0]
            assigned_students.append(student_major_2['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_2['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 3 with major_3 and GPA between 3.5 and 2.9
        eligible_students_major_3 = students_df[
        (students_df['MJR'] == major_3) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_3.empty:
            student_major_3 = eligible_students_major_3.iloc[0]
            assigned_students.append(student_major_3['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_3['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 4 with major_4 and GPA between 3.5 and 2.9
        eligible_students_major_4 = students_df[
        (students_df['MJR'] == major_4) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_4.empty:
            student_major_4 = eligible_students_major_4.iloc[0]
            assigned_students.append(student_major_4['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_4['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 5 with major_5 and GPA below 2.9
        eligible_students_major_5 = students_df[
        (students_df['MJR'] == major_5) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_5.empty:
            student_major_5 = eligible_students_major_5.iloc[0]
            assigned_students.append(student_major_5['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_5['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 6 with major_6 and GPA below 2.9
        eligible_students_major_6 = students_df[
        (students_df['MJR'] == major_6) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_6.empty:
            student_major_6 = eligible_students_major_6.iloc[0]
            assigned_students.append(student_major_6['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_6['ID']]
        else:
            assigned_students.append("NA")

    #######################################################################################################################






    elif num_non_empty_majors == 7:
        major_1 = majors[0]
        major_2 = majors[1]
        major_3 = majors[2]
        major_4 = majors[3]
        major_5 = majors[4]
        major_6 = majors[5]
        major_7 = majors[6]
        print("Checking for Majors:", major_1, major_2, major_3, major_4, major_5, major_6, major_7)

    # Logic for student 1 with major_1 and GPA > 3.5
        eligible_students_major_1 = students_df[
        (students_df['MJR'] == major_1) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) 
        )
    ]

        if not eligible_students_major_1.empty:
            student_major_1 = eligible_students_major_1.iloc[0]
            assigned_students.append(student_major_1['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_1['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 2 with major_2 and GPA > 3.5
        eligible_students_major_2 = students_df[
        (students_df['MJR'] == major_2) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) 
        )
    ]

        if not eligible_students_major_2.empty:
            student_major_2 = eligible_students_major_2.iloc[0]
            assigned_students.append(student_major_2['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_2['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 3 with major_3 and GPA between 3.5 and 2.9
        eligible_students_major_3 = students_df[
        (students_df['MJR'] == major_3) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_3.empty:
            student_major_3 = eligible_students_major_3.iloc[0]
            assigned_students.append(student_major_3['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_3['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 4 with major_4 and GPA between 3.5 and 2.9
        eligible_students_major_4 = students_df[
        (students_df['MJR'] == major_4) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_4.empty:
            student_major_4 = eligible_students_major_4.iloc[0]
            assigned_students.append(student_major_4['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_4['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 5 with major_5 and GPA between 3.5 and 2.9
        eligible_students_major_5 = students_df[
        (students_df['MJR'] == major_5) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_5.empty:
            student_major_5 = eligible_students_major_5.iloc[0]
            assigned_students.append(student_major_5['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_5['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 6 with major_6 and GPA below 2.9
        eligible_students_major_6 = students_df[
        (students_df['MJR'] == major_6) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_6.empty:
            student_major_6 = eligible_students_major_6.iloc[0]
            assigned_students.append(student_major_6['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_6['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 7 with major_7 and GPA below 2.9
        eligible_students_major_7 = students_df[
        (students_df['MJR'] == major_7) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_7.empty:
            student_major_7 = eligible_students_major_7.iloc[0]
            assigned_students.append(student_major_7['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_7['ID']]
        else:
            assigned_students.append("NA")


###########################################################################################################################














    elif num_non_empty_majors == 8:
        major_1 = majors[0]
        major_2 = majors[1]
        major_3 = majors[2]
        major_4 = majors[3]
        major_5 = majors[4]
        major_6 = majors[5]
        major_7 = majors[6]
        major_8 = majors[7]
        print("Checking for Majors:", major_1, major_2, major_3, major_4, major_5, major_6, major_7,major_8)

    # Logic for student 1 with major_1 and GPA > 3.5
        eligible_students_major_1 = students_df[
        (students_df['MJR'] == major_1) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) 
        )
    ]

        if not eligible_students_major_1.empty:
            student_major_1 = eligible_students_major_1.iloc[0]
            assigned_students.append(student_major_1['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_1['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 2 with major_2 and GPA > 3.5
        eligible_students_major_2 = students_df[
        (students_df['MJR'] == major_2) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) 
        )
    ]

        if not eligible_students_major_2.empty:
            student_major_2 = eligible_students_major_2.iloc[0]
            assigned_students.append(student_major_2['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_2['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 3 with major_3 and GPA between 3.5 and 2.9
        eligible_students_major_3 = students_df[
        (students_df['MJR'] == major_3) &
        (students_df['GPA'] > 3.5) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) 
        )
    ]

        if not eligible_students_major_3.empty:
            student_major_2 = eligible_students_major_3.iloc[0]
            assigned_students.append(student_major_3['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_3['ID']]
        else:
            assigned_students.append("NA")
    
        eligible_students_major_4 = students_df[
        (students_df['MJR'] == major_4) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_4.empty:
            student_major_4 = eligible_students_major_4.iloc[0]
            assigned_students.append(student_major_4['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_4['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 4 with major_4 and GPA between 3.5 and 2.9
        eligible_students_major_5 = students_df[
        (students_df['MJR'] == major_5) &
        (3.5 > students_df['GPA']) & (students_df['GPA'] >= 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) 
        )
    ]

        if not eligible_students_major_5.empty:
            student_major_5 = eligible_students_major_5.iloc[0]
            assigned_students.append(student_major_5['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_5['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 5 with major_5 and GPA between 3.5 and 2.9
    # Logic for student 6 with major_6 and GPA below 2.9
        eligible_students_major_6 = students_df[
        (students_df['MJR'] == major_6) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_6.empty:
            student_major_6 = eligible_students_major_6.iloc[0]
            assigned_students.append(student_major_6['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_6['ID']]
        else:
            assigned_students.append("NA")
    # Logic for student 7 with major_7 and GPA below 2.9
        eligible_students_major_7 = students_df[
        (students_df['MJR'] == major_7) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_7.empty:
            student_major_7 = eligible_students_major_7.iloc[0]
            assigned_students.append(student_major_7['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_7['ID']]
        else:
            assigned_students.append("NA")
        eligible_students_major_8 = students_df[
        (students_df['MJR'] == major_8) &
        (students_df['GPA'] < 2.9) &
        (
            (students_df['Choice 1'] == project_id) |
            (students_df['Choice 2'] == project_id) |
            (students_df['Choice 3'] == project_id) |
            (students_df['Choice 4'] == project_id) |
            (students_df['Choice 5'] == project_id) |
            (students_df['Choice 6'] == project_id) |
            (students_df['Choice 7'] == project_id) |
            (students_df['Choice 8'] == project_id) |
            (students_df['Choice 9'] == project_id) |
            (students_df['Choice 10'] == project_id) |
            (students_df['Choice 11'] == project_id) |
            (students_df['Choice 12'] == project_id) |
            (students_df['Choice 13'] == project_id) |
            (students_df['Choice 14'] == project_id) |
            (students_df['Choice 15'] == project_id)
        )
    ]

        if not eligible_students_major_8.empty:
            student_major_8 = eligible_students_major_8.iloc[0]
            assigned_students.append(student_major_8['Student Info'])
            students_df = students_df[students_df['ID'] != student_major_8['ID']]
        else:
            assigned_students.append("NA")

    # Update the current row of the assigned projects DataFrame with assigned student names
    assigned_projects_df.at[idx, 'Assigned Student 1'] = assigned_students[0] if len(assigned_students) > 0 else ""
    assigned_projects_df.at[idx, 'Assigned Student 2'] = assigned_students[1] if len(assigned_students) > 1 else ""
    assigned_projects_df.at[idx, 'Assigned Student 3'] = assigned_students[2] if len(assigned_students) > 2 else ""
    assigned_projects_df.at[idx, 'Assigned Student 4'] = assigned_students[3] if len(assigned_students) > 3 else ""
    assigned_projects_df.at[idx, 'Assigned Student 5'] = assigned_students[4] if len(assigned_students) > 4 else ""
    assigned_projects_df.at[idx, 'Assigned Student 6'] = assigned_students[5] if len(assigned_students) > 5 else ""
    assigned_projects_df.at[idx, 'Assigned Student 7'] = assigned_students[6] if len(assigned_students) > 6 else ""
    assigned_projects_df.at[idx, 'Assigned Student 8'] = assigned_students[7] if len(assigned_students) > 7 else ""
# Save the assigned projects DataFrame to a new Excel file


outputAssignedpath=os.path.join(residual_folder,'assigned_projects.xlsx')
assigned_projects_df.to_excel(outputAssignedpath, index=False)


# # Save the updated student DataFrame to a new Excel file
outputRemainingstudent=os.path.join(residual_folder,'RemainingStudentsC.xlsx')
students_df.to_excel(outputRemainingstudent, index=False)
print("Students assigned to projects based on major choices and GPA criteria. Assigned projects saved to 'assigned_projects.xlsx' and updated students saved to 'Student N Citizenship_.xlsx'.")
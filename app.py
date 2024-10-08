from flask import Flask, request, render_template, redirect, url_for,session
import os
import pandas as pd
import shutil
import time






app = Flask(__name__)
# Generate a random and secure secret key
app.secret_key = os.urandom(32) # 16 bytes = 32 characters


# Function to read Excel file with retry mechanism
def read_excel_with_retry(file_path):
    df = None
    while True:
        try:
            df = pd.read_excel(file_path)

        except Exception as e:
            print(f"Error reading Excel file: {e}")
            time.sleep(0.1)
            # Handle the error (e.g., log, retry, or break out of the loop)
            pass
        else:
            break
    return df

# Function to write DataFrame to Excel file with retry mechanism
def write_excel_with_retry(df, file_path):
    while True:
        try:
            df.to_excel(file_path, index=False)
            time.sleep(0.1)
        except Exception as e:
            print(f"Error writing to Excel file: {e}")
            # Handle the error (e.g., log, retry, or break out of the loop)
            pass
        else:
            break


# Function to read CSV file with retry mechanism
def read_csv_with_retry(file_path):
    df = None
    while True:
        try:
            df = pd.read_csv(file_path)
            time.sleep(0.1)
        except Exception as e:
            print(f"Error reading CSV file: {e}")
            # Handle the error (e.g., log, retry, or break out of the loop)
            pass
        else:
            break
    return df

# Function to write DataFrame to CSV file with retry mechanism
def write_csv_with_retry(df, file_path):
    while True:
        try:
            df.to_csv(file_path, index=False)
            time.sleep(0.1)
        except Exception as e:
            print(f"Error writing to CSV file: {e}")
            # Handle the error (e.g., log, retry, or break out of the loop)
            pass
        else:
            break

# Function to read the current value from 'config.txt' file with retry mechanism
def read_config_file():
    current_value = ""
    while True:
        try:
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()
            time.sleep(0.1)
        except Exception as e:
            print(f"Error reading config.txt: {e}")
            # Handle the error (e.g., log, retry, or break out of the loop)
            pass
        else:
            break
    return current_value

# Function to write the selected folder to 'config.txt' file with retry mechanism
def write_config_file(selected_folder):
    while True:
        try:
            with open('config.txt', 'w') as file:
                file.write(selected_folder)
            time.sleep(0.1)
        except Exception as e:
            print(f"Error writing config.txt: {e}")
            # Handle the error (e.g., log, retry, or break out of the loop)
            pass
        else:
            break

# Function to read folder names from 'db' directory with retry mechanism
def read_folder_names():
    db_folder_path = 'db'
    folders = []
    while True:
        try:
            folders = [folder for folder in os.listdir(db_folder_path) if os.path.isdir(os.path.join(db_folder_path, folder))]
            time.sleep(0.1)
        except Exception as e:
            print(f"Error reading folders: {e}")
            # Handle the error (e.g., log, retry, or break out of the loop)
            pass
        else:
            break
    return folders


@app.route('/',methods=['GET', 'POST'])
def welcome():
    alert_message = None
    session['admin'] = False
    if request.method == 'POST':
        username = request.form['username'].upper()
        password = request.form['password'].upper()
        user_type = request.form.get('user_type')
        try:
            if user_type == "admins":
                # Read Excel file with retry mechanism
                file_path = "./static/admin/adminlogin.xlsx"
                df = read_excel_with_retry(file_path)

                user_exists = ((df['admin_id'] == username) & (df['admin_password'] == password)).any()
                df = None
                if user_exists:
                    session['adminname'] = username
                    session['admin'] = True
                    return redirect(url_for('admin_ui'))
                else:
                    alert_message = "Incorrect username or password"
        except Exception as e:
            print("Error:", e)
            alert_message = "Error occurred"
    return render_template('welcome/welcome.html')

@app.route('/home',methods=['GET', 'POST'])
def home():
    alert_message = None
    session['admin'] = False
    if request.method == 'POST':
        username = request.form['username'].upper()
        password = request.form['password'].upper()
        user_type = request.form.get('user_type')
        try:
            if user_type == "admins":
                # Read Excel file with retry mechanism
                file_path = "./static/admin/adminlogin.xlsx"
                df = read_excel_with_retry(file_path)

                user_exists = ((df['admin_id'] == username) & (df['admin_password'] == password)).any()
                df = None
                if user_exists:
                    session['adminname'] = username
                    session['admin'] = True
                    return redirect(url_for('admin_ui'))
                else:
                    alert_message = "Incorrect username or password"
        except Exception as e:
            print("Error:", e)
            alert_message = "Error occurred"
    return render_template('welcome/welcome.html')


@app.route('/about')
def about():
    return render_template('commonpages/about.html')

@app.route('/404_not_found')
def notfound():
    session['admin'] = False
    return render_template('commonpages/notfound.html')




# Admin UI route
@app.route('/admin_ui', methods=['GET', 'POST'])
def admin_ui():
    username = session.get('adminname')
    admin = session.get('admin')
    if admin:
        db_folder_path = 'db'
        folders = []
        current_value = ""


        for attempt in range(1, 5000):
            try:
                with open('config.txt', 'r') as file:
                    current_value = file.read().strip()
                time.sleep(0.1)
            except Exception as e:
                print(f"Error reading config.txt (Attempt {attempt}/{max_attempts}): {e}")
                # Handle the error (e.g., log, retry, or break out of the loop)
                pass
            else:
                break
        else:
            print("Reached maximum attempts. Could not read config.txt.")


        # Read folder names from the 'db' directory with retry mechanism
        while True:
            try:
                folders = [folder for folder in os.listdir(db_folder_path) if os.path.isdir(os.path.join(db_folder_path, folder))]
                time.sleep(0.1)
            except Exception as e:
                print(f"Error reading folders: {e}")
                # Handle the error (e.g., log, retry, or break out of the loop)
                pass
            else:
                break

        if request.method == 'POST':
            selected_folder = request.form.get('selected_folder')
            if selected_folder is not None:
                # Update the 'config.txt' file with the selected folder with retry mechanism
                while True:
                    try:
                        with open('config.txt', 'w') as file:
                            file.write(selected_folder)
                        time.sleep(0.1)
                    except Exception as e:
                        print(f"Error writing config.txt: {e}")
                        # Handle the error (e.g., log, retry, or break out of the loop)
                        pass
                    else:
                        break
                while True:
                    try:
                        with open('config.txt', 'r') as file:
                            current_value = file.read().strip()
                        time.sleep(0.1)
                    except Exception as e:
                        print(f"Error reading config.txt: {e}")
                        # Handle the error (e.g., log, retry, or break out of the loop)
                        pass
                    else:
                        break
                return render_template('adminpages/admin_ui.html', folders=folders, message=f'Current default value: {current_value}')
            else:
                # Handle the case when 'selected_folder' is None (not selected)
                return render_template('adminpages/admin_ui.html', folders=folders, message=f'Current default value: {current_value}')

        # If it's a GET request
        return render_template('adminpages/admin_ui.html', folders=folders, message=f'Current default value: {current_value}')
    else:
        return render_template('commonpages/notfound.html')


# Admin Management route
@app.route('/admin_management', methods=['GET', 'POST'])
def admin_management():
    admin = session.get('admin')
    if admin:
        file_path = "./static/admin/adminlogin.xlsx"
        df = read_excel_with_retry(file_path)

        if df is None:
            # Set a message for the template
            return render_template('adminpages/admin_management.html', file_not_found_message='File not found. No data to display.')

        if request.method == 'POST':
            try:
                if 'insert' in request.form:
                    # Get data from the form for inserting
                    adminid = request.form.get('adminid').upper()
                    password = request.form.get('password').upper()

                    # Check if the admin ID is unique before inserting
                    if adminid not in df['admin_id'].values:
                        # Append new data to the DataFrame
                        new_data = pd.DataFrame({'admin_id': [adminid], 'admin_password': [password]})
                        df = pd.concat([df, new_data], ignore_index=True)

                        # Write the updated DataFrame to the Excel file
                        write_excel_with_retry(df, file_path)
                        df = df.sort_values(by=['admin_id'])

                        # Set a success message
                        success_message = f"Admin ID '{adminid}' inserted successfully!"
                        return render_template('adminpages/admin_management.html', Admins=df.to_dict('records'), success_message=success_message)
                    else:
                        # Set an error message for duplicate admin ID
                        error_message = "Admin ID already exists. Please choose a unique ID."
                        return render_template('adminpages/admin_management.html', Admins=df.to_dict('records'), error_message=error_message)

                elif 'delete' in request.form:
                    # Get the user ID for deleting
                    userid_to_delete = request.form.get('adminid').upper()

                    if len(df) <= 1:
                        error_message = "Cannot delete the last admin record."
                        return render_template('adminpages/admin_management.html', Admins=df.to_dict('records'), error_message=error_message)

                    # Check if the admin ID exists in the DataFrame
                    elif userid_to_delete in df['admin_id'].values:
                        # Remove the user with the specified ID
                        df = df[df['admin_id'] != userid_to_delete]

                        # Write the updated DataFrame to the Excel file
                        write_excel_with_retry(df, file_path)
                        df = df.sort_values(by=['admin_id'])
                        # Set a success message
                        success_message = f"Admin ID '{userid_to_delete}' deleted successfully!"
                        return render_template('adminpages/admin_management.html', Admins=df.to_dict('records'), success_message=success_message)
                    else:
                        # Set an error message for non-existent admin ID
                        error_message = f"Admin ID '{userid_to_delete}' does not exist."
                        return render_template('adminpages/admin_management.html', Admins=df.to_dict('records'), error_message=error_message)

            except Exception as e:
                print("Error:", e)
                # Set an error message
                error_message = 'An error occurred.'
                return render_template('adminpages/admin_management.html', Admins=df.to_dict('records'), error_message=error_message)

        return render_template('adminpages/admin_management.html', Admins=df.to_dict('records'))



# Admin New Academic Year route
@app.route('/admin_new_academic_year',methods=['GET', 'POST'])
def admin_new_academic_year():
    admin = session.get('admin')
    if admin:
        db_folder_path = 'db'
        folder_names = []

        while True:
            try:
                # Get folder names from the 'db' directory
                folder_names = [folder for folder in os.listdir(db_folder_path) if os.path.isdir(os.path.join(db_folder_path, folder))]
                time.sleep(0.1)
            except Exception as e:
                print(f"Error reading folder names: {e}")
                # Handle the error (e.g., log, retry, or break out of the loop)
                pass
            else:
                break

        return render_template('adminpages/admin_new_academic_year.html', folder_names=folder_names, db_folder_path=db_folder_path)

    return render_template('commonpages/notfound.html')


# Create Folder route
@app.route('/create_folder', methods=['POST'])
def create_folder():
    folder_name = request.form['folder_name']
    folder_path = os.path.join('db', folder_name)

    if os.path.exists(folder_path):
        db_folder_path = 'db'
        folder_names = [folder for folder in os.listdir(db_folder_path) if os.path.isdir(os.path.join(db_folder_path, folder))]
        message = f'The Acadimic Year "{folder_name}" already exists.'
        return render_template('adminpages/admin_new_academic_year.html', message=message, folder_names=folder_names)

    # If the folder does not exist, create it
    os.makedirs(folder_path)

    # Process section data
    sections = ['Student', 'Employer', 'Faculty', 'Alumni']
    for section in sections:
        column_names = []
        column_names.append("RollNumber")
        column_names.append("Password")
        column_names.append("Department")
        column_names.append("Feedback")
        column_names.append("Name")
        column_names.append("MobileNumber")
        column_names.append("Email")
        column_names.append("Suggestion")

        input_prefix = section + "_"

        # Loop through input fields to get column names
        i = 0
        for key, value in request.form.items():
            if key.startswith(input_prefix):
                i = i + 1
                column_names.append(str(i) + "." + value)

        df = pd.DataFrame(columns=column_names)
        excel_file_path = os.path.join(folder_path, f'{section}.xlsx')

        # Write DataFrame to Excel file with retry mechanism
        write_excel_with_retry(df, excel_file_path)

    # Process 'Departments' data
    department_column_names = []

    for key, value in request.form.items():
        if key.startswith('Department'):
            department_column_names.append(value)

    # Create DataFrame for 'Departments' and save it as CSV with retry mechanism
    department_df = pd.DataFrame(columns=department_column_names)
    department_csv_path = os.path.join(folder_path, 'Departments.csv')
    write_csv_with_retry(department_df, department_csv_path)

    db_folder_path = 'db'
    folder_names = [folder for folder in os.listdir(db_folder_path) if os.path.isdir(os.path.join(db_folder_path, folder))]
    message = f'The Acadimic Year "{folder_name}" created successfully.'
    return render_template('adminpages/admin_new_academic_year.html', message=message, folder_names=folder_names)

# Delete Folder route
@app.route('/delete_folder', methods=['POST'])
def delete_folder():
    folder_name = request.form['delete_folder']
    folder_path = os.path.join('db', folder_name)

    if os.path.exists(folder_path):
        # Delete the folder and its contents
        shutil.rmtree(folder_path)
        db_folder_path = 'db'
        folder_names = [folder for folder in os.listdir(db_folder_path) if os.path.isdir(os.path.join(db_folder_path, folder))]
        message = f'The Acadimic Year "{folder_name}" Deleted Sucessfully.'
    else:
        message = f'The Acadimic Year "{folder_name}" Does Not Exist.'

    return redirect(url_for('admin_new_academic_year', message=message, folder_names=folder_names))


# Flask route - admin_set_defaults
@app.route('/admin_set_defaults', methods=['GET', 'POST'])
def admin_set_defaults():
    admin = session.get('admin')

    if admin:
        db_folder_path = 'db'
        folders = read_folder_names()

        # Read the current value from the 'config.txt' file
        current_value = read_config_file()

        if request.method == 'POST':
            selected_folder = request.form.get('selected_folder')
            if selected_folder is not None:
                # Update the 'config.txt' file with the selected folder

                write_config_file(selected_folder)
                current_value = selected_folder
                folders = read_folder_names()

                return render_template('adminpages/admin_ui.html', folders=folders, message=f'Current default value: {current_value}')
            else:
                # Handle the case when 'selected_folder' is None (not selected)
                return render_template('adminpages/admin_ui.html', folders=folders, message=f'Current default value: {current_value}')

        # If it's a GET request
        return render_template('adminpages/admin_ui.html', folders=folders, message=f'Current default value: {current_value}')

    return render_template('commonpages/notfound.html')



@app.route('/admin_new_stakeholder', methods=['GET', 'POST'])
def submit_manual_feedback():
    admin = session.get('admin')
    if admin:

        with open('config.txt', 'r') as file:
            current_value = file.read().strip()


        # Assuming the file path is db/current_value/Departments.csv
        departments_file_path = os.path.join('db', current_value, 'Departments.csv')

        # Read the 'DepartmentName' column using pandas
        departments = []
        Student_department_counts =None
        Alumni_department_counts = None
        Employer_department_counts = None
        Faculty_department_counts = None
        try:
            # Use pandas to read CSV directly
            departments_df = pd.read_csv(departments_file_path)
            departments=departments_df.columns
            print(departments)
            excel_file = ['Student.xlsx', 'Alumni.xlsx', 'Employer.xlsx', 'Faculty.xlsx']

            for excel_file_name in excel_file:
                    try:
                        excel_file_path = 'db/'+current_value + '/' + excel_file_name
                        print(excel_file_path)
                        df =read_excel_with_retry(excel_file_path)
                        selected_columns = ["RollNumber", "Password","Department"]
                        df = df[selected_columns].copy()
                        df = df.dropna()
                        print(df)
                        print(excel_file_name)
                        if(excel_file_name=='Student.xlsx'):
                            Student_department_counts=df
                        elif(excel_file_name=='Alumni.xlsx'):
                             Alumni_department_counts=df
                        elif(excel_file_name=='Employer.xlsx'):
                             Employer_department_counts=df
                        elif(excel_file_name=='Faculty.xlsx'):
                             Faculty_department_counts=df
                    except Exception as e:
                            print(f"Error reading CSV file: {e}")

        except Exception as e:
            print(f"Error reading CSV file: {e}")

        if request.method == 'POST':
            roll_number = str(request.form.get('RollNumber')).strip().upper()
            password = str(request.form.get('Password')).strip().upper()
            department = str(request.form.get('Department')).strip()
            user_type = str(request.form.get('userType'))

            status = process_submission(roll_number, password, department, user_type)


            excel_file = ['Student.xlsx', 'Alumni.xlsx', 'Employer.xlsx', 'Faculty.xlsx']
            Student_department_counts = pd.DataFrame(columns=["RollNumber", "Password"])
            Alumni_department_counts = pd.DataFrame(columns=["RollNumber", "Password"])
            Employer_department_counts = pd.DataFrame(columns=["RollNumber", "Password"])
            Faculty_department_counts = pd.DataFrame(columns=["RollNumber", "Password"])
            for excel_file_name in excel_file:
                    try:
                        excel_file_path = 'db/'+current_value + '/' + excel_file_name
                        print(excel_file_path)
                        df =read_excel_with_retry(excel_file_path)

                        selected_columns = ["RollNumber", "Password"]
                        df = df[selected_columns].copy()
                        df=df.dropna()
                        if(excel_file_name=='Student.xlsx'):
                            Student_department_counts=df
                        elif(excel_file_name=='Alumni.xlsx'):
                             Alumni_department_counts=df
                        elif(excel_file_name=='Employer.xlsx'):
                             Employer_department_counts=df
                        elif(excel_file_name=='Faculty.xlsx'):
                             Faculty_department_counts=df
                    except Exception as e:
                        print(f"Error reading CSV file: {e}")
            print("student",Student_department_counts)
            print("////////////////////////")
            print("faculty",Faculty_department_counts)
            print("////////////////////////")
            print("alumni",Alumni_department_counts)
            print("////////////////////////")
            print("employer",Employer_department_counts)

            return render_template('adminpages/admin_new_stakeholders.html', departments=departments,
                                                                             student=Student_department_counts,
                                                                             alumni= Alumni_department_counts,
                                                                             employer=Employer_department_counts,
                                                                             faculty=Faculty_department_counts,
                                                                             message=status)
        print("student",Student_department_counts)
        print("////////////////////////")
        print("faculty",Faculty_department_counts)
        print("////////////////////////")
        print("alumni",Alumni_department_counts)
        print("////////////////////////")
        print("employer",Employer_department_counts)
        return render_template('adminpages/admin_new_stakeholders.html', departments=departments,
                                                                             student=Student_department_counts,
                                                                             alumni= Alumni_department_counts,
                                                                             employer=Employer_department_counts,
                                                                             faculty=Faculty_department_counts,
                                                                             message=None)
    else:
        return render_template('commonpages/notfound.html')


@app.route('/admin_new_stakeholder_drag', methods=['POST'])
def submit_drag_feedback():
    admin = session.get('admin')
    if admin:

        with open('config.txt', 'r') as file:
            current_value = file.read().strip()

        # Assuming the file path is db/current_value/Departments.csv
        departments_file_path = os.path.join('db', current_value, 'Departments.csv')

        # Read the 'DepartmentName' column using pandas
        departments = []
        try:
            # Use pandas to read CSV directly
            departments_df = pd.read_csv(departments_file_path)
            departments=departments_df.columns
            print(departments)
        except Exception as e:
            print(f"Error reading CSV file: {e}")
        try:
            if 'excelFile' in request.files:
                excel_file = request.files['excelFile']
                if excel_file.filename != '' and allowed_file(excel_file.filename):
                    df = pd.read_excel(excel_file)
                    entries = []

                    for index, row in df.iterrows():
                        roll_number = str(row['RollNumber'])
                        password = str(row['Password'])
                        user_type = str(row['UserType'])
                        department = str(row['Department'])

                        # Validate department name
                        if department not in departments:
                            status = f"Error: Department '{department}' not found in the list of valid departments."
                        else:
                            status = process_submission(roll_number, password, department, user_type)

                        entries.append({'RollNumber': roll_number, 'status': status, 'UserType': user_type})

                    return render_template('adminpages/admin_new_stakeholders.html', departments=departments, entries=entries, message="Excel file is inserted successfully. Check Status in Below Table.")

        except Exception as e:
            print(f"Error reading CSV file: {e}")
            return render_template('adminpages/admin_new_stakeholders.html', departments=departments, message=e )

    else:
        return render_template('commonpages/notfound.html')

def process_submission(roll_number, password,Department, user_type):
    with open('config.txt', 'r') as file:
        current_value = file.read().strip()
    user_data_folder = os.path.join('db', current_value)

    # Validate, process, and store the data in the Excel file
    if (
        pd.notna(roll_number) and
        pd.notna(password) and
        pd.notna(Department) and
        pd.notna(user_type) and
        str(roll_number).strip() != "" and
        str(password).strip() != "" and
        str(Department).strip() != "" and
        str(user_type).strip() != ""
    ):
        user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

        # Check if the file exists
        if os.path.exists(user_data_file):
            # Read user data from the Excel file
            user_data = pd.read_excel(user_data_file)

            # Check if the RollNumber already exists
            if roll_number not in user_data['RollNumber'].values:
                # Create a new DataFrame with the provided data and 'Feedback' set to 'No'
                new_data = pd.DataFrame({'RollNumber': [roll_number], 'Password': [password],'Department':[Department], 'Feedback': ['No']})

                # Concatenate the new data with the existing DataFrame
                user_data = pd.concat([user_data, new_data], ignore_index=True)

                # Save the updated DataFrame back to the Excel file
                user_data.to_excel(user_data_file, index=False)
                return f"Data added successfully : {roll_number} in Department : {Department} "
            else:
                return f"This {roll_number} already exists in Department :- {Department} . Please choose a different one."

    return "Error: in form submission. Please provide valid values for all fields."


def allowed_file(filename):
    ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'xlsm', 'xlsb', 'csv', 'xltx', 'xltm', 'xls', 'xlt', 'xml', 'xlw', 'xlam', 'xla'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS



@app.route('/admin_dashboard', methods=['GET', 'POST'])
def admin_dashboard():
    admin = session.get('admin')
    if admin:
        current_value = read_config_file()
        Student_department_counts = {}
        Alumni_department_counts = {}
        Employer_department_counts = {}
        Faculty_department_counts = {}
        Total_Student_counts = {}
        Total_Alumni_counts = {}
        Total_Employer_counts = {}
        Total_Faculty_counts = {}

        try:
            for attempt in range(1, 5):
                with open('config.txt', 'r') as file:
                    current_value = file.read().strip()
                base_path = 'db/' + current_value

                departments_file_path = base_path + '/Departments.csv'

                departments_df = pd.read_csv(departments_file_path)

                department_names = departments_df.columns
                excel_file = ['Student.xlsx', 'Alumni.xlsx', 'Employer.xlsx', 'Faculty.xlsx']

                for excel_file_name in excel_file:
                    try:
                        excel_file_path = base_path + '/' + excel_file_name
                        df = pd.read_excel(excel_file_path)



                        selected_columns = ["RollNumber", "Department", "Suggestion"]
                        suggestiondf = df[selected_columns].copy()
                        suggestiondf = suggestiondf.dropna()

                        columns_to_ignore = ["RollNumber", "Password", "Feedback", "Name", "MobileNumber", "Email", "Suggestion"]
                        df = df.drop(columns=columns_to_ignore, errors='ignore')

                        df=df.dropna()


                        department_counts = {}
                        transformed_counts = {}
                        for department in department_names:
                             virtualdf = df[df['Department'] == department]
                             virtualdf=virtualdf.drop(columns=["Department"], errors='ignore').to_dict()
                             for question, counts in  virtualdf.items():
                                # Initialize counts for each value
                                value_counts = {1: 0, 2: 0, 3: 0, 4: 0}

                                # Iterate over the counts in the original dictionary
                                for value, count in counts.items():
                                    # Update the corresponding count in the new dictionary
                                    value_counts[count] +=1

                                # Store the transformed counts for the question
                                transformed_counts[question] = value_counts
                                department_counts[department]=transformed_counts


                        if excel_file_name == 'Student.xlsx':
                            Student_department_counts = department_counts

                            Student_department_Suggestions = suggestiondf
                            Total_Student_counts = df.drop(columns=["Department"], errors='ignore').to_dict()

                        elif excel_file_name == 'Alumni.xlsx':
                            Alumni_department_counts = department_counts

                            Alumni_department_Suggestions = suggestiondf
                            Total_Alumni_counts = df.drop(columns=["Department"], errors='ignore').to_dict()
                        elif excel_file_name == 'Employer.xlsx':
                            Employer_department_counts = department_counts

                            Employer_department_Suggestions = suggestiondf
                            Total_Employer_counts = df.drop(columns=["Department"], errors='ignore').to_dict()
                        elif excel_file_name == 'Faculty.xlsx':

                            Faculty_department_counts = department_counts
                            Faculty_department_Suggestions = suggestiondf

                            Total_Faculty_counts = df.drop(columns=["Department"], errors='ignore').to_dict()

                    except Exception as e:
                        print(f"Error reading config.txt, Departments.csv, or Excel file: {e}")
                        return render_template('adminpages/error_template.html',
                                               error_message="Could not read config.txt, Departments.csv, or Excel file.")





            transformed_student_counts = {}

            tdata=[Total_Student_counts,Total_Alumni_counts,Total_Faculty_counts,Total_Employer_counts]

            # Iterate over the original dictionary
            for i in range(len(tdata)):
                transformed_student_counts={}
                for question, counts in tdata[i].items():
                    # Initialize counts for each value
                    value_counts = {1: 0, 2: 0, 3: 0, 4: 0}

                    # Iterate over the counts in the original dictionary
                    for value, count in counts.items():
                        # Update the corresponding count in the new dictionary
                        value_counts[count] +=1

                    # Store the transformed counts for the question
                    transformed_student_counts[question] = value_counts
                    if(i==0):
                        Total_Student_counts=transformed_student_counts

                    elif(i==1):
                        Total_Alumni_counts=transformed_student_counts

                    elif(i==2):
                        Total_Faculty_counts=transformed_student_counts

                    elif(i==3):
                        Total_Employer_counts=transformed_student_counts
            print("Student_department_counts")
            print(Student_department_counts)
            print("/////////////////////")
            print("Alumni_department_counts")
            print(Alumni_department_counts)
            print("/////////////////////")
            print("Employer_department_counts")
            print(Employer_department_counts)
            print("/////////////////////")
            print("Faculty_department_counts")
            print(Faculty_department_counts)
            print("/////////////////////")
            print("Total_Student_count",Total_Student_counts)
            print("/////////////////////")
            print("Total_Alumni_counts",Total_Alumni_counts)
            print("/////////////////////")
            print("Total_Employer_counts",Total_Employer_counts)
            print("/////////////////////")
            print(  "Total_Faculty_counts",Total_Faculty_counts)
            print("/////////////////////")



            return render_template('adminpages/admin_dashboard.html', current_value=current_value,
                                   student_department_counts=Student_department_counts,
                                   alumni_department_counts=Alumni_department_counts,
                                   faculty_department_counts=Faculty_department_counts,
                                   employer_department_counts=Employer_department_counts,
                                   Student_department_Suggestions=Student_department_Suggestions,
                                   Alumni_department_Suggestions=Alumni_department_Suggestions,
                                   Employer_department_Suggestions=Employer_department_Suggestions,
                                   Faculty_department_Suggestions=Faculty_department_Suggestions,
                                   Total_Student_counts=Total_Student_counts,
                                   Total_Alumni_counts=Total_Alumni_counts,
                                   Total_Employer_counts=Total_Employer_counts,
                                   Total_Faculty_counts=Total_Faculty_counts)

        except Exception as e:
            print(e)
            return render_template('error.html', error_message=str(e))

    return render_template('adminpages/admin_dashboard.html')


@app.route('/document_generator', methods=['GET', 'POST'])
def document_generator():
    admin = session.get('admin')
    if admin:
        if request.method == 'POST':
            try:
                current_value = request.form['selected_folder']

                base_path = 'db/' + current_value

                departments_file_path = base_path + '/Departments.csv'

                departments_df = pd.read_csv(departments_file_path)

                department_names = departments_df.columns
                excel_file = ['Student.xlsx', 'Alumni.xlsx', 'Employer.xlsx', 'Faculty.xlsx']

                # Initialize variables to None
                Student_department_Suggestions = None
                Alumni_department_Suggestions = None
                Employer_department_Suggestions = None
                Faculty_department_Suggestions = None

                for excel_file_name in excel_file:
                    try:
                        excel_file_path = base_path + '/' + excel_file_name
                        df = pd.read_excel(excel_file_path)
                        print("Before:", df)

                        selected_columns = ["RollNumber", "Department", "Suggestion"]
                        suggestiondf = df[selected_columns].copy()
                        suggestiondf = suggestiondf.dropna()

                        if excel_file_name == 'Student.xlsx':
                            Student_department_Suggestions = suggestiondf

                        elif excel_file_name == 'Alumni.xlsx':
                            Alumni_department_Suggestions = suggestiondf

                        elif excel_file_name == 'Employer.xlsx':
                            Employer_department_Suggestions = suggestiondf

                        elif excel_file_name == 'Faculty.xlsx':
                            Faculty_department_Suggestions = suggestiondf


                    except Exception as e:
                        print(f"Error reading config.txt, Departments.csv, or Excel file: {e}")
                        return render_template('adminpages/error_template.html',
                                               error_message="Could not read config.txt, Departments.csv, or Excel file.")

                # Check if any variable remains None
                if any(var is None for var in [Student_department_Suggestions, Alumni_department_Suggestions, Employer_department_Suggestions, Faculty_department_Suggestions]):
                    # Handle the case where any variable is None
                    return render_template('adminpages/error_template.html',
                                           error_message="One or more Excel files could not be read.")

                return render_template('adminpages/document_generator.html', department_names=department_names, current_value=current_value,usertype=excel_file_name.replace(".xlsx", ""),Student_department_Suggestions=Student_department_Suggestions, Alumni_department_Suggestions= Alumni_department_Suggestions, Employer_department_Suggestions= Employer_department_Suggestions, Faculty_department_Suggestions= Faculty_department_Suggestions)
            except Exception as e:
                print(e)
                return render_template('error.html', error_message=str(e))
    else:
        return render_template('error.html', error_message="unauthorized user")

@app.route('/pdf_download', methods=['GET', 'POST'])
def pdf_download():
    admin = session.get('admin')
    if admin:
        if request.method == 'POST':
            try:
                current_value = request.form.get('current_value')
                department = request.form.get('department')
                student_selected_rows = request.form.get('student_selected_rows')
                alumni_selected_rows = request.form.get('alumni_selected_rows')
                employer_selected_rows = request.form.get('employer_selected_rows')
                faculty_selected_rows = request.form.get('faculty_selected_rows')
                # Convert lists of dictionaries to DataFrames
                student_selected_rows = pd.DataFrame(eval(student_selected_rows))
                alumni_selected_rows = pd.DataFrame(eval(alumni_selected_rows))
                employer_selected_rows = pd.DataFrame(eval(employer_selected_rows))
                faculty_selected_rows = pd.DataFrame(eval(faculty_selected_rows))

                base_path = 'db/' + current_value
                excel_files = ['Student.xlsx', 'Alumni.xlsx', 'Employer.xlsx', 'Faculty.xlsx']

                Total_Student_counts = {}
                Total_Alumni_counts = {}
                Total_Employer_counts = {}
                Total_Faculty_counts = {}
                for excel_file_name in excel_files:
                    try:
                        excel_file_path = base_path + '/' + excel_file_name

                        if len(department)<0:
                            df = pd.read_excel(excel_file_path)
                            df=df.dropna()
                            df = df[df['Department'] == department]
                            # Convert 'MobileNumber' column to string data type
                            df['MobileNumber'] = df['MobileNumber'].astype(str)

                            # Remove decimal part if present
                            df['MobileNumber'] = df['MobileNumber'].str.split('.').str[0]
                            n_rows = min(5, len(df))
                            formdf = df.sample(n=n_rows, replace=False)

                            columns_to_ignore = ["RollNumber", "Password", "Feedback", "Name", "MobileNumber", "Email", "Suggestion"]
                            df = df.drop(columns=columns_to_ignore, errors='ignore')
                        else:
                            df = pd.read_excel(excel_file_path)
                            df=df.dropna()
                            # Convert 'MobileNumber' column to string data type
                            df['MobileNumber'] = df['MobileNumber'].astype(str)

                            # Remove decimal part if present
                            df['MobileNumber'] = df['MobileNumber'].str.split('.').str[0]
                            n_rows = min(5, len(df))
                            formdf =df.sample(n=n_rows, replace=False)
                            columns_to_ignore = ["RollNumber", "Password", "Feedback", "Name", "MobileNumber", "Email", "Suggestion"]
                            df = df.drop(columns=columns_to_ignore, errors='ignore')

                        if excel_file_name == 'Student.xlsx':
                            Studentform=formdf
                            print("Studentform")
                            print(Studentform)
                            print( Studentform.columns)
                            Total_Student_counts = df.drop(columns=["Department"], errors='ignore').to_dict()
                        elif excel_file_name == 'Alumni.xlsx':
                            Alumniform =formdf
                            print("Alumniform ")
                            print(Alumniform )
                            Total_Alumni_counts = df.drop(columns=["Department"], errors='ignore').to_dict()
                        elif excel_file_name == 'Employer.xlsx':
                            Employerform =formdf
                            print("Employerform")
                            print(Employerform)
                            Total_Employer_counts = df.drop(columns=["Department"], errors='ignore').to_dict()
                        elif excel_file_name == 'Faculty.xlsx':
                            Facultyform=formdf
                            print("Facultyform")
                            print(Facultyform)

                            Total_Faculty_counts = df.drop(columns=["Department"], errors='ignore').to_dict()
                    except Exception as e:
                        print(e)
                        print("pdf_generator")
                        return render_template('error.html', error_message=str(e))

                transformed_student_counts = {}

                tdata=[Total_Student_counts,Total_Alumni_counts,Total_Faculty_counts,Total_Employer_counts]
                # Iterate over the original dictionary
                for i in range(len(tdata)):
                    transformed_student_counts={}
                    for question, counts in tdata[i].items():
                        # Initialize counts for each value
                        value_counts = {1: 0, 2: 0, 3: 0, 4: 0}

                        # Iterate over the counts in the original dictionary
                        for value, count in counts.items():
                            # Update the corresponding count in the new dictionary
                            value_counts[count] +=1

                        # Store the transformed counts for the question
                        transformed_student_counts[question] = value_counts
                        if(i==0):
                            Total_Student_counts=transformed_student_counts

                        elif(i==1):
                            Total_Alumni_counts=transformed_student_counts

                        elif(i==2):
                            Total_Faculty_counts=transformed_student_counts

                        elif(i==3):
                            Total_Employer_counts=transformed_student_counts

                if(department):
                    pass
                else:
                    department="College wide"



                return render_template('adminpages/pdf_download.html',current_value=current_value,
                                       Total_Student_counts=Total_Student_counts,
                                       Total_Alumni_counts=Total_Alumni_counts,
                                       Total_Employer_counts=Total_Employer_counts,
                                       Total_Faculty_counts=Total_Faculty_counts,
                                       student_selected_rows=student_selected_rows,
                                       alumni_selected_rows=alumni_selected_rows,
                                       employer_selected_rows=employer_selected_rows,
                                       faculty_selected_rows=faculty_selected_rows,
                                       Studentform=Studentform,Alumniform=Alumniform,
                                       Employerform=Employerform,Facultyform=Facultyform,department=department

                                       )

            except Exception as e:
                print(e)
                print("pdf_generator")
                return render_template('error.html', error_message=str(e))
    else:
        return render_template('error.html', error_message="Unauthorized access")







@app.route('/login', methods=['GET', 'POST'])
def login():
    alert_message = None
    if request.method == 'POST':
        userid = request.form['userid'].strip().upper()
        password = request.form['password'].strip().upper()
        user_type = request.form.get('user_type')

        with open('config.txt', 'r') as file:
            current_value = file.read().strip()

        user_data_folder = os.path.join('db', current_value)
        user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

        if os.path.exists(user_data_file):
            user_data = pd.read_excel(user_data_file)


            # Check if the provided username and password match any row in the DataFrame
            matching_user = user_data[(user_data['RollNumber'].astype(str) == userid) & (user_data['Password'].astype(str) == password)]

            if not matching_user.empty:
                feedback_status = matching_user['Feedback'].iloc[0]
                session['userid'] = True
                if feedback_status == 'Yes':
                    return render_template('feedbackpages/fbalreadydone.html',userid=userid,user_type=user_type)
                elif not matching_user.empty and user_type == 'Faculty':
                    return redirect(url_for('facultyfb',userid=userid,user_type=user_type))
                elif not matching_user.empty and user_type == 'Student':
                    return redirect(url_for('studentfb',userid=userid,user_type=user_type))
                elif not matching_user.empty and user_type == 'Alumni':
                    return redirect(url_for('alumnifb',userid=userid,user_type=user_type))
                elif not matching_user.empty and user_type == 'Employer':
                    return redirect(url_for('employerfb',userid=userid,user_type=user_type))
                else:
                    alert_message = "Incorrect username or password"
            else:
                alert_message = "Incorrect username or password"
        else:
            alert_message = "Invalid user type"

    return render_template('loginpages/login.html', alert_message=alert_message)





@app.route('/fbalreadydone', methods=['GET', 'POST'])
def fbalreadydone():
    if request.method == 'POST':
        # Assuming you have the necessary code to identify the user and update the feedback column
        userid = request.form['userid']
        user_type = request.form.get('user_type')

        with open('config.txt', 'r') as file:
            current_value = file.read().strip()

        user_data_folder = os.path.join('db', current_value)
        user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

        if os.path.exists(user_data_file):
            user_data = pd.read_excel(user_data_file)

            # Check if the provided username matches any row in the DataFrame
            matching_user = user_data[user_data['RollNumber'].astype(str) == userid]

            if not matching_user.empty:
                # Update the Feedback column to 'No'
                user_data.loc[matching_user.index, 'Feedback'] = 'No'

                # Save the updated DataFrame back to the Excel file
                with pd.ExcelWriter(user_data_file) as writer:
                    user_data.to_excel(writer, index=False)
                # Redirect to the login route
                return redirect(url_for('login'))
            else:
                return  render_template('commonpages/notfound.html')
    else:
        return  render_template('commonpages/notfound.html')


@app.route('/facultyfb', methods=['GET', 'POST'])
def facultyfb():
    if session['userid'] == True:
        if request.method == 'POST':
            try:
                with open('config.txt', 'r') as file:
                    current_value = file.read().strip()

                user_type = "Faculty"
                user_data_folder = os.path.join('db', current_value)
                user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

                df = pd.read_excel(user_data_file)
                print(df)

                # Specify names to be removed
                names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

                # Remove specified names from the list of column names
                column_names = df.columns.tolist()
                for name in names_to_remove:
                    if name in column_names:
                        column_names.remove(name)

                userid = request.form.get('userid')
                print("userid",userid)
                matching_row = df[(df['RollNumber'].astype(str) == userid)]
                #matching_row = df[df['RollNumber'] == userid]
                print("matchedrow",matching_row)
                if not matching_row.empty:
                    # Update the existing row with feedback values
                    try:
                        feedback_data = {
                            'Name': request.form.get('Name'),
                            'MobileNumber': int(request.form.get('MobileNumber')),
                            'Email': request.form.get('Email'),
                            'Feedback': 'Yes',
                            'Suggestion': request.form.get('Suggestion')
                        }

                        for column_name in column_names:
                            feedback_value = request.form.get(column_name)
                            if feedback_value is not None:  # Skip columns not present in the form
                                # Explicitly convert to integer if possible
                                try:
                                    feedback_data[column_name] = int(feedback_value)
                                except ValueError:
                                    # If conversion to int fails, leave it as is
                                    feedback_data[column_name] = feedback_value

                        # Create a new DataFrame with the updated values
                        updated_df = df.copy()

                        # Get the index of the matching row
                        index_to_update = matching_row.index[0]

                        # Update the row with feedback values
                        updated_df.loc[index_to_update, ['Name','MobileNumber','Email','Feedback','Suggestion'] + column_names] = feedback_data

                        try:
                            # Use pd.ExcelWriter to explicitly close the Excel file after writing
                            with pd.ExcelWriter(user_data_file) as writer:
                                updated_df.to_excel(writer, index=False)

                        except Exception as e:
                            print("Error writing to Excel file:", e)
                            return render_template('error.html')

                        return render_template('feedbackpages/fbdone.html')

                    except Exception as e:
                        print("Error:", e)
                        return render_template('error.html',error_message=str(e))
                else:
                    return render_template('error.html',error_message="if condition not working")

            except Exception as e:
                print("Error:", e)
                return render_template('error.html',error_message=str(e))

        try:
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()

            userid = request.args.get('userid')
            user_type = "Faculty"
            user_data_folder = os.path.join('db', current_value)
            user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

            df = pd.read_excel(user_data_file)

            # Specify names to be removed
            names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

            # Remove specified names from the list of column names
            column_names = df.columns.tolist()
            for name in names_to_remove:
                if name in column_names:
                    column_names.remove(name)

            return render_template('feedbackpages/facultysfb.html', userid=userid, column_names=column_names,
                                user_type=user_type)

        except Exception as e:
            print("Error:", e)
            return render_template('error.html',error_message=str(e))
    else:
        return render_template('loginpages/login.html', alert_message="UnAuthorized Login Attempt")





@app.route('/studentfb', methods=['GET', 'POST'])
def studentfb():

    if session['userid'] == True:
        if request.method == 'POST':
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()

            user_type = "Student"
            user_data_folder = os.path.join('db', current_value)
            user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')
            df = pd.read_excel(user_data_file)

            # Specify names to be removed
            names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

            # Remove specified names from the list of column names
            column_names = df.columns.tolist()
            for name in names_to_remove:
                if name in column_names:
                    column_names.remove(name)

            userid = request.form.get('userid')
            matching_row = df[(df['RollNumber'].astype(str) == userid)]
            if not matching_row.empty:
                # Update the existing row with feedback values
                try:
                    feedback_data = {
                        'Name': request.form.get('Name'),
                        'MobileNumber': int(request.form.get('MobileNumber')),
                        'Email': request.form.get('Email'),
                        'Feedback': 'Yes',
                        'Suggestion': request.form.get('Suggestion')

                    }

                    for column_name in column_names:
                        feedback_value = request.form.get(column_name)
                        if feedback_value is not None:  # Skip columns not present in the form
                            # Explicitly convert to integer if possible
                            try:
                                feedback_data[column_name] = int(feedback_value)
                            except ValueError:
                                # If conversion to int fails, leave it as is
                                feedback_data[column_name] = feedback_value

                    # Create a new DataFrame with the updated values
                    updated_df = df.copy()

                    # Get the index of the matching row
                    index_to_update = matching_row.index[0]

                    # Update the row with feedback values
                    updated_df.loc[index_to_update, ['Name','MobileNumber','Email','Feedback','Suggestion'] + column_names] = feedback_data

                    try:
                        # Use pd.ExcelWriter to explicitly close the Excel file after writing
                        with pd.ExcelWriter(user_data_file) as writer:
                            updated_df.to_excel(writer, index=False)
                        session['userid'] = False

                    except Exception as e:
                        print("Error writing to Excel file:", e)
                        session['userid'] = False
                        return render_template('error.html',error_message=str(e))

                    return render_template('feedbackpages/fbdone.html')

                except Exception as e:
                    print("Error:", e)
                    session['userid'] = False
                    return render_template('error.html',error_message=str(e))
            else:
                session['userid'] = False
                return render_template('error.html')

        try:
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()

            userid = request.args.get('userid')
            user_type = "Student"
            user_data_folder = os.path.join('db', current_value)
            user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

            df = pd.read_excel(user_data_file)

            # Specify names to be removed
            names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

            # Remove specified names from the list of column names
            column_names = df.columns.tolist()
            for name in names_to_remove:
                if name in column_names:
                    column_names.remove(name)

            return render_template('feedbackpages/studentsfb.html', userid=userid, column_names=column_names,
                                user_type=user_type)

        except Exception as e:
            print("Error:", e)
            session['userid'] = False
            return render_template('error.html')
    else:
        session['userid'] = False
        return render_template('loginpages/login.html', alert_message="UnAuthorized Login Attempt")


@app.route('/alumnifb', methods=['GET', 'POST'])
def alumnifb():
    if session['userid'] == True:
        if request.method == 'POST':
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()

            user_type = "Alumni"
            user_data_folder = os.path.join('db', current_value)
            user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')
            df = pd.read_excel(user_data_file)

            # Specify names to be removed
            names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

            # Remove specified names from the list of column names
            column_names = df.columns.tolist()
            for name in names_to_remove:
                if name in column_names:
                    column_names.remove(name)

            userid = request.form.get('userid')
            matching_row = df[(df['RollNumber'].astype(str) == userid)]
            if not matching_row.empty:
                # Update the existing row with feedback values
                try:
                    feedback_data = {
                        'Name': request.form.get('Name'),
                        'MobileNumber': int(request.form.get('MobileNumber')),
                        'Email': request.form.get('Email'),
                        'Feedback': 'Yes',
                        'Suggestion': request.form.get('Suggestion')
                    }

                    for column_name in column_names:
                        feedback_value = request.form.get(column_name)
                        if feedback_value is not None:  # Skip columns not present in the form
                            # Explicitly convert to integer if possible
                            try:
                                feedback_data[column_name] = int(feedback_value)
                            except ValueError:
                                # If conversion to int fails, leave it as is
                                feedback_data[column_name] = feedback_value

                    # Create a new DataFrame with the updated values
                    updated_df = df.copy()

                    # Get the index of the matching row
                    index_to_update = matching_row.index[0]

                    # Update the row with feedback values
                    updated_df.loc[index_to_update, ['Name','MobileNumber','Email','Feedback','Suggestion'] + column_names] = feedback_data

                    try:
                        # Use pd.ExcelWriter to explicitly close the Excel file after writing
                        with pd.ExcelWriter(user_data_file) as writer:
                            updated_df.to_excel(writer, index=False)
                            session['userid'] = False

                    except Exception as e:
                        print("Error writing to Excel file:", e)
                        session['userid'] = False
                        return render_template('error.html')

                    return render_template('feedbackpages/fbdone.html')

                except Exception as e:
                    print("Error:", e)
                    session['userid'] = False
                    return render_template('error.html')
            else:
                session['userid'] = False
                return render_template('error.html')

        try:
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()

            userid = request.args.get('userid')
            user_type = "Alumni"
            user_data_folder = os.path.join('db', current_value)
            user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

            df = pd.read_excel(user_data_file)

            # Specify names to be removed
            names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

            # Remove specified names from the list of column names
            column_names = df.columns.tolist()
            for name in names_to_remove:
                if name in column_names:
                    column_names.remove(name)

            return render_template('feedbackpages/alumnisfb.html', userid=userid, column_names=column_names,
                                user_type=user_type)

        except Exception as e:
            print("Error:", e)
            session['userid'] = False
            return render_template('error.html')
    else:
        session['userid'] = False
        return render_template('loginpages/login.html', alert_message="UnAuthorized Login Attempt")



@app.route('/employerfb', methods=['GET', 'POST'])
def employerfb():
    if session['userid'] == True:
        if request.method == 'POST':
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()

            user_type = "Employer"
            user_data_folder = os.path.join('db', current_value)
            user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')
            df = pd.read_excel(user_data_file)

            # Specify names to be removed
            names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

            # Remove specified names from the list of column names
            column_names = df.columns.tolist()
            for name in names_to_remove:
                if name in column_names:
                    column_names.remove(name)

            userid = request.form.get('userid')
            matching_row = df[(df['RollNumber'].astype(str) == userid)]
            if not matching_row.empty:
                # Update the existing row with feedback values
                try:
                    feedback_data = {
                        'Name': request.form.get('Name'),
                        'MobileNumber': int(request.form.get('MobileNumber')),
                        'Email': request.form.get('Email'),
                        'Feedback': 'Yes',
                        'Suggestion': request.form.get('Suggestion')
                    }

                    for column_name in column_names:
                        feedback_value = request.form.get(column_name)
                        if feedback_value is not None:  # Skip columns not present in the form
                            # Explicitly convert to integer if possible
                            try:
                                feedback_data[column_name] = int(feedback_value)
                            except ValueError:
                                # If conversion to int fails, leave it as is
                                feedback_data[column_name] = feedback_value

                    # Create a new DataFrame with the updated values
                    updated_df = df.copy()

                    # Get the index of the matching row
                    index_to_update = matching_row.index[0]

                    # Update the row with feedback values
                    updated_df.loc[index_to_update, ['Name','MobileNumber','Email','Feedback','Suggestion'] + column_names] = feedback_data

                    try:
                        # Use pd.ExcelWriter to explicitly close the Excel file after writing
                        with pd.ExcelWriter(user_data_file) as writer:
                            updated_df.to_excel(writer, index=False)
                            session['userid'] = False

                    except Exception as e:
                        print("Error writing to Excel file:", e)
                        session['userid'] = False
                        return render_template('error.html')

                    return render_template('feedbackpages/fbdone.html')

                except Exception as e:
                    print("Error:", e)
                    session['userid'] = False
                    return render_template('error.html')
            else:
                session['userid'] = False
                return render_template('error.html')

        try:
            with open('config.txt', 'r') as file:
                current_value = file.read().strip()

            userid = request.args.get('userid')
            user_type = "Employer"
            user_data_folder = os.path.join('db', current_value)
            user_data_file = os.path.join(user_data_folder, f'{user_type}.xlsx')

            df = pd.read_excel(user_data_file)

            # Specify names to be removed
            names_to_remove = ['RollNumber', 'Password','Department', 'Feedback', 'Name', 'MobileNumber', 'Email', 'Suggestion']

            # Remove specified names from the list of column names
            column_names = df.columns.tolist()
            for name in names_to_remove:
                if name in column_names:
                    column_names.remove(name)

            return render_template('feedbackpages/employersfb.html', userid=userid, column_names=column_names,
                                user_type=user_type)

        except Exception as e:
            print("Error:", e)
            session['userid'] = False
            return render_template('error.html')
    else:
        session['userid'] = False
        return render_template('loginpages/login.html', alert_message="UnAuthorized Login Attempt")


@app.route('/FeedBackDone', methods=['GET', 'POST'])
def FeedBackDone():
    session['userid'] = False
    return render_template('feedbackpages/fbdone.html')

if __name__ == '__main__':
    app.run(debug=True)

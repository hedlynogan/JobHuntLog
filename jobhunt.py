import pandas as pd
from datetime import datetime

# Prompt the user for their name
user_name = input("Please enter your name for the EDD Job Hunt Log title: ")

# Define the columns for EDD job reporting requirements
columns = [
    'Week Ending Date', 'Employer Name', 'Type of Work Sought', 'Method of Contact',
    'Date of Contact', 'Position Applied For', 'Contact Person', "Contact Person's Title",
    'Contact Phone or Email', 'Result of Contact', 'Follow-up Actions Planned'
]

# Add a title to the spreadsheet using the user's name
current_date = datetime.now().strftime('%Y-%m-%d')
title = f"{user_name} Job Log - {current_date}"

# Create a DataFrame with the title and defined columns
# The title will be in the first row, and the column headers will be in the second row
# edd_job_reporting_log_with_title = pd.DataFrame(columns=[''])
title_df = pd.DataFrame({'A': [title]})
columns_df = pd.DataFrame(columns=columns)
edd_job_reporting_log_with_title = pd.concat([title_df, columns_df], ignore_index=True)

# Save the DataFrame with the title to an Excel file
edd_file_name_with_title = f"EDD_Job_Hunt_Log_for_{user_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
with pd.ExcelWriter(edd_file_name_with_title, engine='xlsxwriter') as writer:
    edd_job_reporting_log_with_title.to_excel(writer, index=False, header=False, startrow=2)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.write('A1', title)  # Write the title
    worksheet.set_row(0, 30)  # Set the title row height
    worksheet.set_row(1, 20)  # Set the columns row height

print(f"Spreadsheet created: {edd_file_name_with_title}")

# edd_job_reporting_log_with_title.loc[0] = title
# edd_job_reporting_log_with_title = edd_job_reporting_log_with_title._append(pd.DataFrame(columns=columns), ignore_index=True)

# Save the DataFrame with the title to an Excel file
# edd_file_name_with_title = f"EDD_Job_Hunt_Log_for_{user_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
# edd_job_reporting_log_with_title.to_excel(edd_file_name_with_title, index=False, header=False)

# print(f"Spreadsheet created: {edd_file_name_with_title}")




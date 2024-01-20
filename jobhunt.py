import pandas as pd
from datetime import datetime

# Define the structure of the spreadsheet for EDD job reporting requirements
# Commonly required fields may include:
columns = [
    'Week Ending Date',  # The last day of the job search week
    'Employer Name',  # Name of the employer where the applicant applied or interviewed
    'Type of Work Sought',  # The type of job the applicant is seeking
    'Method of Contact',  # How the applicant contacted the employer (e.g., online, in-person, phone)
    'Date of Contact',  # The date when the employer was contacted
    'Position Applied For',  # The specific position or job title the applicant applied for
    'Contact Person',  # Name or title of the person contacted at the employer
    "Contact Person's Title",  # Title of the contact person
    'Contact Phone or Email',  # Phone number or email address of the contact person
    'Result of Contact',  # The result of the contact (e.g., interview scheduled, no response, job offer)
    'Follow-up Actions Planned'  # Any follow-up actions the applicant intends to take
]

# Create a DataFrame with the defined columns
edd_job_reporting_log = pd.DataFrame(columns=columns)

# Save the DataFrame to an Excel file
edd_file_name = f"EDD_Job_Reporting_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
edd_job_reporting_log.to_excel(edd_file_name, index=False)

edd_file_name

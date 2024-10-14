# Automated Report Refresh and Distribution via VBA

This project includes two macros developed to streamline the monthly/weekly process of refreshing reports in *Analysis for Microsoft Excel* and sending them via email. The macros automate the task of data refresh and file distribution, reducing manual effort and improving efficiency.

## Key Components

### 1. File Refresh Macro
This macro is responsible for automatically refreshing files in *Analysis for Microsoft Excel*. It:

- Opens all the necessary files listed in **Column B** of the Excel sheet.
- Refreshes the data within those files.
- Saves the updated versions to the specified location.

### 2. Email Sending Macro
Once the files are refreshed, this macro automates the process of emailing the reports to designated recipients. It:

- Sends the refreshed files via *Outlook* to the controllers listed in **Column A**.
- Attaches the corresponding files listed in **Column B**.
- Adds custom email headers from **Column C** and body text from **Column D**.

## How It Works

1. The macros are triggered in sequence, with the file refresh macro executing first and the email sending macro following afterward.
2. The Excel sheet acts as the central command, storing the file paths, recipient lists, email headers, and body text.
3. The combined process takes approximately **10-15 minutes** to complete.

### Challenges and Observations

- **Execution Time**: The total execution time for both macros is slightly longer than typical macros, primarily because of the large file sizes and multiple file operations.
- **Outlook Integration**: The interaction between Outlook and VBA causes a slight delay during the email sending process, which results in the longer execution time.

---



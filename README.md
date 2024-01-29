# INVENTORY-REPORT-AUTOMATION-FOR-MOVEMENT-MANUFACTURING

This Python script provides a tool for processing Excel files, merging them, and generating a pivot table chart. Additionally, it supports uploading classification and tag information, updating the merged file, and creating a final pivot table.

Usage
Upload Excel Files:

Execute the script and upload Excel files (e.g., FPS.xlsx, MCFP.xlsx) when prompted.
Processed files will be stored in the 'uploads' directory.
Merge Excel Files:

Processed files are merged into a single file named 'Merged_Files.xlsx'.
Unwanted data is filtered based on specific conditions.
Upload Additional Information:

Optionally upload Classification_Base.xlsx, Tag1.xlsx, and Tag2.xlsx.
Classification and tag information is added to the merged file.
Create Pivot Table:

A pivot table is generated based on the 'Tag' and 'Classification' columns in the merged file.
The resulting pivot table is stored in a new sheet named 'Pivot table'.
Chart Generation:

The pivot table data is used to create a stacked bar chart.
The chart is embedded in a PowerPoint presentation ('output_chart.pptx').
Email with Attachment and Image:

An email is sent with the PowerPoint file ('output_chart.pptx') attached.
An embedded image extracted from the presentation is also included.
Requirements
Ensure the required libraries are installed by running:

python
Copy code
pip install pandas openpyxl XlsxWriter python-pptx
Instructions
Run the Script:

Execute the entire script.
Upload Excel files when prompted.
Review Processed Files:

Check the 'uploads' directory for processed and merged files.
Optional: Upload Additional Files

Respond to prompts to upload Classification_Base.xlsx, Tag1.xlsx, and Tag2.xlsx.
Download Outputs:

Download the merged file ('Merged_Files.xlsx') and PowerPoint presentation ('output_chart.pptx').
Email Output:

Provide email details (sender_email, recipient_email, smtp_server, smtp_port, smtp_username, smtp_password).
The script will send an email with the PowerPoint file attached and an embedded image.
Note: Ensure that the email service provider allows access to less secure apps, or use an app-specific password.

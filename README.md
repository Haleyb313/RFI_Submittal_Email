# RFI_Submittal_Email
Weekly email with report tables inside email body and attached as PDFs

Goal: Send a weekly email that contains the RFI and Submittal reports
  Sub-goal:  Allow recipients to copy data directly out of the email body

Process:

First open Excel and run macros to 
  1) Clear out old, previously created files, 
  2) Make PDF copies of the RFI and Submittal tables from their respective sheets
  3) Save each table as an HTML file
  
The code runs faster than Excel completes the macros, so I have an If Statment for both my PDF and HTML files to count how many files are inside the folder and wait until the correct total exists.
  
Next, make the email by building the HTML code
  1) Retrieve my base body text
  2) Use an If Statment to check if both HTML files exist
  3) Append the HTML file for the RFI table (thus keeping all CSS and Styling from Excel)
  4) Append the HTML file for the Submittal table (thus keeping all CSS and Styling from Excel)

Next, add the PDF attachments
  1) Use an If Statment to check if both PDF files exist
  2) Attach both files to the email

Finally, send the email

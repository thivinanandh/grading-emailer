# Python script for PDF Report generation and Email 
This script reads the excel sheet with marks and generates a summary of reports to be emailed to the users. 

## Files Included
- A template html file ( template of the report with html and css )
- A Yaml file for configuration 
- Scoring Sheet in Excel. 

## Steps 
### Edit the Configration (yaml) file 
---
#### Excel File Related 


- inputFileName :  name of the sheet 
- masterMarkSheetName : sheet name in excel file 
- columnStart : Start of column number in excel (Excluding Serial No)
- HeaderEmail:  Column header which stores email id
- HeaderName: Column header which stores names

#### Section-1
for displaying all raw marks scored by the student 
- section-1-Heading: Table heading ## Not advised to be changed 
- ActualMarks: Column headers of Actual marks scoredto be included in the table
- NormalisedMarks: Column headers of Actual marks scoredto be included in the table


#### Section-2
Summary of all marks (preferably for 100)
- section-2-Needed: flag to display section 2
- section-2-marks: Column headers to be included in the table


#### Email 

- sendEmail : Sends email if `True`, else only generates reports
- EmailSubject: Subject of Email 
- cc : persons to be cced in mail 

### 2: Run the python script 
---

```
python3 mailcode.py
```


### 3: Send Email 
When it comes to send email the user will be presented wih three options. 
- `Send Email` - Only sends email of that particular student   **Use this first to validate the email sent before using the below command to disable prompt**
- `Send all Email without Prompt`  - This will disable the prompt and sends email to all users without asking for any prompt 
- `Stop` - Will exit the program 









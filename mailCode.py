# %%

# from turtle import shearfactor
# from unittest import main
from jinja2 import FileSystemLoader
import pandas as pd
import numpy as np 
import matplotlib.pyplot as plt
import weasyprint
from tqdm import tqdm
import time

import yaml  #pyyaml required
# %%
# read the yaml file

with open("parameters.yaml", "r") as stream:
    try:
        inputParameters = yaml.safe_load(stream)
    except yaml.YAMLError as exc:
        print(exc)
        exit(0);
# %%
dataMaster = pd.read_excel(inputParameters['inputFileName'],sheet_name=inputParameters['masterMarkSheetName'],skiprows=inputParameters['skipRows'],nrows=inputParameters['nRows'],
                    usecols=range(0,inputParameters['columnsEnd']) , engine='openpyxl')
dataMaster

EmailId = dataMaster[inputParameters["HeaderEmail"]].to_list();
Name = dataMaster[inputParameters["HeaderName"]].to_list();
SendAllEMAILS = False #### DANGEROUS VARIABLE, DO NOT EDIT


## print the SMTP email id and password from the yaml file
print("SMTP Email ID: ",inputParameters['SMTPEmailID'])
print("SMTP Email Password: ",inputParameters['SMTPEmailPassword'])



# %%

for i in tqdm(range(dataMaster.shape[0])):  #
    studentId = i;
    emailId = EmailId[studentId]
    personalDetails = {};
    personalDetails["name"] = Name[studentId].upper()
    personalDetails["email"] = EmailId[studentId]
    from jinja2 import Environment,FileSystemLoader
    import os
    templateLoader = FileSystemLoader(searchpath="./")
    templateEnv = Environment(loader=templateLoader)
    jinjaTemplateName = inputParameters['jinjaTemplateFileName']
    template = templateEnv.get_template(jinjaTemplateName)

    filename = "reports/" + str(emailId) + ".html"
    
    import subprocess
    p = subprocess.Popen(f'mkdir -p reports'.split(" "))
    p.wait()

    p = subprocess.Popen(f'cp {jinjaTemplateName} {filename}'.split(" "))
    p.wait()


    pdffilename = "reports/" + str(emailId) + ".pdf"
    PsetMainDict = {}

    heading = inputParameters['section-1-Heading']

    dictmarks  = dataMaster.loc[studentId]

    ## get the list of All Main Marks 
    MainMarksList = [key for key,value in inputParameters['ActualMarks'].items()  ]
    MainMarks = list(dataMaster[MainMarksList].loc[studentId])

    ## Normalised Marks
    NormalisedMarksList = [key for key,value in inputParameters['NormalisedMarks'].items()  ]
    
    NormalisedMarks = list(dataMaster[NormalisedMarksList].loc[studentId])
    NormalisedMarks = ['%.2f' % elem if "ubm" not in str(elem) else elem for elem in NormalisedMarks ]
    RowItems = [key for key,value in inputParameters['ActualMarks'].items()  ]
    MaxMarks   = [value for key,value in inputParameters['ActualMarks'].items()  ]
    NormMaxMarks  = [value for key,value in inputParameters['NormalisedMarks'].items()  ]

    mainMarkList = [];
    mainMarkList.append(RowItems)
    mainMarkList.append(MaxMarks)
    mainMarkList.append(MainMarks)
    mainMarkList.append(NormMaxMarks)
    mainMarkList.append(NormalisedMarks)

    if(inputParameters['section-2-Needed']):
        ## Final table 
        columns = [key for key,value in inputParameters['section-2-marks'].items()  ]
        MainMarksFinal = list(dataMaster[columns].loc[studentId])
        MainMarksFinal = ['%.2f' % float(elem) if "ubm" not in str(elem) else elem for elem in MainMarksFinal ]
        MaxMarksFinal   =  [value for key,value in inputParameters['section-2-marks'].items()  ]
        FinalHeading  = inputParameters['section-2-Heading']
        FinalRowItems =  [key for key,value in inputParameters['section-2-marks'].items()  ]

        mainMarkFinalList = []
        mainMarkFinalList.append(FinalRowItems)
        mainMarkFinalList.append(MaxMarksFinal)
        mainMarkFinalList.append(MainMarksFinal)
    else:
        mainMarkFinalList = []
        FinalRowItems = []
        FinalHeading = []

    # print(PenaltyPresent)
    with open(filename ,"w") as fh:
        fh.write(template.render(
            ##Variables
            personalDetails = personalDetails,
            PsetHeader = [],     ## Archived Codes
            dictmarks=dictmarks,
            heading=heading,
            mainMarkList= mainMarkList,
            RowItems = RowItems,
            len=len,
            printSection2 = inputParameters['section-2-Needed'],
            FinalHeading= FinalHeading,
            mainMarkFinalList = mainMarkFinalList,
            FinalRowItems = FinalRowItems,
            SubjectName = inputParameters['SubjectName'],
            instructionArray = inputParameters['instructionArray'],
        ))

    
    pdf = weasyprint.HTML(filename).write_pdf()
    open(pdffilename, 'wb').write(pdf)
    # print("Generated Reports for ", EmailId[studentId])


    if(not inputParameters['sendEmail']):
        continue;

    if(not SendAllEMAILS):
        c = input("Type 'Send Email' as confirmation for sending one email \n PLease check the Email Status on sent Email \n Enter 'Send all Email without Prompt' to send all further emails without prompt \n Enter 'Stop' to exit ")

        if(c == 'Send all Email without Prompt'):
            print("Made True")
            SendAllEMAILS = True
        
    if(c == "Stop" or c == "stop"):
        exit(0)

    if((c == 'Send Email' or SendAllEMAILS) and inputParameters['sendEmail']):
        from email.mime.text import MIMEText
        from email.mime.image import MIMEImage
        from email.mime.application import MIMEApplication
        from email.mime.multipart import MIMEMultipart
        import smtplib
        import os


        smtp = smtplib.SMTP('smtp.office365.com', 587)
        smtp.ehlo()
        smtp.starttls()
        smtp.login(inputParameters['SMTPEmailID'], inputParameters['SMTPEmailPassword'])


        msg = MIMEMultipart()

        subject = inputParameters['EmailSubject']
        msg['Subject'] = subject

        one_attachment = pdffilename

        cc=inputParameters['cc']
        to=[emailId]

        msg['To'] =','.join(to)
        msg['Cc']=','.join(cc)   
        toAddress = to + cc    

        with open(one_attachment, 'rb') as f:
            file = MIMEApplication(
                f.read(), name=os.path.basename(one_attachment)
            )
            file['Content-Disposition'] = f'attachment; \
            filename="{os.path.basename(one_attachment)}"'
            msg.attach(file)

        smtp.sendmail(from_addr=inputParameters['SMTPEmailID'],
                    to_addrs=toAddress, msg=msg.as_string())
        smtp.quit()

        
        if(SendAllEMAILS):
            time.sleep(inputParameters['sleepTime'])
# %%

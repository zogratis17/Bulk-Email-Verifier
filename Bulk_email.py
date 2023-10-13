import csv,openpyxl
import pandas as pd
import re,email,is_disposable_email
from email_validator import validate_email, EmailNotValidError

excel=openpyxl.Workbook()
sheet=excel.active
sheet.title='Verified Emails'
sheet.append(['Email','Validate Email','Domain Address','Disposable Email','Deliverable Email','Reason'])


with open('email.csv',encoding='utf-8-sig') as csvfile:
    reader=csv.DictReader(csvfile)
    email_list=[]
    for row in reader:
        email_list.append(row['Email']) 


Size_of_List=len(email_list)
isvalidemail=[]
isDomainAddress=[]
isDisposableMail=[]
isDeliverableMail=[]
isReason=[]


def domainAddress(res): 
    res = res.split('@')[1]
    isDomainAddress.append(res)

def checkemail(s):
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
    if re.match(pattern,s):
        isvalidemail.append("Valid Email")
    else:
        isvalidemail.append("Invalid Email")

def disposableEmail(email): 
    result = is_disposable_email.check(email)
    if(result==True):
        isDisposableMail.append("Yes")
    else:
        
        isDisposableMail.append("No")

def emailValidate(email):  
    try:
        emailinfo = validate_email(email, check_deliverability=True)
        email = emailinfo.normalized
        isDeliverableMail.append("Yes")
        isReason.append("-")
    except EmailNotValidError as e:
        isDeliverableMail.append("No")
        isReason.append(str(e))         


for i in range(0,Size_of_List):
    email=email_list[i]    

    checkemail(email)

    domainAddress(email)

    disposableEmail(email)

    emailValidate(email)

try:
    for i in range(0,Size_of_List):
        sheet.append([email_list[i],isvalidemail[i],isDomainAddress[i],isDisposableMail[i],isDeliverableMail[i],isReason[i]]) 
    excel.save('Email Verification.xlsx')
    print("Successfull")
except Exception as e:
    print("error")

from traceback import print_tb
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
from email.mime.application import MIMEApplication
import sys
from email.mime.image import MIMEImage

# defining funtion for program scale sorting:

def sortProgramScale(val):
    val=str(val)
    if len(val)<5:
        return False
    if int(val[0])>= 2:
        return True
    if int(val[1])>= 3:
        return True
    if int(val[2])>= 4:
        return True
    if int(val[3])>= 5:
        return True
    if int(val[4])>= 2:
        return True
    return False

#Reading Excel
df =pd.read_excel('MIS_Report.xlsx')
#Selecting required data from excel and renaming
req_col =['VC_Number','DR_status','Program_scale','GCI_F','Owner','Description','Applicability_yes','Revision','Sequence']
data =df[req_col]
data.rename(columns = { 'Applicability_yes' : 'No_of_Usecase_Mapped'}, inplace = True)

data3 = data
data3 = data3[(data3['Program_scale'] != '(null)') ]
data3 = data3[(data3['Program_scale'] != 'Nil') ]

#Latest Revison Sequence

data3.rename(columns = { 'index' : 'SR'}, inplace = True )
data3.sort_values(by=['Revision','Sequence'], inplace=True)
duplicateRows = data3[data3.duplicated(['VC_Number'])]
if 'NR' in duplicateRows['Revision'].values :
        print("\nThis value NR exists in Dataframe")
        SR_val=duplicateRows[duplicateRows['Revision']=='NR'].SR.values
        for i in SR_val:
            data.drop(data.SR[i], axis=0,inplace=True)
else :
    print("\nThis value does not exists in Dataframe")
data3.drop_duplicates(subset='VC_Number', keep='last', inplace=True)

# making a empty list to store condtions of program scale
ts=[]
for i in data3['Program_scale']:
    f = sortProgramScale(i)
    ts.append(f)
data3['With Scale >= 23451']=ts
data3=data3[data3['With Scale >= 23451'] == True]

#sorting based on Release Vault
data = data[(data['Owner'] == 'Release Vault') ]
# Sorting based on Program scale
data = data[(data['Program_scale'] != '(null)') ]
data = data[(data['Program_scale'] != 'Nil') ]

# making a empty list to store condtions of program scale
ps=[]
for i in data['Program_scale']:
    f = sortProgramScale(i)
    ps.append(f)

#Adding new colum as with scale >= 23451
data['With Scale >= 23451']=ps
#sorting and taking only True condtions
data=data[data['With Scale >= 23451'] == True]

# Creating seprated Data frames for DRO DR1 DR2 DR3
datadr0 = data[(data['DR_status'] == 'DR0') ]
datadr1 = data[(data['DR_status'] == 'DR1') ]
datadr2 = data[(data['DR_status'] == 'DR2') ]
datadr3 = data3[(data3['DR_status'] == 'DR3') ]

#Countin DR wise VC numbers
DR0_count= datadr0.DR_status.count()
DR1_count= datadr1.DR_status.count()
DR2_count= datadr2.DR_status.count()
DR3_count= datadr3.DR_status.count()

#Counting Number of zero usecases mapped
zerousecasedr=datadr0[datadr0['No_of_Usecase_Mapped']==0].No_of_Usecase_Mapped.count()
zerousecasedr1=datadr1[datadr1['No_of_Usecase_Mapped']==0].No_of_Usecase_Mapped.count()
zerousecasedr2=datadr2[datadr2['No_of_Usecase_Mapped']==0].No_of_Usecase_Mapped.count()
zerousecasedr3=datadr3[datadr3['No_of_Usecase_Mapped']==0].No_of_Usecase_Mapped.count()

#Counting  number of VC having GCI less tha 85
gcidr0 = datadr0[(datadr0['GCI_F'] < 85 )].GCI_F.count()
gcidr1 = datadr1[(datadr1['GCI_F'] < 85 )].GCI_F.count()
gcidr2 = datadr2[(datadr2['GCI_F'] < 85 )].GCI_F.count()
gcidr3 = datadr3[(datadr3['GCI_F'] < 85 )].GCI_F.count()

#Creating summery table object
table1={
'DR Status': ['DR0','DR1','DR2','DR3 [Released & WIP]'],
'No. of VCs with Program Scale ≥ 23451':[DR0_count,DR1_count,DR2_count,DR3_count],
'No. of VCs with Zero Use Cases' :[zerousecasedr,zerousecasedr1,zerousecasedr2,zerousecasedr3],
'No. of VCs with GCI < 85%' :[gcidr0,gcidr1,gcidr2,gcidr3]
}

# Creating table1  data 
tabledf = pd.DataFrame(table1) 

#Converting Data frame into HTML Table
table1=tabledf.to_html(table_id="summery",index=False)


#Creating DR status wise tables 
req_col_dr=['VC_Number','Description','Program_scale','No_of_Usecase_Mapped','GCI_F']

# To add suffix * in DR3 VC taking owner in required columns
req_col_dr3=['VC_Number','Description','Program_scale','No_of_Usecase_Mapped','GCI_F','Owner']

datadr0=datadr0[req_col_dr]
datadr1=datadr1[req_col_dr]
datadr2=datadr2[req_col_dr]
datadr3=datadr3[req_col_dr3]

#Adding suffix as * 
datadr3.loc[datadr3['Owner'] == 'Release Vault', 'VC_Number'] = datadr3['VC_Number']+"*"
datadr3=datadr3[req_col_dr]

#Renaming tables colums as required
datadr0.rename(columns={'VC_Number':'VC No.','No_of_Usecase_Mapped':'No. of DVA Use Cases','Program_scale':'Program Scale','GCI_F':'GCI(A)'},inplace = True)
datadr1.rename(columns={'VC_Number':'VC No.','No_of_Usecase_Mapped':'No. of DVA Use Cases','Program_scale':'Program Scale','GCI_F':'GCI(A)'},inplace = True)
datadr2.rename(columns={'VC_Number':'VC No.','No_of_Usecase_Mapped':'No. of DVA Use Cases','Program_scale':'Program Scale','GCI_F':'GCI(A)'},inplace = True)
datadr3.rename(columns={'VC_Number':'VC No.','No_of_Usecase_Mapped':'No. of DVA Use Cases','Program_scale':'Program Scale','GCI_F':'GCI(A)'},inplace = True)



#Sorting data frames according to Use Case 
datadr0.sort_values(by=['No. of DVA Use Cases'], inplace=True)
datadr1.sort_values(by=['No. of DVA Use Cases'], inplace=True)
datadr2.sort_values(by=['No. of DVA Use Cases'], inplace=True)
datadr3.sort_values(by=['No. of DVA Use Cases'], inplace=True)


# Filter for GCI< 85 and DVA Use Cases 0
datadr0 = datadr0[(datadr0['GCI(A)'] < 85) | (datadr0['No. of DVA Use Cases'] == 0)  ]
datadr1 = datadr1[(datadr1['GCI(A)'] < 85) | (datadr1['No. of DVA Use Cases'] == 0)  ]
datadr2 = datadr2[(datadr2['GCI(A)'] < 85) | (datadr2['No. of DVA Use Cases'] == 0)  ]
datadr3 = datadr3[(datadr3['GCI(A)'] < 85) | (datadr3['No. of DVA Use Cases'] == 0)  ]

# Filtering decimal values from Program Scale

datadr0['Program Scale'] = datadr0['Program Scale'].astype(str).apply(lambda x: x.replace('.0',''))
datadr1['Program Scale'] = datadr1['Program Scale'].astype(str).apply(lambda x: x.replace('.0',''))
datadr2['Program Scale'] = datadr2['Program Scale'].astype(str).apply(lambda x: x.replace('.0',''))
datadr3['Program Scale'] = datadr3['Program Scale'].astype(str).apply(lambda x: x.replace('.0',''))

#Converting Data frames into HTML Table
if datadr1.empty:
    dr1t=""
    dr1html=""
else:
    dr1t =datadr1.to_html(table_id="dr1",index=False)
    dr1html='<h3>DR1 VCs</h3>'

if datadr0.empty:
    dr0t=""
    dr0html=''
else:
    dr0t = datadr0.to_html(table_id="dro",index=False)
    dr0html='<h3>DR0 VCs</h3>'

if datadr2.empty:
    dr2t=""
    dr2html=''
else:
    dr2t =datadr2.to_html(table_id="dr2",index=False)
    dr2html='<h3>DR2 VCs</h3>'
if datadr3.empty:
    dr3t=""
    dr3html=''
else:
    dr3t =datadr3.to_html(table_id="dr3",index=False)
    dr3html='<h3>DR3 VCs</h3>'



#Taking body contet and required style content from files
with open ('/var/www/html/DVA_MIS_Report/DVA_MIS/bin2/mailContent/style.txt') as file:
    table_style_html =file.read()

with open ('/var/www/html/DVA_MIS_Report/DVA_MIS/bin2/mailContent/body.txt') as body:
    body_html =body.read()

#Setting Summney table heading 
gatewise='<h3>Gate Wise Summary</h3>'

#Setting the signature
signature = '<p> <br/> Thank you ! <br/> DPDS & ERC QA </p>'


#Embeding image into body
image = '<br><img src="cid:image1"><br>'

#Creating whole HTML Code to render it into outlook
finalcode = body_html + table_style_html + gatewise + table1 + dr0html + dr0t + dr1html + dr1t + dr2html + dr2t + dr3html + dr3t + image + signature

#Taking inputs from shell DVA_NC.sh

VT=sys.argv[3]          #Vehicle Type
BU=sys.argv[4]          #Buisness Unit  
WK=sys.argv[5]          #Week   
MO=sys.argv[6]          #Month      
YE=sys.argv[7]          #Year


# Setting Up To Recipient list
if (VT=='BUS'):
    To= ['sarang.kavishwar@tatamotors.com', 'amulv@tatamotors.com', 'a.dungarwal@tatamotors.com', 'jagannath.sarkar@tatamotors.com','manoj.surana@tatamotors.com', 'mrinal.pandey@tatamotors.com', 'nitin.kulhare@tatamotors.com', 's.byakod@tatamotors.com', 's.agrahari@tatamotors.com', 'vikrant.bende@tatamotors.com']
elif(VT=='HCV'):
    To=['mwadje@tatamotors.com', 'ravindra.deshmukh@tatamotors.com', 'pranab.soumandal@tatamotors.com', 'amit.gupta@tatamotors.com', 'brijesh.gupta@tatamotors.com', 'jitendrakumar.singh@tatamotors.com', 'sushobhan.chatterjee@tatamotors.com', 'abhijit.chavan@tatamotors.com', 'nilesh.khankar@tatamotors.com', 'purushottam.k@tatamotors.com', 'dhirendrap.singh@tatamotors.com', 'pratik.lahane@tatamotors.com', 'manoranjan.sahu@tatamotors.com', 'narottam.pankaj@tatamotors.com', 'sandeep.ghosh@tatamotors.com','yogesh.adatiya@tatamotors.com']
elif(VT=='LCV'):
    To=['chetan.k@tatamotors.com', 'mangesh.uplenchwar@tatamotors.com', 'bhaskargodbole@tatamotors.com', 'm.belsare@tatamotors.com', 'bhupendra.bhat@tatamotors.com', 'mushtaq.saudagar@tatamotors.com', 'rakesh.nanda@tatamotors.com', 'hemant.potphode@tatamotors.com', 'shrikant.moyade@tatamotors.com', 'amit.rathod@tatamotors.com', 'ramakanta.swain@tatamotors.com', 'chandrashekhar.tayde@tatamotors.com', 'dhaval.salgaonkar@tatamotors.com']
elif(VT=='LMV'):
    To=['thakare.p@tatamotors.com', 'ruparaj.m@tatamotors.com', 'devendra.sangve@tatamotors.com', 'chetan.chawadimani@tatamotors.com', 'bhaskar.sathi@tatamotors.com', 'chaitanya.b@tatamotors.com', 'milind.dixit@tatamotors.com', 'santosh.sonar@tatamotors.com', 'vijay.lodhi@tatamotors.com', 'j.bhalerao@tatamotors.com', 'renuka.avachat@tatamotors.com', 'pramod.c@tatamotors.com', 'kulkarni.niketan@tatamotors.com', 'nilesh.kankariya@tatamotors.com', 'sachin.babar@tatamotors.com', 'vidyadhar.vaidya@tatamotors.com']
elif(VT=='MCV'):
    To=['chetan.k@tatamotors.com', 'mangesh.uplenchwar@tatamotors.com', 'bhaskargodbole@tatamotors.com', 'm.belsare@tatamotors.com', 'bhupendra.bhat@tatamotors.com', 'mushtaq.saudagar@tatamotors.com', 'rakesh.nanda@tatamotors.com', 'hemant.potphode@tatamotors.com', 'shrikant.moyade@tatamotors.com', 'amit.rathod@tatamotors.com', 'ramakant.swain@tatamotors.com', 'chandrashekhar.tayde@tatamotors.com', 'dhaval.salgaonkar@tatamotors.com']
elif(VT=='UVVAN'):
    To=['pramod.c@tatamotors.com','renuka.avachat@tatamotors.com','j.bhalerao@tatamotors.com','ruparaj.m@tatamotors.com']
elif(VT=='MUV'):
    To=['nitin.kamble@tatamotors.com', 'abhishek_singh@tatamotors.com', 'musaib.momin@tatamotors.com', 'joshi.mahesh@tatamotors.com', 'pranjal.shendre@tatamotors.com', 'm.gajanan@tatamotors.com', 'pednekar.tushar@tatamotors.com']
elif(VT=='PRD'):
    To=['ajp.ttl@tatamotors.com','samirt.ttl@tatamotors.com']
elif(VT=='CAR'):
    To=['makarand.deval@tatamotors.com', 'cv.kulkarni@tatamotors.com', 'atrajesh@tatamotors.com', 'awatechandrakant@tatamotors.com', 'kumar.p@tatamotors.com', 'yogesh.bhandari@tatamotors.com', 'sachin.lale@tatamotors.com', 'deshmukh.kishor@tatamotors.com', 
'umk770063@tatamotors.com', 'samir.ghodekar@tatamotors.com', 'p.yogeshb@tatamotors.com', 'mahavir.chimad@tatamotors.com', 'kanad.karandikar@tatamotors.com', 'priyadarshi.singh@tatamotors.com', 'subhajit.mahanty@tatamotors.com']
else:
    To=['ajp.ttl@tatamotors.com']

#To=['ajp.ttl@tatamotors.com']

#Setting up Emails:

sender = 'DQM Notifications <noreply@tatamotors.com>'
#To=['ajp.ttl@tatamotors.com']

CC=['kaushik.biswas@tatamotors.com', 'srijna.prasad@tatamotors.com', 'ajaym.ttl@tatamotors.com']
#CC=['samirt.ttl@tatamotors.com']
#BCC Only DPDS Team 
BCC=['vedd.ttl@tatamotors.com', 'samirt.ttl@tatamotors.com', 'ajp.ttl@tatamotors.com', 'gn923253.ttl@tatamotors.com']
#BCC=['ajp.ttl@tatamotors.com']
MAIL=To+CC+BCC


#Creating MIME mail messaged Object
msg = MIMEMultipart()
msg['Subject'] = f'WCQ Result 8 - GCI MIS | {BU} {VT} | WK{WK} {MO}, {YE}'
# msg['Subject'] = "Subjet Line"
msg['From'] = sender
msg['To'] = ", ".join(To)
msg['Cc'] = ", ".join(CC)
msg['Bcc'] = ", ".join(BCC)


#Attaching the HTML contente to mail
part1 = MIMEText(finalcode, 'html')
msg.attach(part1)

# Image attachemnt 

# This example assumes the image is in the current directory
fp = open('/var/www/html/DVA_MIS_Report/DVA_MIS/bin2/mailContent/table1.PNG', 'rb')
msgImage = MIMEImage(fp.read())
fp.close()

# Define the image's ID as referenced above
msgImage.add_header('Content-ID', '<image1>')
msg.attach(msgImage)



#PDF attachemnt file

filename = sys.argv[2]
filename = 'MIS_'+f'{filename}'+'.pdf'
# filename='sample.pdf'

#Attaching Report PDF
with open(filename, "rb") as f:
        attach = MIMEApplication(f.read(),_subtype="pdf")
attach.add_header('Content-Disposition','attachment',filename=str(filename))
msg.attach(attach)

#Creating SMTP Object for local host
smtpObj = smtplib.SMTP('localhost')


smtpObj.sendmail(sender, MAIL, msg.as_string()) 
smtpObj.quit()        
print ("Successfully sent email")



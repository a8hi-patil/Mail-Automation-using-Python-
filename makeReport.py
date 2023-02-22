#************************Python Libraries********************************

# ============================================================================
import os
from xml.dom.minidom import TypeInfo
import numpy as np
import pandas as pd
from glob import glob
from fpdf import FPDF
import matplotlib.pyplot as plt
import sys
import csv
import xlrd
from datetime import datetime, timedelta
# import datetime

WIDTH = 280
HEIGHT = 150
""" Date """
fromdt=sys.argv[5]
todt=sys.argv[6]

date_from = datetime.strptime(fromdt, '%Y/%m/%d')
FromDate = date_from.strftime('%d-%b-%Y')

date_to = datetime.strptime(todt, '%Y/%m/%d')
ToDate = date_to.strftime('%d-%b-%Y')
# FromDate = "01-12-2021"
# ToDate = "31-12-2021"

    
#**********************Main_Directories*****************************
Excel = "/var/www/html/DVA_MIS_Report/DVA_MIS/workArea/MIS_Report.xlsx"
Input_Path = '../data/'


#**********************Header&Footer********************************
class PDF(FPDF):

  def header(self):
    # TOP,LEFT, Right BorderLine
    self.set_draw_color(r=0, g=0, b=0)
    # self.line(10, 10, 290, 10)
    # self.line(10, 10, 10, 205)
    # self.line(290, 10, 290, 205)
    self.line(5, 7, 293, 7)
    self.line(5, 7, 5, 208)
    self.line(293, 7, 293, 208)
    # Logo Image
    self.image(name='/var/www/html/DVA_MIS_Report/DVA_MIS/resource/Tata-Logo.jpg', x=235, y=16, w=50, h=6)
    self.image(name='/var/www/html/DVA_MIS_Report/DVA_MIS/resource/images.png', x=19, y=12, w=17, h=15)


  def footer(self):    
    #Bottom BorderLine
    self.set_draw_color(r=0, g=0, b=0)
    self.line(5, 208, 293, 208)
    # self.line(10, 205, 290, 205)
    
    # y2 = self.get_y()
    # x2 = self.get_x()
    # print(y2)
    # print(x2)
    self.set_y(-15)
    self.set_x(179)
    self.image(name='/var/www/html/DVA_MIS_Report/DVA_MIS/resource/dpdslogo.jpg')
    self.set_font('Times', '', 11)
    self.set_text_color(191, 191, 191)
    self.set_y(-17)
    self.cell(0, 10, 'Digital Product Development Systems', 0, 0, 'C')
    self.set_font('Times', 'B', 12)
    self.set_x(85)
    self.set_text_color(0,0,0)
    
    # self.cell(0, 10, 'DPDS', 0, 0, 'C')
    
    
    # Go to 1.5 cm from bottom
    self.set_y(-12)
    self.set_font('Times', '', 11)
    self.set_text_color(191, 191, 191)
    self.cell(0, 10,u"\u00A9 Copywrite,Confidential,Tata Motors Limited" , 0, 0, 'L')
    self.cell(0, 10, 'Page %s' % self.page_no(), 0, 0, 'R')
    #Bottom BorderLine
    self.set_draw_color(r=0, g=0, b=0)
    # self.line(10, 205, 290, 205)
    
    
    
def SortData(weeklyData_path):
    #******sorting all csv files in ascending order*******
    files = sorted(glob(weeklyData_path + 'MIS*.csv'))
    # print(files)
    VC_Number = []
    Desc = []
    DR = []
    Rev = []
    Seq = []
    DVA_Usecases = []
    Applicability_Yes =[]
    # RELEASEDATE =[]
    DFQ_Attached = []
    NotOk = []
    DVA_OK = []
    DVA_NC = []
    DVA_Index = []
    Path = []
    split = ""
    Program_scale=[]
    GCI=[]
    GCI_F=[]
    DMU=[]
    Owner=[]

   #************reading each csv file**************
    m=0
    for i in files:
        
        
        #**********reading i(th) csv file*******************
                
        
        
        with open(i, 'r') as f:
            reader = csv.reader(f)
            desc_count = 0
            vc_count = 0
            rev_count = 0
            seq_count = 0
            DR_count = 0
            date_count = 0
            program_scale_count=0
            gci_count=0
            dmu_count=0
            counter = 0
            new_applicability = 0
            flag = 0
            PDF_Yes = 0
            PDF_No = 0
            NC_Yes = 0
            
            for n in reader:
                if vc_count == 1:
                    cell2 = n[4]
                    print(cell2)
                    VC_Number.append(cell2)
                    # print(VC_Number)
                if desc_count == 2:
                    cell = n[4]
                    # print(cell)
                    Desc.append(cell)
                    # print(Desc)
                if rev_count == 1:
                    cell3 = n[7]
                    # print(cell3)
                    Rev.append(cell3)
                    # print(Rev)
                    '''       if date_count == 5:
                    cell1 = n[4]
                    #cell1 => 2021/06/23-11:39:38:709
                    datesplit=cell1.split("-")
                    date=datesplit[0].strip()
                    # date => 2021/08/23 strip-to remove whitespace
                    date2=date.split("/")
                    cell2=date2[0]+"-"+date2[1]+"-"+date2[2]
                    #cell2 => 2021-08-23
                    # print(cell2)
                    RELEASEDATE.append(cell2)  '''
                if seq_count == 1:
                    cell4 = n[10]
                    # print(cell4)
                    Seq.append(cell4)
                    # print(Seq)
                if DR_count == 3:
                    cell5 = n[4]
                    # print(cell5)
                    DR.append(cell5)
                    # print(DR)
                if program_scale_count==2:
                    try:
                        cell6=n[7]
                        Program_scale.append(cell6)
                    except IndexError:
                        cell6='Nil'
                        Program_scale.append(cell6)
                    print('________________-----Program Scale --------------________________')
                    print(cell6)
                    print(type(cell6))
                    #print(int(cell6))
                    
                
                if program_scale_count==10:
                    try:
                        cell7=n[4]
                        if (cell7=='Nil' or cell7==''):
                            GCI.append(cell7)
                        else:    
                            
                            cell7=int(float(cell7))
                            print(cell7)
                            GCI.append(cell7)
                    except IndexError:
                        cell7='Nil'
                        GCI.append(cell7)


                if program_scale_count==9:
                    try:
                        cell8=n[4]
                        if (cell8=='Nil' or cell8=='' or cell8=='nan'):
                            cell8=0
                            DMU.append(cell8)
                        
                        else:
                            cell8=int(cell8)
                            DMU.append(cell8)
                    except IndexError:
                        cell8=0
                        DMU.append(cell8)
                if program_scale_count==4:
                    cell9 = n[4]
                    Owner.append(cell9)
                
                if( counter >= 14):
                    try:
                        if(n[13]=='Yes'):
                            new_applicability = new_applicability + 1
                            if(n[16] == 'Yes'):
                                PDF_Yes += 1
                                if(n[23]=='Yes' or n[23]==''):
                                    NC_Yes += 1
                    except:
                        new_applicability = new_applicability + 0




                desc_count += 1
                vc_count += 1
                rev_count += 1
                seq_count += 1
                DR_count += 1
                date_count += 1
                program_scale_count+=1
                gci_count+=1
                counter+=1
            
        
        
        
        pd.set_option("display.max_rows", None, "display.max_columns", None)
        
        #*************skip first 6 rows**********************
        # df1 = pd.read_csv(i,encoding='latin-1', skiprows = 6)
        df1 = pd.read_csv(i,encoding='latin-1', skiprows = 13)
        # print(df1)
        # print(i)
        
        # Path_remove = os.path.relpath(i, '../')
        splitpath=i.split("html/")
        # print(splitpath)
        
        Pathsplit = "http://172.26.113.57/"+splitpath[1]
        Path.append(Pathsplit)
        # Path = i
        # print(Path)
        
        
        
        #'''Rename Columne Name'''
        df1.rename(columns = { 'No of Use cases Mapped' : 'Usecase'}, inplace = True )
        df1.rename(columns = { "Use cases Applicability 'YES'" : 'Applicability_Yes'}, inplace = True )
        df1.rename(columns = { 'DFQ Documents Number' : 'DFQ_Documents'}, inplace = True )
        df1.rename(columns = { 'PDF' : 'PDF'}, inplace = True )
        df1.rename(columns = { 'DVA Meet the Expectations (Yes/No)' : 'Expectations'}, inplace = True )
        
        
        Usecases = df1.Usecase.sum()
        DVA_Usecases.append(Usecases)
        # print(DVA_Usecases)
        
        Applicability = df1.Applicability_Yes.sum()
        # Applicability_Yes.append(Applicability)
        Applicability_Yes.append(new_applicability)
        # print(Applicability_Yes)
        
        df2=df1[df1['DFQ_Documents'].isnull() | ~df1[df1['DFQ_Documents'].notnull()].duplicated(subset='DFQ_Documents')]
        
        # print(df2)
        # print(df1.shape)
        # print(df2.shape)
        
        Count_PDF = df2[(df2['PDF'] == 'Yes')].PDF.count()
        DFQ_Attached.append(Count_PDF)
        
        Count_No = df2[(df2['Expectations'] == 'No')].Expectations.count()
        NotOk.append(Count_No)
        # print(NotOk)
        
        Count_Ok = Count_PDF - Count_No
        if (Count_Ok< 0):
            Count_Ok=Count_Ok * -1
        
        #DVA_OK.append(Count_Ok)
        DVA_OK.append(NC_Yes)
        #print('#######################@@@@@@@@@@@@@@@@@@@@@!!!!!!!!!!!!!!!!!@@@@@@@@@@@@')
        #print(DVA_OK)
        # Count_PDF = len(a)
        # DFQ_Attached.append(Count_PDF)
        # Index = round((Count_PDF/Applicability)*100,0)
        # DVA_Index.append(Index)
        
        
        if new_applicability == 0:
            DVA_Index.append(0)
        else:
            # Index = round((Count_PDF/Applicability)*100,0)
            Index = round((NC_Yes/new_applicability)*100,0)
            # Index ="{:.2f}".format(Index)
            # Index = float(Index)
            #Index = round((Count_Ok/Applicability)*100,0)
            DVA_Index.append(Index)
        
        if Count_PDF == 0:
            DVA_NC.append(0)
        else:
            NC = round((Count_No/Count_PDF)*100,0)
            DVA_NC.append(NC)
        # print('#########$$$$$$$$$$$$$$$$$$$$$$$$$$$$DVA INDEX')
        print(DVA_Index)

        if ( GCI[m]=='Nil' or GCI[m]=='nan' or GCI[m]==''):
        
            GCI_V=((DVA_Index[m]/2) + DMU[m]/2)
            GCI_V = "{:.1f}".format(GCI_V)
        
            GCI_F.append(GCI_V)
            
        else :

            GCI_F.append(GCI[m])
        
        # print(GCI_F)  
        # print("gci ===",GCI[m])
        # print("DMU ====",DMU[m])
        # print(DVA_Index[m])
        # print('value of m --',m)
        # print(DVA_Index)
        m=m+1
        
        
     
    #****************Creating a dataframe of reqired data*******************
    
    df = pd.DataFrame({'VC_Number':VC_Number,'DR_status':DR,'Revision':Rev,
                        'Revision_Encode':Rev,
                        'Sequence':Seq,
                        'Program_scale':Program_scale,
                        'DVA_OK':DVA_OK,
                        'DMU':DMU,
                        'GCI':GCI,
                        'GCI_F':GCI_F,
                        'Owner':Owner,
                        # 'Demerit_VScore':Vscore,'Demerit_Rating':Sscore,
                        # 'GC_Index':Ascore,
                        'No of Usecae Mapped':DVA_Usecases,
                        'Applicability_yes':Applicability_Yes,
                        'DFQ_Attached' :DFQ_Attached,
                        'DVA_Index' : DVA_Index,
                        'DFQ NotOk' : NotOk,
                        'DVA NC' : DVA_NC,
                        'Path':Path,'Description':Desc,}) 
                        
    # data['Revision_Encode'].replace({0:'NR',1:'A'},inplace=True)
    #*************DVA_index Column Color with conditions(not use)****************
    def color(val):
        if val >= 85 :
            color = 'green'
        elif 80 <= val < 85 :
            color = 'yellow'
        else:
            color = 'red'
        return 'background-color: %s' % color
    
    
    #**************Clubing Multiple csv to single excel*********************
    #**************Excel is given as input to pd.read_excel*********************
    # df.style. \
        # applymap(color, subset=pd.IndexSlice[:, ['DVA_Index']]).\
        # to_excel(Excel, engine="openpyxl" , index_label='index')
        
    df.to_excel(Excel, engine="openpyxl" , index_label='index')
    
    #df.style.applymap(color,subset=pd.IndexSlice[:, ['DVA_Index']].to_csv("../workArea/Report52.csv",encoding='latin-1'))
    #df.to_excel('../workArea/Report52.xlsx')
    
    

def DR0(pdf):
    
    DR="DR0"
    a = Latest_Rev(DR)
    Count_DR0 = a[(a['DR_status'] == DR)].DR_status.count()
    # print(Count_DR0)
    
    if Count_DR0 != 0:
        Table(a,pdf,DR)
    
def DR1(pdf):
    
    DR="DR1"
    a = Latest_Rev(DR)
    Count_DR1 = a[(a['DR_status'] == DR)].DR_status.count()
    # print(Count_DR1)

    
    if Count_DR1 != 0:
        Table(a,pdf,DR)
    
    

def DR2(pdf):
    
    DR="DR2"
    a = Latest_Rev(DR)
    Count_DR2 = a[(a['DR_status'] == DR)].DR_status.count()
    # print(Count_DR2)
    
    if Count_DR2 != 0:
        Table(a,pdf,DR)
 
    
def DR3(pdf):
    
    DR="DR3"
    a = Latest_Rev(DR)
    Count_DR3 = a[(a['DR_status'] == DR)].DR_status.count()
    # print(Count_DR3)
    
    
    if Count_DR3 != 0:
        Table(a,pdf,DR)
        
def DR3P(pdf):
    
    DR="DR3P"
    a = Latest_Rev(DR)
    Count_DR3P = a[(a['DR_status'] == DR)].DR_status.count()
    # print(Count_DR3P)
    
    if Count_DR3P != 0:
        Table(a,pdf,DR)
        
def DR4(pdf):
    
    DR="DR4"
    a = Latest_Rev(DR)
    Count_DR4 = a[(a['DR_status'] == DR)].DR_status.count()
    # print(Count_DR3P)
    
    if Count_DR4 != 0:
        Table(a,pdf,DR)

def Latest_Rev(DR):
    
    #*******************Opening the above excel******************************
    
    #data = pd.read_csv('../workArea/Report52.csv', encoding='latin-1')
    dataCom = pd.read_excel(Excel, engine="openpyxl")
    # print(dataCom)
    dataCom.rename(columns = { 'index' : 'SR'}, inplace = True )
    data = dataCom[(dataCom['DR_status'] == DR) ]
    data.sort_values(by=['Revision','Sequence'], inplace=True)
    duplicateRows = data[data.duplicated(['VC_Number'])]

    if 'NR' in duplicateRows['Revision'].values :
        print("\nThis value NR exists in Dataframe")
        SR_val=duplicateRows[duplicateRows['Revision']=='NR'].SR.values
        for i in SR_val:
            data.drop(data.SR[i], axis=0,inplace=True)
    else :
        print("\nThis value does not exists in Dataframe")
        
    data.drop_duplicates(subset='VC_Number', keep='last', inplace=True)
    data.sort_values(by=['DVA_Index','Applicability_yes'], inplace=True)
    
    return data
        

def HeaderPDF(pdf):    
    '''Page frontend(header and footer define at top)'''
    '''
    set_font = font style and size
    ln = space between 2 lines
    set_x = from x it will start to write
    cell = write in box
    write = write the sentence (not in rec box)
    draw_color = border color
    set_fill_color = cell color (IMP- True value should be written in cell to apply color)
    '''
    
    '''Pg Title'''
    pdf.set_draw_color(r=215, g=215, b=215)
    pdf.set_font('Times', 'I', 11)
    pdf.set_text_color(125, 125, 125)
    pdf.set_x(x=19) 
    pdf.write(5, 'This is a system generated report.')
    pdf.set_text_color(0, 0, 0)
     
    pdf.ln(5)
    pdf.set_font('Times', 'B', 17)
    pdf.set_x(x=5.2)
    pdf.set_fill_color(46, 117, 182)
    pdf.set_text_color(217,217,217) 
    pdf.set_draw_color(r= 46, g = 117, b=182)
    #pdf.cell(287.6, 13,  'Geometry Conformance Index', 1, 0, 'C',True)
    pdf.cell(287.6, 13,  'GEOMETRY CONFORMANCE INDEX (A)', 1, 0, 'C',True)
    pdf.set_text_color(0, 0, 255) 
    pdf.set_fill_color(0, 0, 0)
    
    pdf.set_text_color(125, 125, 125) 
    
    pdf.ln(15)
     
    pdf.set_font('Times', 'B', 12)
    pdf.set_x(x=25) 
    
    pdf.write(5, 'Some paragraph Some paragraph Some paragraph Some paragraph ')
    
    pdf.set_font('Times', '', 12)
    pdf.write(5, ' Some paragraph Some paragraph Some paragraph Some paragraph')
    pdf.ln(5)
    pdf.set_x(x=25) 
    #pdf.write(5,  "Some paragraph Some paragraph Some paragraph Some paragraph")
    pdf.write(5,  "Some paragraph Some paragraph Some paragraph Some paragraph ")
    
    pdf.ln(7)
    pdf.set_font('Times', '', 12)
    pdf.set_x(x=30) 
    pdf.write(5, "Some paragraph Some paragraph Some paragraph Some paragraph:-")
    pdf.ln(7)
    pdf.set_x(x=32)
    pdf.set_draw_color(r= 245, g = 35, b=35)
    pdf.set_fill_color(245,35,35)
    pdf.cell(15, 5, '', 1, 0, 'C',True)
    pdf.write(5, 'A<70%      ')
    pdf.set_draw_color(r= 255, g = 255, b=40)
    pdf.set_fill_color(255,255,40)
    pdf.cell(15, 5, '', 1, 0, 'C',True)
    pdf.write(5, '70%<= A <85%    ')
    pdf.set_draw_color(r= 0, g = 176, b=80)
    pdf.set_fill_color(0,176,80)
    pdf.cell(15, 5, '', 1, 0, 'C',True)
    pdf.write(5, 'A => 85%   ')
    pdf.set_draw_color(r= 150, g = 150, b=150)
    pdf.set_fill_color(150,150,150)
    pdf.cell(15, 5, '', 1, 0, 'C',True)
    pdf.write(5, 'any attrubute = 0     ')
    #pdf.set_draw_color(r= 255, g = 172, b=20)
    #pdf.set_fill_color(255,172,20)
    #pdf.cell(15, 5, '', 1, 0, 'C',True)
    #pdf.write(5, 'DVA NC >0%')
    pdf.set_font('Times', 'B', 15)
    pdf.write(5, '*')
    pdf.set_font('Times','', 12)
    pdf.write(5, ' = Indicates special attribute')
    
    
    pdf.ln(10)
    pdf.set_font('Times', 'B', 14)
    pdf.set_x(x=5.2)
    pdf.set_draw_color(222,235,247)
    pdf.set_fill_color(222,235,247)
    pdf.cell(287.6, 9, '' , 1, 0, 'C',True)
    
    
    pdf.set_x(x=30)
    # y1 = pdf.get_y()
    # print(y1)
    pdf.set_text_color(115,115,115)
    pdf.write(9, 'From Date  ')
    
    pdf.set_text_color(60,60,60)
    pdf.write(9, '{}  '.format(FromDate))
    # pdf.write(9, '24-05-2021  ')
    
    pdf.set_text_color(115,115,115)
    pdf.write(9, 'To Date  ')
    
    pdf.set_text_color(60,60,60)
    pdf.write(9, '{}  '.format(ToDate))
    
    pdf.set_text_color(115,115,115)
    pdf.write(9, 'for  ')
    
    
    # VT="HCV"
    # BU="CVBU"
    pdf.set_text_color(60,60,60)
    # pdf.write(9, '{} - {}'.format(VT,BU))
    pdf.write(9, '{} - {}'.format(sys.argv[3],sys.argv[4]))
    
    #pdf.write(9, 'From Date 1-5-2021 to Date {} for {} - {} '.format(Date,sys.argv[3],sys.argv[4]))
    # pdf.set_y(y=y1)
    # pdf.set_x(x=30)
    # pdf.set_text_color(70,70,70)
    # pdf.write(9, '                    {}               {}         LMV - CAR '.format(Date,Date))
    
    pdf.set_fill_color(0, 0, 0)
    pdf.set_text_color(0,0,0)
    
        
    pdf.ln(10)
    
    
   

def Table(data,pdf,DR):
     
    # print(data['DVA_Index'])
    DR=DR
    
    Count_VC = data.VC_Number.count()
    # print(Count_VC)
    
    #Count_Grey = data[(data['Applicability_yes'] == 0 )].Applicability_yes.count()
    # print(Count_Grey)
    #Count_Grey = data[(data['Applicability_yes'] == 0 )].Applicability_yes.count()
    
    #Count_R = data[(data['DVA_Index'] < 80 )].DVA_Index.count()
    Count_R = data[(data['GCI_F'] < 70 )].GCI_F.count()
    # Count_Red = Count_R - Count_Grey
    Count_Red = Count_R 
    
    # Count_Yellow = data[80 <= data['DVA_Index'] < 85].DVA_Index.count()
    # print(Count_Yellow)
    
    #Count_Green = data[(data['DVA_Index'] >= 85 )].DVA_Index.count()
    Count_Green = data[(data['GCI_F'] >= 85 )].GCI_F.count()
    # print(Count_Green)
    

    
    # Count_Yellow = Count_VC - (Count_Red + Count_Green + Count_Grey)
    Count_Yellow = Count_VC - (Count_Red + Count_Green )
    
    
      
    VC_data = pd.DataFrame()
    VC_data['Values'] = list(data['VC_Number'])
    VC_data['Values1'] = list(data['Description'])
    VC_data['Values2'] = list(data['DR_status'])
    VC_data['Values3'] = list(data['Revision'])
    VC_data['Values4'] = list(data['Sequence'])
    VC_data['Values5'] = list(data['No of Usecae Mapped'])
    VC_data['Values6'] = list(data['Applicability_yes'])
    VC_data['Values7'] = list(data['DFQ_Attached'])
    VC_data['Values8'] = list(data['DVA_Index'])
    VC_data['Values9'] = list(data['DFQ NotOk'])
    VC_data['Values10'] = list(data['DVA NC'])
    VC_data['Values11'] = list(data['GCI'])
    VC_data['Values12'] = list(data['Path'])
    VC_data['Values13'] = list(data['Program_scale'])
    VC_data['Values14'] = list(data['DVA_OK'])
    VC_data['Values15'] = list(data['DMU'])
    VC_data['Values16'] = list(data['GCI_F'])
    VC_data['Values17'] = list(data['Owner'])
    
    
    
               
    '''Page frontend(header and footer define at top)'''
    '''
    set_font = font style and size
    ln = space between 2 lines
    set_x = from x it will start to write
    cell = write in box
    write = write the sentence (not in rec box)
    draw_color = border color
    set_fill_color = cell color (IMP- True value should be written in cell to apply color)
    '''
    
   
    pdf.ln(7)
    pdf.set_font('Times', 'B', 12)
    
    y12 = pdf.get_y()
    print(y12)
    if y12 >160:
        pdf.add_page()
    pdf.set_x(x=65) 
    pdf.set_text_color(0, 0, 125) 
    pdf.write(5, '{} VC Release DVA Index Details'.format(DR))
    pdf.set_text_color(105, 105, 105) 
    # pdf.write(5, '  (Total No of VC ={}, Grey={}, Red={}, Yellow={}, Green={})'.format(Count_VC,Count_Grey,Count_Red,Count_Yellow,Count_Green))
    pdf.write(5, '  (Total No of VC ={}, Red={}, Yellow={}, Green={})'.format(Count_VC,Count_Red,Count_Yellow,Count_Green))
    pdf.set_text_color(0, 0,0) 
    pdf.ln(8)
            
    '''Table of VC and DVA Index'''
    pdf.set_x(x=10)
    pdf.set_font('Times', 'B', 10)
    pdf.set_draw_color(r= 100, g = 100, b=100)
    pdf.set_fill_color(200,210,252)
    xPos=pdf.get_x()
    yPos=pdf.get_y()
    pdf.cell(24, 10, 'VC Number', 1, 0, 'C',True)
    pdf.set_xy(xPos + 24 , yPos)
    pdf.cell(76, 10, 'Description', 1, 0, 'C',True)
    pdf.set_xy(xPos + 100 , yPos)
    pdf.cell(10, 10, 'Rev', 1, 0, 'C',True)
    pdf.set_xy(xPos + 110, yPos)
   
    pdf.cell(10, 10, 'Seq', 1, 0, 'C',True)
    pdf.set_xy(xPos + 120, yPos)
    
    # pdf.multi_cell(23, 5, 'Mapped DVA for VC', 1, 'C',True)
    # pdf.set_xy(xPos + 143 , yPos)
    
    pdf.multi_cell(18, 5, 'Program Scale', 1,'C',True)
    pdf.set_xy(xPos + 138, yPos)

    pdf.multi_cell(30, 5, 'Applicability Selected by COC', 1,'C',True)
    pdf.set_xy(xPos + 168 , yPos)
    
    pdf.multi_cell(25, 5, 'DVA Report OK', 1,'C',True)
    pdf.set_xy(xPos + 191, yPos)
  

    # pdf.multi_cell(20, 5, 'DFQ Attached', 1,'C',True)
    # pdf.set_xy(xPos + 193 , yPos)
    
    # pdf.multi_cell(20, 5, 'DFQ Not Ok', 1,'C',True)
    # pdf.set_xy(xPos + 213 , yPos)
    
    pdf.multi_cell(23, 5, 'DVA Index(D)', 1,'C',True)
    pdf.set_xy(xPos + 214 , yPos)

    pdf.cell(26, 10, 'DMU rating (S)', 1,0,'C',True)
    pdf.set_xy(xPos + 240 , yPos)

    pdf.cell(15, 10, 'GCI', 1,0,'C',True)
    pdf.set_xy(xPos + 255 , yPos)

    pdf.cell(23, 10, 'File Link', 1, 0, 'C',True)
    pdf.set_xy(xPos + 273 , yPos)
        
    # pdf.cell(23, 10, 'DVA NC %', 1, 0, 'C',True)
    # pdf.set_xy(xPos + 259 , yPos)
    
    # pdf.multi_cell(34, 5, 'Geometry Conformance Index ', 1,'C',True)
    # pdf.set_xy(xPos + 260 , yPos)
       
    pdf.ln(10)
    
    pdf.set_x(x=10)
    y1 = pdf.get_y()
    x1 = pdf.get_x()
    # print(x1)
    
    
    for i in range(0, len(VC_data)):
        # print('KkkkkkkkkKKKKKKKKKKKKKKKKKkkkkkkkkkkkkkk')
        # print('line 695')
        print(len(VC_data))
        pdf.set_font('Times', 'B', 9)
        pdf.set_draw_color(r= 100, g = 100, b=100)
        if(str(VC_data.Values17.iloc[i])=="Release Vault"):
            vc_rel = str(VC_data.Values.iloc[i]) + "*"
            pdf.cell(24, 8.4, '%s' % (vc_rel), 1, 0, 'C')
        else:
            pdf.cell(24, 8.4, '%s' % (str(VC_data.Values.iloc[i])), 1, 0, 'C')
        pdf.set_font('Times', '', 9)
        # pdf.cell(83, 10, '%s' % (str(VC_data.Values1.iloc[i])), 1, 0, 'L')
        
        
        # des_split=str(VC_data.Values1.iloc[i]).split()
        # print(des_split)
        # listtpstr1=' '.join(map(str, des_split[0:7]))
        # listtpstr2=' '.join(map(str, des_split[7:]))
        
        length=pdf.get_string_width(str(VC_data.Values1.iloc[i]))
        # print(length)
        
        if 110>length>70:
            xPos=pdf.get_x()
            yPos=pdf.get_y()
            pdf.multi_cell(76, 4.2,'%s' % (str(VC_data.Values1.iloc[i])), 1,'L')
            pdf.set_xy(xPos + 76 , yPos)
        elif length<70:
            pdf.cell(76, 8.4, '%s' % (str(VC_data.Values1.iloc[i])), 1, 0, 'L')
        else:
            xPos=pdf.get_x()
            yPos=pdf.get_y()
            pdf.set_font('Times', '', 8)
            pdf.multi_cell(76, 2.8,'%s' % (str(VC_data.Values1.iloc[i])), 1,'L')
            pdf.set_xy(xPos + 76 , yPos)
        
        
        pdf.cell(10, 8.4, '%s' % (str(VC_data.Values3.iloc[i])), 1, 0, 'C')  #NR
        pdf.cell(10, 8.4, '%s' % (str(VC_data.Values4.iloc[i])), 1, 0, 'C')  #seq

        if (str(VC_data.Values13.iloc[i])=='' or str(VC_data.Values13.iloc[i])=='nan' or str(VC_data.Values13.iloc[i])=='Nil' or str(VC_data.Values13.iloc[i])=='(null)' ) :
            pdf.cell(18, 8.4, '%s' % (str('-')), 1, 0, 'C') #prgm scl
        else :
            print('--------------------Program scale---------------------')
            print((VC_data.Values13.iloc[i]))
            print(type((VC_data.Values13.iloc[i])))
            print(int((VC_data.Values13.iloc[i])))
            pdf.cell(18, 8.4, '%s' % (int(VC_data.Values13.iloc[i])), 1, 0, 'C') #prgm scl
        
        xyz = int(VC_data.Values6.iloc[i])
        print('COC applicability xyz == ',xyz)
        if(xyz==0):
            pdf.set_fill_color(150,150,150) #grey
            pdf.cell(30, 8.4, '%s' % (str(VC_data.Values6.iloc[i])), 1, 0, 'C',fill=True)  #COC applicability
        else:
            pdf.cell(30, 8.4, '%s' % (str(VC_data.Values6.iloc[i])), 1, 0, 'C')  #COC applicability

        pdf.cell(23, 8.4, '%s' % (str(VC_data.Values14.iloc[i])), 1, 0, 'C')  #DVA report ok

        pdf.cell(23, 8.4, '%s' % (str(VC_data.Values8.iloc[i])), 1, 0, 'C')  #DVA index

        if(str(VC_data.Values15.iloc[i])=='' or str(VC_data.Values15.iloc[i])=='nan' or str(VC_data.Values15.iloc[i])=='Nil'):
            pdf.cell(26, 8.4, '%s' % (str('-')), 1, 0, 'C')  #DMU rating
        else:
            pdf.cell(26, 8.4, '%s' % (str(VC_data.Values15.iloc[i])), 1, 0, 'C')  #DMU rating


        # if VC_data.Values8.iloc[i] >= 85:
        #     pdf.set_fill_color(0,176,80)
        # elif 80 <= VC_data.Values8.iloc[i] < 85:
        #     pdf.set_fill_color(255,255,40)
        # elif VC_data.Values6.iloc[i] == 0:
        #     pdf.set_fill_color(150,150,150)
        # elif VC_data.Values8.iloc[i] < 80:
        #     pdf.set_fill_color(245,35,35)
        op=int(VC_data.Values16.iloc[i])   #GCI
        
        # if int(VC_data.Values6.iloc[i]) == 0:
        #     pdf.set_fill_color(150,150,150) #grey
        if op >= 85:
            pdf.set_fill_color(0,176,80) #green
        elif 70 <= op < 85:
            pdf.set_fill_color(255,255,40) #yellow
        # elif op == 0 :
        #     if(str(VC_data.Values6.iloc[i])== '' or str(VC_data.Values6.iloc[i]) =='NA' or str(VC_data.Values6.iloc[i])=='Nil'):
        #         if (int(VC_data.Values6.iloc[i])>= 0):
        #             print('-------------Inside RED-------------')
        #             pdf.set_fill_color(245,35,35) #red
        #         else:
        #             pdf.set_fill_color(150,150,150)  #grey
        #     else:
        #         pdf.set_fill_color(150,150,150) #grey

        elif op < 70:
            pdf.set_fill_color(245,35,35)  #red
        
        
        
        pdf.cell(15, 8.4, '%s' % (float(VC_data.Values16.iloc[i])), 1, 0, 'C',fill=True)



            
        # if (str(VC_data.Values11.iloc[i])=='Nil' or str(VC_data.Values11.iloc[i])=='' or str(VC_data.Values11.iloc[i])=='NaN' or str(VC_data.Values11.iloc[i])=='nan'):
        #     #pdf.cell(15, 8.4, '%s' % (str('Nil')), 1, 0, 'C',fill=True)  #GCI
        #     GCI_1=round((VC_data.Values14.iloc[i]/VC_data.Values6.iloc[i])*100*0.5,0)
        #     GCI_2=format((VC_data.Values14.iloc[i]/VC_data.Values6.iloc[i])*100*0.5,".2f")
        #     print('value of i = ',GCI_2)          
        #     print('Line 755')
        #     print('value 14 =',(int(VC_data.Values14.iloc[i])))  #DVA okay
        #     print('type =',type(VC_data.Values14.iloc[i]))
        #     print('value 6 =',(int(VC_data.Values6.iloc[i])))    #COC applicability
        #     print('gci type = ',type(GCI_1))
        #     if int(VC_data.Values14.iloc[i]) == 0 and int(VC_data.Values6.iloc[i])== 0 :
        #         GCI_2=0
        #     print(GCI_1)
        #     print('Abhijit ')
        #     print('################@@@@@@@@@@@@@@@@GCGCGCGCGCGCGCG@@@@@@@@@@@@@@@@@%%%%%%%%%%%%%%%%%')
        #     if GCI_1 >= 85:
        #         pdf.set_fill_color(0,176,80) #green
        #     elif 70 <= GCI_1 < 85:
        #         pdf.set_fill_color(255,255,40) #yellow
        #     # elif GCI_1 == 0:
        #     #     pdf.set_fill_color(150,150,150)  #grey
        #     elif GCI_1 < 70:
        #         pdf.set_fill_color(245,35,35)  #red
            
        #     pdf.cell(15, 8.4, '%s' % (GCI_2), 1, 0, 'C',fill=True)   
        # else :
        #     print('this is excuting')
        #     print(type(VC_data.Values11.iloc[i]))
        #     print(VC_data.Values11.iloc[i])
        #     #GCI_1=format(int(VC_data.Values11.iloc[i]),".2f")
        #     GCI_1=int(VC_data.Values11.iloc[i])
        #     print(GCI_1)
        #     if GCI_1 >= 85:
        #         pdf.set_fill_color(0,176,80) #green
        #     elif 70 <= GCI_1 < 85:
        #         pdf.set_fill_color(255,255,40) #yellow
        #     # elif GCI_1 == 0:
        #     #     pdf.set_fill_color(150,150,150)  #grey
        #     elif GCI_1 < 70:
        #         pdf.set_fill_color(245,35,35)  #red
        #     #pdf.cell(15, 8.4, '%s' % (str(GCI_1)), 1, 0, 'C',fill=True)
        #     pdf.cell(15, 8.4, '%s' % (int(VC_data.Values11.iloc[i])), 1, 0, 'C',fill=True)

        # else :
            
        #     GCI_1=round((VC_data.Values14.iloc[i]/VC_data.Values6.iloc[i])*100*0.5,0)
        #     print('value of i = ',i)          
        #     print('Line 755')
        #     print('value 14 =',int(VC_data.Values14.iloc[i]))
        #     print('gci = ',GCI_1)
        #     print('Abhijit ')
        #     print('################@@@@@@@@@@@@@@@@GCGCGCGCGCGCGCG@@@@@@@@@@@@@@@@@%%%%%%%%%%%%%%%%%')
        #     if GCI_1 >= 85:
        #         pdf.set_fill_color(0,176,80) #green
        #     elif 70 <= GCI_1 < 85:
        #         pdf.set_fill_color(255,255,40) #yellow
        #     # elif GCI_1 == 0:
        #     #     pdf.set_fill_color(150,150,150)  #grey
        #     elif GCI_1 < 70:
        #         pdf.set_fill_color(245,35,35)  #red
        #     pdf.cell(15, 8.4, '%s' % (str(GCI_1)), 1, 0, 'C',fill=True)
            
            

        

        #pdf.cell(23, 8.4, '%s' % (str(VC_data.Values8.iloc[i])), 1, 0, 'C', fill=True)
        
        ''' To color the DVA Index cell depending on condition '''
        '''
        if VC_data.Values8.iloc[i] >= 85:
            pdf.set_font('Times', 'B', 9)
            #pdf.set_text_color(198,0,0)
            pdf.set_fill_color(0,176,80)
            pdf.cell(15, 8.4, '%s' % (str(VC_data.Values8.iloc[i])), 1, 0, 'C', fill=True)
        elif 80 <= VC_data.Values8.iloc[i] < 85:
            pdf.set_font('Times', 'B', 9)
            # pdf.set_text_color(0,175,0)
            pdf.set_fill_color(255,255,40)
            pdf.cell(15, 8.4, '%s' % (str(VC_data.Values8.iloc[i])), 1, 0, 'C', fill=True)
        elif VC_data.Values6.iloc[i] == 0:
            pdf.set_font('Times', 'B', 9)
            # pdf.set_text_color(198,0,0)
            pdf.set_fill_color(150,150,150)
            pdf.cell(15, 8.4, '%s' % (str(VC_data.Values8.iloc[i])), 1, 0, 'C', fill=True)
        elif VC_data.Values8.iloc[i] < 80:
            pdf.set_font('Times', 'B', 9)
            # pdf.set_text_color(0,175,0)
            pdf.set_fill_color(245,35,35)
            pdf.cell(15, 8.4, '%s' % (str(VC_data.Values8.iloc[i])), 1, 0, 'C', fill=True) 
        '''
        #pdf.cell(23, 8.4, '%s' % (str(VC_data.Values9.iloc[i])), 1, 0, 'C')   #DMU rating
                                                                                #CGI
        # pdf.set_fill_color(255,255,255)

        '''   
        if VC_data.Values10.iloc[i] > 0:
            pdf.set_font('Times', 'B', 10)
            # pdf.set_text_color(255,0,0)
            pdf.set_fill_color(255,172,20)
            pdf.cell(23, 8.4, '%s' % (str(VC_data.Values10.iloc[i])), 1, 0, 'C', fill=True)
        else:
            pdf.set_font('Times', 'B', 10)
            # pdf.set_text_color(0,0,0)
            pdf.set_fill_color(255,255,255)
            pdf.cell(23, 8.4, '%s' % (str(VC_data.Values10.iloc[i])), 1, 0, 'C', fill=True)
        '''
        # pdf.cell(18, 8.4, '%s' % (str(VC_data.Values13.iloc[i])), 1, 0, 'C')  <
 
        # pdf.cell(23, 8.4, '%s' % (str(VC_data.Values10.iloc[i])), 1, 0, 'C')
        # pdf.cell(32, 8.4, '%s' % (str(VC_data.Values11.iloc[i])), 1, 0, 'C')
        
        '''Set text to 0 otherwise it will continue the cell color given above'''
        pdf.set_text_color(0, 0, 0)
        pdf.cell(23, 8.4, '' , 1, 2, 'L')
        pdf.set_font('Times', '', 9)
        pdf.set_text_color(0, 0, 255)           
        pdf.write(-8.4, ' Download', '%s' % (str(VC_data.Values12.iloc[i]))) 
        # print(str(VC_data.Values8.iloc[i]))
        pdf.set_text_color(0, 0, 0) 
        # pdf.set_draw_color(r= 0, g = 0, b=0)
        pdf.ln(0)
        pdf.set_x(x=10)      
    


def create_analytics_report(weeklyData_path):

    pdf = PDF(orientation = 'L', unit = 'mm', format='A4')
    pdf.set_top_margin(margin = 35)
    # pdf.set_auto_page_break(True, 20)
    
    SortData(weeklyData_path)
    
    """Add Page"""
    pdf.add_page()
    HeaderPDF(pdf)
    DR0(pdf)
    DR1(pdf) 
    DR2(pdf) 
    DR3(pdf) 
    # DR3P(pdf) 
    # DR4(pdf)
    
    
    
    pdf.ln(10)
    pdf.set_font('Times', '', 12)
    pdf.set_text_color(125, 125, 125)
    pdf.set_x(x=40) 
    pdf.write(5, '----------------------------------------------------------------------End of Report----------------------------------------------------------------------')
    pdf.set_text_color(0, 0, 0)
    
    
        
    '''Saving PDF'''
    # pdf.output('MIS_Dec_2021_CVBU_HCV.pdf', 'F')
    # pdf.output(Output_Pdf_Path,"F")
    pdf.output(sys.argv[2],"F")
    
   


if __name__ == '__main__':

    # weeklyData_path = "/var/www/html/DVA_MIS_Report/DVA_MIS/database/2022/Nov/NC_test/"
    # weeklyData_path = "/var/www/html/DVA_MIS_Report/DVA_MIS/database/2021/Dec/Test/HCV/"
    weeklyData_path = sys.argv[1] 
    create_analytics_report(weeklyData_path)
    


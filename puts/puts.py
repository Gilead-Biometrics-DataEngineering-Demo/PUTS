# -*- coding: utf-8 -*-
######### start of header ######################################
# Program Name:  puts.py
# Author:        Larry Sleeper (lsleeper)
# Description:   PUTS (Python Utility Testing system)
#                Python shared module containing common routines for utility testing
# Category:      Shared Module
# Macros called: 
# Parameter:                
# Usage:         import puts
#                                                                       
# Change History: 
#      2020-02-16  lsleeper - Original programming   
#      2021-03-02  cchowthi - Added puts_diff_html, details and pdf_frm02060
#      2021-10-27  cchowthi - Added docx_frm02060
#
########## end of header ###########################################/


import subprocess
import os
import re
import sys
import shutil
import tempfile
import zipfile
from datetime import datetime
from lxml.html.diff import htmldiff
import fitz
import pandas as pd
import filecmp
import importlib
import math 
from pkg_resources import resource_filename


# initialize puts test                         
def puts_init(python_program_name):
                return (python_program_name.split('.')[0])

# call check_output with command in cmd and 
# send output to file specified in filename
def check_output_to_file(cmd, filename):
                print ('Executing command ', cmd)
                args=cmd.split(' ')
                #stdout=subprocess.check_output(args) 
                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True, executable='/bin/bash')
                stdout = process.communicate()[0] 
                status = process.returncode                                                               
                print ('Writing output to file ', filename)
                with open(filename, 'wb') as filehandle: 
                                header = 'Command:' + cmd + "\n\n"
                                filehandle.write(header.encode())                                                            
                                filehandle.write(stdout)

# check if file contains a string
def file_has_string(filename, search_string):                      
                if os.path.exists(filename):
                                with open(filename) as myfile:
                                                if search_string in myfile.read():
                                                                return True
                return False

# test passed
def test_pass():
                print ('PASS')

# test failed
def test_fail(e):
                print ('FAIL', e)

# remove file if it exists
def remove_file_if_exists(filename):
                if os.path.isfile(filename):
                                os.remove(filename)
                                                                
#def log_write(mode='PUTS Qualification System',returncode=0,status='Failure',ur_num=0,ur_txt='N/A'):
  #log= open('/biometrics/system_programming/qualification/global_programs/PUTS_Testing_Output_' + Current_Date + '.txt',"a+")
  #log.write(mode + "," + utility + ", returncode:" + str(returncode) + "," + status + "," + str(ur_num) + "," + ur_txt + ", Date Performed: " + str(Current_Date) + ", " + str(cwd) + "\n")
  #log.close()
                                                                
# check if two dataframes are the same
def puts_diff_dataframe(df1, df2):
                if df2.equals(df1):
                                return True
                else:
                                return False
                                
# check if two dataframes are the same
def puts_diff_excel(xls1, xls2, worksheet):
                df1=pd.read_excel(xls1, sheet_name=worksheet)
                df2=pd.read_excel(xls2, sheet_name=worksheet)
                return puts_diff_dataframe(df1, df2)
                
# check if two files are the same
def puts_diff_file(f1, f2):
                return filecmp.cmp(f1, f2, shallow=True)
                
# check if two PDF files have the same number of pages and the same content.
def puts_diff_pdf(pdf1, pdf2):
                
                pdf1 = fitz.open(pdf1)
                pdf2 = fitz.open(pdf2)
                pdf1_pages = pdf1.pageCount
                pdf2_pages = pdf2.pageCount
                
                if pdf1_pages != pdf2_pages:
                                #print (pdf1, 'has different number of pages than', pdf2)
                                return False
                
                #print (pdf1, 'has same number of pages than', pdf2)
                
                for page_number in range(pdf1_pages):
                                page1 = pdf1.loadPage(page_number)
                                page2 = pdf2.loadPage(page_number)
                                page1_content = page1.getText("text")
                                page2_content = page2.getText("text")
                                if page1_content != page2_content:
                                                #print (pdf1, 'has different content on page', page_number, 'than', pdf2)
                                                return False
                                                
                #print (pdf1, 'and', pdf2, 'have the same content')                           
                return True

# Check if two html/xml files are the same. Compare by tag. Can provide list of strings in ignoreifcontains
# to ignore any particular tags (ex. 'Date/time created')
def puts_diff_html(f1, f2, ignoreifcontains=[]):
    
    with open(f1, "r", encoding='utf-8') as f:
        s1= f.read()
    with open(f2, "r", encoding='utf-8') as f:
        s2= f.read()
    set1 = set([x for x in re.split("<.*?>", s1) if not any(y in str(x) for y in ignoreifcontains)])
    set2 = set([x for x in re.split("<.*?>", s2) if not any(y in str(x) for y in ignoreifcontains)])
    return set1==set2

# Retrieve metadata from current working directory and python program    
class details:
    utility = re.sub(r'.*?([^\/]+)\/qualification\/.*',r'\1',os.getcwd())
    path = re.sub(r'(.*?[^\/]+)\/qualification\/.*',r'\1',os.getcwd())
    qcpid = os.environ.get('USER')
    try:
        sys.path.insert(0, path)
        module = importlib.import_module(utility)
        pdid = module.__author__
        pdate = module.__date__
        sys.path.remove(path) 
    except:
        pdid = ''
        pdate = ''
    qcdate = datetime.today().strftime('%Y-%m-%d')
    qcmethod = 'UT'
                
# Create representation of GT-27043A from test results
def pdf_frm02060(session):
    
    global rects, heads  
    
    # Initiate new page
    def new_page():
        Page = newPDF.newPage()
        Page.setMediaBox(fitz.Rect(0,0,800,600))
        page_head(Page)
        return Page
    
    # Draw new block
    def block(Page, d, border=True, bwidth=0.3, bcolor=(0, 0, 0), fcolor=(1, 1, 1), image='', text='', font='Helvetica', size=8, xpad=0, ypad=4, talign=0):
        r = fitz.Rect(d)        
        img = Page.newShape()
        img.drawRect(r)
        if border:
            img.finish(width=bwidth, color=bcolor, fill=fcolor)
       
        if text:
            img.insertTextbox(fitz.Rect(d[0] + xpad, d[1] + ypad, d[2] - xpad, d[3] - ypad), text, fontsize=size, fontname=font, align=talign)
        img.commit()
        if image:
            with open(image, 'rb') as f:
                logo = f.read()   
            Page.insertImage(fitz.Rect(d[0] + xpad, d[1] + ypad, d[2] - xpad, d[3] - ypad), stream=logo)  
                                
    # Insert page header    
    def page_head(Page):     
        logo = resource_filename(__name__, 'suppl/gilead.png')  
        block(Page, (40, 30, 250, 65), image=logo, ypad=2)
        block(Page, (250, 30, 500, 65), text='Utility and System Program Acceptance', font='Helvetica-Bold', size=12, ypad=8, talign=1)
        block(Page, (500, 30, 760, 65), text='Representation of GT-27043A (2.0)', font='Helvetica-Bold', size=12, ypad=8, talign=1)  

    # Insert table header
    def table_head(Page):        
        global rects, heads           
        for h, d in zip(heads,rects):    
            block(Page, d, text=h, fcolor=(0.75, 0.75, 0.75), font='Helvetica-Bold', talign=1)   
   
    # Define variables
    pno = 0
    rects = [(40, 100, 110, 120), #Utility
                (110, 100, 150, 120), #UR ID
                (150, 100, 300, 120), #User Requirement
                (300, 100, 360, 120), #PD ID
                (360, 100, 420, 120), #Program Date
                (420, 100, 460, 120), #QCP ID
                (460, 100, 520, 120), #QC Date
                (520, 100, 560, 120), #Method
                (560, 100, 610, 120), #Pass / Fail
                (610, 100, 760, 120)] #Exception Comment  
    heads = ['Utility','UR ID','User Requirement','PD ID','Program Date','QCP ID','QC Date','Method','Pass / Fail','Exception Comment']
    
    # Open new PDF and create page        
    newPDF = fitz.open()
    Page = new_page()
    table_head(Page)
    y0 = rects[0][3]  
    height = 10

    # Insert session data into table
    alldata = [[details.utility, 
                item.urid, 
                item.user_requirement, 
                details.pdid,
                details.pdate,
                details.qcpid, 
                details.qcdate, 
                details.qcmethod,
                item.Pass_fail,
                ''] for item in session]
                                            
    for datarow in alldata:
        maxfactor = max([(fitz.getTextLength(h, fontname="Helvetica-Bold", fontsize=12)-4)/(d[2]-d[0]-4) for h, d in zip(datarow,rects)]) + 1
        for h, d in zip(datarow,rects):   
            if y0 + height * math.ceil(maxfactor) >= 540: 
                # Initiate new page/table if towards end of page
                y0 = d[3]
                pno += 1
                Page = new_page()
                table_head(Page)
            d = (d[0], y0 ,d[2], y0 + height * math.ceil(maxfactor) + 4)  
            block(Page, d, text=h, xpad=2, ypad=2)
        y0 += height * math.ceil(maxfactor)

    # Add signature page
    Page = new_page()
    pno += 1

    block(Page, (40, 100, 800, 130), border=False, text="Release Version: _________________ (for system programs N/A may be appropriate)", font='Helvetica-Bold', size=12)
    block(Page, (40, 140, 800, 170), border=False, text="Acceptance:", font='Helvetica-Bold', size=12)
    block(Page, (40, 170, 800, 190), border=False, 
            text="TL Name (print):___________________________________  Signature: ____________________________________________  Date: _____________________", 
            font='Helvetica-Bold', size=9)
    block(Page, (40, 210, 800, 240), border=False, text="Deployment:", font='Helvetica-Bold', size=12)
    block(Page, (40, 240, 800, 260), border=False, 
            text="GL Name (print):___________________________________  Signature: ____________________________________________  Date: _____________________", 
            font='Helvetica-Bold', size=9)   
    
    # Add footers to each page (program name and page x of y)
    for i in range(pno + 1):
        Page = newPDF[i] 
        block(Page, (660, 560, 760, 580), border=False, text="Page " + str(i+1) + " of " + str(pno+1), talign=2)
        block(Page, (40, 560, 100, 580), border=False, text=alldata[0][0])
    
    # Save to current directory    
    newPDF.save(os.getcwd() + "/FRM-02060 Utility and System Program Acceptance " + alldata[0][0] + ".pdf")   

# Complete GT-27043A template from test results
def docx_frm02060(session):
    
    alldata = [[details.utility, 
                item.urid, 
                item.user_requirement, 
                details.pdid, 
                details.pdate, 
                details.qcpid, 
                details.qcdate, 
                details.qcmethod, 
                item.Pass_fail, 
                ''] for item in session]

    cellcode = '<w:tc><w:tcPr><w:tcW w:w="COLWIDTH" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="auto"/></w:tcBorders><w:vAlign w:val="center"/></w:tcPr><w:p w:rsidR="00A705EA" w:rsidRDefault="00A705EA" w:rsidP="00D36AFB"><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>TEXT</w:t></w:r></w:p></w:tc>'
    template = resource_filename(__name__, 'suppl/FRM-02060 Utility and System Program Acceptance Template.docx')
    filename = os.getcwd() + '/FRM-02060 Utility and System Program Acceptance ' + alldata[0][0] + '.docx'
    os.system('cp ' + template.replace(' ', '\\ ') + ' ' + filename.replace(' ', '\\ '))
    docx = zipfile.ZipFile(template)
    docdata = docx.read('word/document.xml').decode('utf-8')
    tblre = re.compile('<w:tbl(?: .*?)?>.*?</w:tbl>')
    trre = re.compile('<w:tr(?: .*?)?>.*?</w:tr>')
    tcre = re.compile('<w:tc(?: .*?)?>.*?</w:tc>')
    tre = re.compile('<w:t(?: .*?)?>(.*?)</w:t>')
    datasets = {}
    tbl = 0
    widths = []
    header = ''
    
    # Clear rows
    for tble in tblre.findall(docdata):
        tbl = tbl + 1
        if ' Comment' in tble:
            newtble = tble
            for row in trre.findall(tble):
                if header:
                    newtble = newtble.replace(row, '')
                if ' Comment' in row:
                    header = row
                    for cell in tcre.findall(row):
                        width = re.sub('.*?w:tcW w:w="(\\d+)".*', '\\1', cell)
                        widths.append(width)

            docdata = docdata.replace(tble, newtble)

    # Add rows from test data
    table = header
    for datarow in alldata:
        table += '<w:tr w:rsidR="00A705EA" w:rsidRPr="00D26556" w:rsidTr="00E26658"><w:trPr><w:trHeight w:val="288"/></w:trPr>'
        table += ''.join(cellcode.replace('COLWIDTH', x).replace('TEXT', y).replace('\n', '') for x, y in zip(widths, datarow))
        table += '</w:tr>'

    docdata = docdata.replace(header, table)
    tmpDir = tempfile.mkdtemp()
    docx.extractall(tmpDir)
    with open(os.path.join(tmpDir, 'word/document.xml'), 'w') as (f):
        f.write(docdata)
    filenames = docx.namelist()
    with zipfile.ZipFile(filename, 'w') as (docxout):
        for f in filenames:
            docxout.write(os.path.join(tmpDir, f), f)

    shutil.rmtree(tmpDir)      
      
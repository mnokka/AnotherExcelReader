#encoding=latin1

# POC tool to read Excel using Python
# Data will be used to create subtasks / add attachments to Jira main issues
# Created either via this tool or Excel import plugin
# 
#
#
# Author mika.nokka1@gmail.com for Ambientia
#
# TODO 
# Use Pandas instead?
#
#from __future__ import unicode_literals

import openpyxl 
import sys, logging
import argparse
#import re
from collections import defaultdict
from CreateIssue import Authenticate  # no need to use as external command
from CreateIssue import DoJIRAStuff, CreateSubTask
from CreateIssue import CreateIssue 
import glob

 
__version__ = "0.2.1394"


logging.basicConfig(level=logging.DEBUG) # IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out



def main(argv):
    
    JIRASERVICE=""
    JIRAPROJECT=""
    PSWD=''
    USER=''
  
    logging.debug ("--Python starting Excel reading --") 

 
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com for Ambientia
    
    USAGE:
    -filepath  | -p <Path to Excel file directory>
    -filename   | -n <Excel filename>

    """.format(__version__,sys.argv[0]))

    parser.add_argument('-f','--filepath', help='<Path to Excel file directory>')
    parser.add_argument('-n','--filename', help='<Excel filename>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
    
    parser.add_argument('-w','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    parser.add_argument('-s','--service', help='<JIRA service>')
    parser.add_argument('-p','--project', help='<JIRA project>')
   
        
    args = parser.parse_args()
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
           
    filepath = args.filepath or ''
    filename = args.filename or ''
    
    JIRASERVICE = args.service or ''
    JIRAPROJECT = args.project or ''
    PSWD= args.password or ''
    USER= args.user or ''
    
    # quick old-school way to check needed parameters
    if (filepath=='' or  filename=='' or JIRASERVICE=='' or  JIRAPROJECT==''  or PSWD=='' or USER=='' ):
        parser.print_help()
        sys.exit(2)
        


    Parse(filepath, filename,JIRASERVICE,JIRAPROJECT,PSWD,USER)


############################################################################################################################################
# Parse excel and create dictionary of
# 1) Jira main issue data
# 2) Jira subtask(s) (remark(s)) data for main issue
# 3) Info of attachment for main issue (to be added using inhouse tooling
#  
#NOTE: Uses hardcoded sheet/column value

def Parse(filepath, filename,JIRASERVICE,JIRAPROJECT,PSWD,USER):
    logging.debug ("Filepath: %s     Filename:%s" %(filepath ,filename))
    files=filepath+"/"+filename
    logging.debug ("File:{0}".format(files))
   
    Issues=defaultdict(dict) 

    MainSheet="general_report" 
    wb= openpyxl.load_workbook(files)
    #types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    #sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    CurrentSheet=wb[MainSheet] 
    #logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    #logging.debug ("First row:{0}".format(CurrentSheet['A4'].value))

    ########################################
    #CONFIGURATIONS AND EXCEL COLUMN MAPPINGS
    DATASTARTSROW=5 # data section starting line
    C=3 #SUMMARY
    D=4 #Issue Type
    E=5 #Status Always "Open"    
    G=7 #ResponsibleNW
    H=8 #Creator
    I=9 #Inspection date --> Original Created date in Jira Changed as Inspection Date
    #K=11 #LINKED_ISSUES 
    M=13 #Shipnumber
    P=16 #PerformerNW
    #Q=17 #ResponsibleNW
    #U=20 #Responsible Phone Number --> Not taken, field just exists in Jira
    S=19 #DepartmentNW
    V=22 #Deck
    W=23 #Block
    X=24 # Firezone
    #AA=26 #DeckNW
    
    
    
    
    
   

    
    #for cell in CurrentSheet['A']:
    #    logging.debug  ("Row value:{0}".format(cell.value))
  
    
    ##############################################################################################
    #Go through main excel sheet for main issue keys (and contents findings)
    # Create dictionary structure (remarks)
    # NOTE: Uses hardcoded sheet/column values
    # NOTE: As this handles first sheet, using used row/cell reading (buggy, works only for first sheet) 
    #
    i=DATASTARTSROW # brute force row indexing
    for row in CurrentSheet[('B{}:B{}'.format(DATASTARTSROW,CurrentSheet.max_row))]:  # go trough all column B (KEY) rows
        for mycell in row:
            KEY=mycell.value
            logging.debug("ROW:{0} Original ID:{1}".format(i,mycell.value))
            Issues[KEY]={} # add to dictionary as master key (KEY)
            
            #Just hardocode operations, POC is one off
            #LINKED_ISSUES=(CurrentSheet.cell(row=i, column=K).value) #NOTE THIS APPROACH GOES ALWAYS TO THE FIRST SHEET
            #logging.debug("Attachment:{0}".format((CurrentSheet.cell(row=i, column=K).value))) # for the same row, show also column K (LINKED_ISSUES) values
            #Issues[KEY]["LINKED_ISSUES"] = LINKED_ISSUES
            
            SUMMARY=(CurrentSheet.cell(row=i, column=C).value)
            if not SUMMARY:
                SUMMARY="Summary for this task has not been defined"
            Issues[KEY]["SUMMARY"] = SUMMARY
            
            ISSUE_TYPE=(CurrentSheet.cell(row=i, column=D).value)
            Issues[KEY]["ISSUE_TYPE"] = ISSUE_TYPE
            
            STATUS=(CurrentSheet.cell(row=i, column=E).value)
            Issues[KEY]["STATUS"] = STATUS
            
            RESPONSIBLE=(CurrentSheet.cell(row=i, column=G).value)
            Issues[KEY]["RESPONSIBLE"] = RESPONSIBLE.encode('utf-8')
            
            #REPORTER=(CurrentSheet.cell(row=i, column=G).value)
            #Issues[KEY]["REPORTER"] = REPORTER
            
            
            CREATOR=(CurrentSheet.cell(row=i, column=H).value)
            Issues[KEY]["CREATOR"] = CREATOR
            
            CREATED=(CurrentSheet.cell(row=i, column=I).value) #Inspection date
            Issues[KEY]["CREATED"] = CREATED
            
            SHIP=(CurrentSheet.cell(row=i, column=M).value)
            Issues[KEY]["SHIP"] = SHIP
            
            PERFORMER=(CurrentSheet.cell(row=i, column=P).value)
            Issues[KEY]["PERFORMER"] = PERFORMER.encode('utf-8')
            
              
            #RESPHONE=(CurrentSheet.cell(row=i, column=U).value)
            #Issues[KEY]["RESPHONE"] = RESPHONE
            
            DEPARTMENT=(CurrentSheet.cell(row=i, column=S).value)
            Issues[KEY]["DEPARTMENT"] = DEPARTMENT
            
            DECK=(CurrentSheet.cell(row=i, column=V).value)
            Issues[KEY]["DECK"] = DECK
            
            BLOCK=(CurrentSheet.cell(row=i, column=W).value)
            Issues[KEY]["BLOCK"] = BLOCK
            
            FIREZONE=(CurrentSheet.cell(row=i, column=X).value)
            Issues[KEY]["FIREZONE"] = FIREZONE
            
            
            
            
            
            #Create sub dictionary for possible subtasks (to be used later)
            Issues[KEY]["REMARKS"]={}
            
            logging.debug("---------------------------------------------------")
            i=i+1
    #print Issues
    print Issues.items() 
    
    #key=18503 # check if this key exists
    #if key in Issues:
    #    print "EXISTS"
    #else:
    #    print "NOT THERE"
    #for key, value in Issues.iteritems() :
    #    print key, value

    ############################################################################################################################
    # Check any remarks (subtasks) for main issue
    # NOTE: Uses hardcoded sheet/column values
    #
    #removed currently dfue excel changes

    Authenticate(JIRASERVICE,PSWD,USER)
    jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)

    #create main issues
    for key, value in Issues.iteritems() :
        print "ORIGINAL ISSUE KEY:{0}\nVALUE:{1}".format(key, value)
        #print "1)LINKED_ISSUES:{0}".format(Issues[key]["LINKED_ISSUES"])
        #print "2)REPORTER:{0}".format(Issues[key]["REPORTER"])
        print "3)REMARKS:{0}".format(Issues[key]["REMARKS"])
        print "4)SUMMARY:{0}".format((Issues[key]["SUMMARY"]).encode('utf-8'))
        print "5)ISSUE_TYPE:{0}".format((Issues[key]["ISSUE_TYPE"]).encode('utf-8'))    
        print "6)STATUS:{0}".format(Issues[key]["STATUS"])  
        print "7)CREATOR:{0}".format(Issues[key]["CREATOR"])  
        print "8)CREATED:{0}".format(Issues[key]["CREATED"]) 
        print "9)SHIP:{0}".format(Issues[key]["SHIP"]) 
        print "10)PERFORMER:{0}".format(Issues[key]["PERFORMER"]) #.encode('utf8'))    
        print "11)RESPONSIBLE:{0}".format(Issues[key]["RESPONSIBLE"]) #.encode('utf8'))         
        #print "12)RESPHONE:{0}".format(Issues[key]["RESPHONE"])     
        print "13)DEPARTMENT:{0}".format(Issues[key]["DEPARTMENT"])      
        print "14)BLOCK:{0}".format(Issues[key]["BLOCK"])     
        #print "15)CRONO:{0}".format(Issues[key]["CRONO"])          
        print "16)DECK:{0}".format(Issues[key]["DECK"])      
        print "16)FIREZONE:{0}".format(Issues[key]["FIREZONE"])
   
        JIRADESCRIPTION="Inspection Report"
        JIRASUMMARY=(Issues[key]["SUMMARY"]).encode('utf-8')          
        JIRASUMMARY=JIRASUMMARY.replace("\n", " ") # Perl used to have chomp, this was only Python way to do this
        JIRASUMMARY=JIRASUMMARY[:254] ## summary max length is 255
        KEY=key
        #REPORTER=Issues[key]["REPORTER"]
        CREATOR=Issues[key]["CREATOR"]
        CREATED=Issues[key]["CREATED"] # 30.1.2018  9:32:15 fromat from excel
        SHIP=Issues[key]["SHIP"]
        RESPONSIBLE=Issues[key]["RESPONSIBLE"]
        PERFORMER=Issues[key]["PERFORMER"]
        BLOCK=Issues[key]["BLOCK"]
        DEPARTMENT=Issues[key]["DEPARTMENT"]
        DECK=Issues[key]["DECK"]
        DECK=DECK.encode('utf-8') 
        ISSUETYPE=Issues[key]["ISSUE_TYPE"]

        
        # ISO 8601 conversion to Exceli time
        time2=CREATED.strftime("%Y-%m-%dT%H:%M:%S.000-0300")  #-0300 is UTC delta to Finland, 000 just keeps Jira happy
        print "CREATED ISOFORMAT TIME2:{0}".format(time2)
        CREATED=time2
        INSPECTED=CREATED # just reusing value

        
   
        IssueID=CreateIssue(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION,KEY,CREATOR,CREATED,INSPECTED,SHIP,PERFORMER,RESPONSIBLE,BLOCK,DEPARTMENT,DECK,ISSUETYPE)
        print "Issue:{0}".format(IssueID)
        #print "IssueKey:{0}".format(IssueID.key)
        
        filesx=filepath+"/*{0}*".format(key)
        print "filesx:{0}".format(filesx)
        
        
        attachments=glob.glob("{0}".format(filesx))
        if (len(attachments) > 0): # if any attachment with key embedded to name found
            print "Found attachments for key:{0}".format(IssueID)
            print "Found these:{0}".format(attachments)
            for item in attachments: # add them all
                jira.add_attachment(issue=IssueID, attachment=attachments[0])
                print "Attachment:{0} added".format(item)
        else:
            print "NO attachments  found for key:{0}".format(IssueID)
        
        
        Remarks=Issues[key]["REMARKS"] # take a copy of remarks and use it
        print "-------------------------------------------------------------------------"
        PARENT=IssueID
        #create subtask(s) under one parent
        for subkey , subvalue in Remarks.iteritems():
            #print subkey, subvalue
            print "    Remark key:{0}".format(subkey)
            print "    A) DECK:{0}".format(Remarks[subkey]["DECK"])
            print "    B) BLOCK:{0}".format(Remarks[subkey]["BLOCK"])
            print "    C) NUMBEROF:{0}".format(Remarks[subkey]["NUMBEROF"])
            JIRASUMMARY="Subtask for BGR:{0}".format(subkey)
            JIRADESCRIPTION="BLOCK:{0}    DECK:{1}".format(Remarks[subkey]["BLOCK"],Remarks[subkey]["DECK"])
            SubIssueID=CreateSubTask(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION,PARENT)
            print "Subtask:{0}".format(SubIssueID)
            
        print "*************************************************************************"
        

 
  

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 
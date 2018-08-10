#encoding=latin1

# POC tool to add attachments to existing Jira issues
# Attachment will have Key field in their name. This field exists in Jira issue, thus one can search matching Jira issue for operation
# Forked and modified from master version (which creates issues With attachment in fthe first round)
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
import re
import os
import time


start = time.clock()
__version__ = "0.2.1394"


logging.basicConfig(level=logging.DEBUG) # IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out



def main(argv):
    
    JIRASERVICE=""
    JIRAPROJECT=""
    PSWD=''
    USER=''
  
    logging.debug ("--Python starting checking Jira issues for attachemnt adding --") 

 
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com for Ambientia
    
    USAGE:
    -filepath  | -p <Path to Excel file directory>
    -filename   | -n <Excel filename>

    """.format(__version__,sys.argv[0]))

    parser.add_argument('-f','--filepath', help='<Path to attachment directory>')
    parser.add_argument('-q','--excelfilepath', help='<Path to excel directory>')
    parser.add_argument('-n','--filename', help='<Main tasks Excel filename>')
    parser.add_argument('-m','--subfilename', help='<Subtasks Excel filename>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
    
    parser.add_argument('-w','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    parser.add_argument('-s','--service', help='<JIRA service>')
    parser.add_argument('-p','--project', help='<JIRA project>')
    parser.add_argument('-z','--rename', help='<rename files>') #adhoc operation activation
   
        
    args = parser.parse_args()
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
           
    filepath = args.filepath or ''
    excelfilepath = args.excelfilepath or ''
    filename = args.filename or ''
    subfilename=args.subfilename or ''
    
    JIRASERVICE = args.service or ''
    JIRAPROJECT = args.project or ''
    PSWD= args.password or ''
    USER= args.user or ''
    RENAME= args.rename or ''
    
    # quick old-school way to check needed parameters
    if (filepath=='' or  JIRASERVICE=='' or  JIRAPROJECT==''  or PSWD=='' or USER=='' or subfilename=='' or excelfilepath=='' or filename==''):
        parser.print_help()
        sys.exit(2)
        


    Parse(filepath,JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename)


############################################################################################################################################
# Parse attachment files and add to matching Jira issue
#

#NOTE: Uses hardcoded sheet/column value

def Parse(filepath, JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename):
    logging.debug ("Attachment Filepath: %s   " %(filepath))
    files=excelfilepath+"/"+filename
    logging.debug ("File:{0}".format(files))
   
    Issues=defaultdict(dict) 
   
    #main excel definitions
    MainSheet="general_report" 
    wb= openpyxl.load_workbook(files)
    #types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    #sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    CurrentSheet=wb[MainSheet] 
    #logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    #logging.debug ("First row:{0}".format(CurrentSheet['A4'].value))

   
    #subtasks excel definitions
    logging.debug ("ExcelFilepath: %s     ExcelFilename:%s" %(excelfilepath ,subfilename))
    subfiles=excelfilepath+"/"+subfilename
    logging.debug ("SubFiles:{0}".format(subfiles))
   
    
    SubMainSheet="general_report" 
    subwb= openpyxl.load_workbook(subfiles)
    #types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    #sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    SubCurrentSheet=subwb[SubMainSheet] 
    #logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    #logging.debug ("First row:{0}".format(CurrentSheet['A4'].value))
   
   
   
    

    
    
    
    ########################################
    #CONFIGURATIONS AND EXCEL COLUMN MAPPINGS, both main and subtask excel
    DATASTARTSROW=5 # data section starting line MAIN TASKS EXCEL
    DATASTARTSROWSUB=5 # data section starting line SUB TASKS EXCEL
    C=3 #SUMMARY
    D=4 #Issue Type
    E=5 #Status Always "Open"    
    G=7 #ResponsibleNW
    H=8 #Creator
    I=9 #Inspection date --> Original Created date in Jira Changed as Inspection Date
    J=10 # Subtask TASK-ID
    K=11 #system number, subtasks excel 
    M=13 #Shipnumber
    N=14 #system number
    P=16 #PerformerNW
    Q=17 #Performer, subtask excel
    R=18 #Responsible ,subtask excel
    #U=20 #Responsible Phone Number --> Not taken, field just exists in Jira
    S=19 #DepartmentNW
    V=22 #Deck
    W=23 #Block
    X=24 # Firezone
    AA=27 #Subtask DeckNW
   

    
   
    #print Issues.items() 
    
    #key=18503 # check if this key exists
    #if key in Issues:
    #    print "EXISTS"
    #else:
    #    print "NOT THERE"
    #for key, value in Issues.iteritems() :
    #    print key, value

    #
               
    


    ##########################################################################################################################
    
    
    
    Authenticate(JIRASERVICE,PSWD,USER)
    jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)

    
    #Deactivated renaming command     
    attachments=glob.glob("{0}/*/*".format(filepath))
    if (len(attachments) > 0): # if any attachment with key embedded to name found
        
        # RENAME ATTACHMENT FILES USING DIRECTORY ID NUMBER
        # FILE Attachment 1016:..\..\MIKAN_TYO\ASIAKKAAT\Meyer\tsp\04_Attachment Remarks\394_3429854\IMG_0330.JPG RENAMING -->
        # Attachment 1016:..\..\MIKAN_TYO\ASIAKKAAT\Meyer\tsp\04_Attachment Remarks\394_3429854\3429854_IMG_0330.JPG
        if (RENAME):
            i=1
            for item in attachments: # add them all
                #jira.add_attachment(issue=IssueID, attachment=attachments[0])
                print "*****************************************"
                print "Attachment {0}:{1}".format(i,item)
                regex = r"(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                regex2=r"(.*?)(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                #test_str = "\\394_3428553\\"
                match = re.search(regex, item)
                match2 = re.search(regex2, item)
                hit=match.group(4)
                origname=match.group(6)
                path=match2.group(1)+match2.group(2)+match2.group(3)+match2.group(4)+match2.group(5)
                print "Attachment remark ID:{0}".format(hit)
                print "Original name:{0}".format(origname)
                newname=hit+"_"+origname
                print "New name:{0}".format(newname)
                print "Path: {0}".format(path)
                newfile=path+"\\"+newname
                print "GOING TO DO RENAMING:{0} ---->  {1}".format(item,newfile)
                # removed for safety   os.rename(item, newfile)
                print "Done!!!"
                i=i+1
        else:
            print "--> Renaming bypassed"
        
        i=1
        for item in attachments: # add them all
            print "Attachment {0}:{1}".format(i,item)
            i=i+1
            
            
        #Find remark's original parent ID using 1) remark ID in the file name 2) remark excel 
        
        print "--> SUBTASK EXCEL: {0}".format(subfilename)
        
        
    ### MAIN EXCEL ###########################################################################################
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
            # ISO 8601 conversion to Exceli time
            time2=CREATED.strftime("%Y-%m-%dT%H:%M:%S.000-0300")  #-0300 is UTC delta to Finland, 000 just keeps Jira happy
            print "CREATED ISOFORMAT TIME2:{0}".format(time2)
            CREATED=time2
            INSPECTED=CREATED # just reusing value
            Issues[KEY]["INSPECTED"] = INSPECTED
            
            
            SHIP=(CurrentSheet.cell(row=i, column=M).value)
            Issues[KEY]["SHIP"] = SHIP
            
            PERFORMER=(CurrentSheet.cell(row=i, column=P).value)
            Issues[KEY]["PERFORMER"] = PERFORMER # .encode('utf-8')
            
              
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
            
                
            SYSTEMNUMBER=(CurrentSheet.cell(row=i, column=N).value)
            Issues[KEY]["SYSTEMNUMBER"] = SYSTEMNUMBER
            
            
            
            
            #Create sub dictionary for possible subtasks (to be used later)
            Issues[KEY]["REMARKS"]={}
            
            logging.debug("---------------------------------------------------")
            i=i+1
    #print Issues
    #print Issues.items() 
    
    #key=18503 # check if this key exists
    #if key in Issues:
    #    print "EXISTS"
    #else:
    #    print "NOT THERE"
    #for key, value in Issues.iteritems() :
    #    print key, value
   
        
    #######REMARK EXCEL #####################################################################################################################
    # Check any remarks (subtasks) for main issue
    # NOTE: Uses hardcoded sheet/column values
    #
    #removed currently dfue excel changes

    i=DATASTARTSROWSUB # brute force row indexing
    for row in SubCurrentSheet[('B{}:B{}'.format(DATASTARTSROWSUB,SubCurrentSheet.max_row))]:  # go trough all column B (KEY) rows
        for submycell in row:
            PARENTKEY=submycell.value
            logging.debug("SUBROW:{0} Original PARENT ID:{1}".format(i,PARENTKEY))
            #Issues[KEY]={} # add to dictionary as master key (KEY)
            
            #Just hardocode operations, POC is one off
            #LINKED_ISSUES=(CurrentSheet.cell(row=i, column=K).value) #NOTE THIS APPROACH GOES ALWAYS TO THE FIRST SHEET
            #logging.debug("Attachment:{0}".format((CurrentSheet.cell(row=i, column=K).value))) # for the same row, show also column K (LINKED_ISSUES) values
            #Issues[KEY]["LINKED_ISSUES"] = LINKED_ISSUES
            if PARENTKEY in Issues:
                print "Subtask has a known parent {0}".format(PARENTKEY)
                #REMARKKEY=SubCurrentSheet['J{0}'.format(i)].value  # column J holds Task-ID NW
                REMARKKEY=(SubCurrentSheet.cell(row=i, column=J).value)
                print "REMARKKEY:{0}".format(REMARKKEY)
                #Issues[KEY]["REMARKS"]={}
                Issues[PARENTKEY]["REMARKS"][REMARKKEY] = {}
                
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["REMARKKEY"] = REMARKKEY
                
                # Just hardcode operattions, POC is one off
                #DECK=SubCurrentSheet['AA{0}'.format(i)].value  # column AA holds DECK
                SUBDECK=(SubCurrentSheet.cell(row=i, column=AA).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["DECK"] = SUBDECK
                
                SUBBLOCK=(SubCurrentSheet.cell(row=i, column=X).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["BLOCK"] = SUBBLOCK
                
                SUBPERFORMER=(SubCurrentSheet.cell(row=i, column=Q).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["PERFORMER"] = SUBPERFORMER
                
                SUBRESPONSIBLE=(SubCurrentSheet.cell(row=i, column=R).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["RESPONSIBLE"] = SUBRESPONSIBLE
                
                SUBDEPARTMENT=(SubCurrentSheet.cell(row=i, column=W).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["DEPARTMENT"] = SUBDEPARTMENT
                
                SUBISSUETYPE=(SubCurrentSheet.cell(row=i, column=D).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["ISSUETYPE"] = SUBISSUETYPE
                
                SUBSYSTEMNUMBER=(SubCurrentSheet.cell(row=i, column=N).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SYSTEMNUMBER"] = SUBSYSTEMNUMBER
                
                SUBSUMMARY=(SubCurrentSheet.cell(row=i, column=C).value)
                if not SUBSUMMARY:
                    SUBSUMMARY="Summary for this subtask has not been defined"
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SUMMARY"] = SUBSUMMARY
                
                SUBCREATED=(SubCurrentSheet.cell(row=i, column=I).value) #Inspection date
                # ISO 8601 conversion to Exceli time
                subtime2=SUBCREATED.strftime("%Y-%m-%dT%H:%M:%S.000-0300")  #-0300 is UTC delta to Finland, 000 just keeps Jira happy
                print "CREATED SUBTASK ISOFORMAT TIME2:{0}".format(subtime2)
                SUBCREATED=subtime2
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SUBCREATED"] = SUBCREATED
                
                JIRASUBDESCRIPTION="Remark for Inspection Report"
                SUBTASKID=REMARKKEY
            
            else:
                    print "ERROR: Unknown parent found"
            print "----------------------------------"
            i=i+1
           
       
         
    # Now orignal dictionary has been re-created (used to crate Jira issues)
    # Use file name embedded remark ID to find orignal main issue key (in old ticketing system) and remark summmay text
    # these info needed, when deciding to which Jira remark (subtask) to attach curren attachment file
    
             
    find=3135983
    #find=3128668

    
    
    attachments=glob.glob("{0}/*/*".format(filepath)) # get all attachments
    if (len(attachments) > 0): # if any attachment with key embedded to name found
        
        # RENAME ATTACHMENT FILES USING DIRECTORY ID NUMBER
        # FILE Attachment 1016:..\..\MIKAN_TYO\ASIAKKAAT\Meyer\tsp\04_Attachment Remarks\394_3429854\IMG_0330.JPG RENAMING -->
        # Attachment 1016:..\..\MIKAN_TYO\ASIAKKAAT\Meyer\tsp\04_Attachment Remarks\394_3429854\3429854_IMG_0330.JPG

            i=1
            for item in attachments: # check them all
                #jira.add_attachment(issue=IssueID, attachment=attachments[0])
                print "*****************************************"
                print "Attachment {0}:{1}".format(i,item)
                regex = r"(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                regex2=r"(.*?)(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                #test_str = "\\394_3428553\\"
                match = re.search(regex, item)
                match2 = re.search(regex2, item)
                hit=match.group(4)
                origname=match.group(6)
                path=match2.group(1)+match2.group(2)+match2.group(3)+match2.group(4)+match2.group(5)
                print "Attachment remark ID:{0}".format(hit)
                
                
                find=int(hit)
                 # uses "find" to define which remark original ID is being searched
                for key, value in Issues.iteritems() :
                    #print key, value
                    #print "************************************"
                    for key2, value2 in value.iteritems():
                        #print key2, value2
                        if key2=="REMARKS":
                            #print key2,value2
                            #print "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
                            for key3, value3 in value2.iteritems():
                                #print key3,value3
                    
                                if key3==find:
                                    print "*********** HIHIT ******"
                                    #print key3, value3
                                    #print key,value
                                    print "ORIGINAL KEY:{0}  ORIGINAL REMARK KEY:{1}".format(key,key3)
                                    print "SUMMARY:{0}".format(value3["SUMMARY"].encode('utf-8'))
                
                
                
                i=i+1
    else:
        print "--> No attachments??"
    
    
    
  
        
    end = time.clock()
    totaltime=end-start
    print "Time taken:{0} seconds".format(totaltime)
       
            
    print "*************************************************************************"
        

 
  

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 
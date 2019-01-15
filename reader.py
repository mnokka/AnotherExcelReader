#encoding=utf8

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
import unidecode
import array as arr

start = time.clock()
__version__ = u"0.3.1394" 

# should pass via parameters
#ENV="demo"
ENV=u"PROD"

logging.basicConfig(level=logging.DEBUG) # IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out



def main(argv):
    
    JIRASERVICE=u""
    JIRAPROJECT=u""
    PSWD=u''
    USER=u''
  
    logging.debug (u"--Python starting checking Jira issues for attachemnt adding --") 

 
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
    parser.add_argument('-x','--ascii', help='<ascii file names>') #adhoc operation activation
        
    args = parser.parse_args()
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
           
    filepath = args.filepath or ''
    excelfilepath = args.excelfilepath or ''
    filename = args.filename or ''
    subfilename=args.subfilename or 'NOTDEFINED' #not needed
    
    JIRASERVICE = args.service or ''
    JIRAPROJECT = args.project or ''
    PSWD= args.password or ''
    USER= args.user or ''
    RENAME= args.rename or ''
    ASCII=args.ascii or ''
    
    # quick old-school way to check needed parameters
    if (filepath=='' or  JIRASERVICE=='' or  JIRAPROJECT==''  or PSWD=='' or USER==''  or excelfilepath=='' or filename==''):
        parser.print_help()
        print "args: {0}".format(args)
        sys.exit(2)
        
    #adhoc ascii conversion to file names (��� and german letters off)
    if (ASCII):
       DoAscii(filepath,JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename,ENV)
       exit()
    
    
    Parse(filepath,JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename,ENV)


######################################################

def DoAscii(filepath,JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename,ENV):


 #Deactivated renaming command     
    attachments=glob.glob("{0}/*/*".format(filepath))
    if (len(attachments) > 0): # if any attachment with key embedded to name found
        
        # convert file names to ascii
        if (1):
            i=1
            for item in attachments: # add them all
                #jira.add_attachment(issue=IssueID, attachment=attachments[0])
                print "*****************************************"
                print "Attachment {0}:{1}".format(i,item)
                regex = r"(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                regex2=r"(.*?)(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                
                match = re.search(regex, item)
                match2 = re.search(regex2, item)
             
                path=match2.group(1)+match2.group(2)+match2.group(3)+match2.group(4)+match2.group(5)
                origname=match2.group(7)
                print "Original name:{0}".format(origname)
                #newname=unidecode.unidecode(u'{0}').format(origname)
                
                #newname=unidecode.unidecode(origname)
                #newname=origname.encode('utf-8')
                command="unidecode -c \"{0}\"".format(origname) # did not get working directly
                print "Command: {0}".format(command)
                newname=os.popen(command).read()
                
                
                print "GOING TO DO UNIDECODING:{0} ---->  {1}".format(origname,newname)
                print "New name:{0}".format(newname)
                print "Path: {0}".format(path)
                newfile=path+"\\"+newname
                print "Newfile: {0}".format(newfile)
                if (item==newfile):
                    print "No need for ascii renaming"
                else:
                    new_item= item
                    x=new_item.replace("\\","\\\\")
                    y=newfile.replace("\\","\\\\")
                    y=y.replace("\n","") #remove linefeed created somewhere earlier
                    print "GOING TO DO RENAMING:{0} ---->{1}".format(x,y)
                    os.rename(x,y)
                    print "Done!!!"
                    print "-------------------------------------------------------------------"
                i=i+1
        else:
            print "--> Renaming bypassed"








############################################################################################################################################
# Parse attachment files and add to matching Jira issue
#

#NOTE: Uses hardcoded sheet/column value

def Parse(filepath, JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename,ENV):
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
    #subfiles=excelfilepath+"/"+subfilename
    #logging.debug ("SubFiles:{0}".format(subfiles))
   
    
    #SubMainSheet="general_report" 
    #subwb= openpyxl.load_workbook(subfiles) 
    #types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    #sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    #SubCurrentSheet=subwb[SubMainSheet] 
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
    #attachments=glob.glob("{0}/*/*".format(filepath))


            
    attachments=[]
    rootDir2 =os.path.realpath(filepath) # from args, assuming .
    for dirName, subdirList, fileList in os.walk(rootDir2):
        print "***********************************************"
        print('Found directory: %s' % dirName)
       # print("Listing files:")
   
        for name in fileList:
            fullpathed = os.path.join(dirName, name)
            #print(fullpathed)
            attachments.append(fullpathed)
        #for name in dirs:
        #    print(os.path.join(dirName, name)) 
            
            

    if (len(attachments) > 0): # if any attachment with key embedded to name found
        
        # RENAME ATTACHMENT FILES USING DIRECTORY ID NUMBER
        # FILE Attachment 1016:..\..\MIKAN_TYO\ASIAKKAAT\Meyer\tsp\04_Attachment Remarks\394_3429854\IMG_0330.JPG RENAMING -->
        # Attachment 1016:..\..\MIKAN_TYO\ASIAKKAAT\Meyer\tsp\04_Attachment Remarks\394_3429854\3429854_IMG_0330.JPG
        if (RENAME):
            #NOT IMPLEMENTED FOR PHASE2
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
            if os.path.isdir(item):
                 print "IS DIRECTORY"
            #else:
            #    print "NOT DIRECTORY"
            i=i+1

    
        #Find remark's original parent ID using 1) remark ID in the file name 
        #print "--> SUBTASK EXCEL: {0}".format(subfilename)
        
  
        
   
    if (len(attachments) > 0): # if any attachment with key embedded to name found
        


            i=1
            go=0
            INVENTORY={} #set dictionary
            for item in attachments: # check them all
                #jira.add_attachment(issue=IssueID, attachment=attachments[0])
                print "\n\n****PROCESSING ITEM *************************************"
                print "Attachment {0}:{1}".format(i,item)
               
                regex = r"(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                #regex = r"(\\)(\d\d\d)(_)(\d+)(\\)(.*)"
                regex2=r"(.*?)(\\)(INSP_\d\d\d)(_No)(\d+)(_)(.*)"
                #test_str = "\\394_3428553\\"
                match = re.search(regex, item)
                match2 = re.search(regex2, item)
                
                if (match):
                    hit=match.group(4)
                    print "Attachment old issue number ID:{0}".format(hit)
                    go=1
                
                elif (match2):
                    hit=match2.group(5)
                    print "Attachment old issue number ID:{0}".format(hit)
                    go=1         
                else:
                    print "no match"
                    go=0
               
                #find=int(hit)
                 # uses "find" to define which remark original ID is being searched
                 
                 
                summary_text="kissa" #not needed here
                
                if (go==1):
                    key3=hit 
                    #Set custome filed hard way, one really should use the names
                    if (ENV=="demo"):
                        key_field="cf[12317]"
                    if (ENV=="PROD"):
                        key_field="cf[12900]"
                
                        jql_query="project = {0} and {1} ~ {2}".format(JIRAPROJECT,key_field,key3)
                        print "Query:{0}".format(jql_query)
                        #project = NB1394FERU and cf[12900] ~ 470
                        ask_it=jira.search_issues(jql_query)
                        print "Query:{0}".format(jql_query)
                        print "Feedback:{0}".format(ask_it) 
                        
                        
                                 
                    for issue in ask_it:
                        print "-----> GOING TO ADD ATTACHMENT:{0}\n TO JIRA ISSUE:{1}\n  (Summary:{2})".format(item,issue.key,issue.fields.summary.encode('utf-8'))      
                        
                        #use dictionary to keep record of how many attachment for one issue
                        if (issue.key in INVENTORY):
                            value=INVENTORY.get(issue.key,"10000") # 1000 is default value
                            value=value+1 
                            INVENTORY[issue.key]=value
                        else:
                            INVENTORY[issue.key]=1 # first issue attachment, create entry for dictionary
                        
                        # this makes the chamge!
                        #jira.add_attachment(issue=issue.key, attachment=item)
                        #print "Attachment:{0} added".format(item) 
                 
                i=i+1
                go=0

    
    for key,value in INVENTORY.items():
        print "ISSUE:{0}  => ATTACHMENTS ADDITIONS: {1}".format(key,value)  
        
    print "FORCE ENDING 1"
    sys.exit(5)  
      

      
      
    
    
    
    
  
        
    end = time.clock()
    totaltime=end-start
    print "Time taken:{0} seconds".format(totaltime)
       
            
    print "*************************************************************************"
        

 
  

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 
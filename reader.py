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
    #parser.add_argument('-n','--filename', help='<Main tasks Excel filename>')
    #parser.add_argument('-m','--subfilename', help='<Subtasks Excel filename>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
    
    parser.add_argument('-w','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    parser.add_argument('-s','--service', help='<JIRA service>')
    parser.add_argument('-p','--project', help='<JIRA project>')
    parser.add_argument('-n','--rename', help='<rename files>') #adhoc operation activation
   
        
    args = parser.parse_args()
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
           
    filepath = args.filepath or ''
    #filename = args.filename or ''
    #subfilename=args.subfilename or ''
    
    JIRASERVICE = args.service or ''
    JIRAPROJECT = args.project or ''
    PSWD= args.password or ''
    USER= args.user or ''
    RENAME= args.rename or ''
    
    # quick old-school way to check needed parameters
    if (filepath=='' or  JIRASERVICE=='' or  JIRAPROJECT==''  or PSWD=='' or USER=='' ):
        parser.print_help()
        sys.exit(2)
        


    Parse(filepath,JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME)


############################################################################################################################################
# Parse attachment files and add to matching Jira issue
#

#NOTE: Uses hardcoded sheet/column value

def Parse(filepath, JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME):
    logging.debug ("Filepath: %s   " %(filepath))
    #files=filepath+"/"+filename
    #logging.debug ("File:{0}".format(files))
   
    Issues=defaultdict(dict) 

    
    
    
    
   

    
   
    print Issues.items() 
    
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
                os.rename(item, newfile)
                print "Done!!!"
                i=i+1
        else:
            print "--> Renaming bypassed"
        
        
        
       
            
    print "*************************************************************************"
        

 
  

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 
# Create Issue to given JIRA
# Requires .netrc file for authentication
#
# 6.12.2016 mika.nokka1@gmail.com for Ambientia
# 
# NOTE: For POC removed .netrc authetication, using pure arguments
# NOTE: NOT TESTED with changes added using normal commandline usage!!!!
# (used via importing only)
# 

import datetime 
import time
import argparse
import sys
import netrc
import requests, os
from requests.auth import HTTPBasicAuth
# We don't want InsecureRequest warnings:
import requests
requests.packages.urllib3.disable_warnings()
import itertools, re, sys
from jira import JIRA
import random


__version__ = "0.1"
thisFile = __file__

    
def main(argv):

    JIRASERVICE=""
    JIRAPROJECT=""
    JIRASUMMARY=""
    JIRADESCRIPTION=""
    PSWD=''
    USER=''
    
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com
    
    .netrc file used for authentication. Remember chmod 600 protection
    Creates issue for given JIRA service and project in JIRA
    Used to crate issue when build fails in Bamboo
    
    EXAMPLE: python {1}  -j http://jira.test.com -p BUILD -s "summary text"


    """.format(__version__,sys.argv[0]))

    parser.add_argument('-p','--project', help='<JIRA project key>')
    parser.add_argument('-j','--jira', help='<Target JIRA address>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
    parser.add_argument('-s','--summary', help='<JIRA issue summary>')
    parser.add_argument('-d','--description', help='<JIRA issue description>')
    
    parser.add_argument('-ps','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    
    args = parser.parse_args()
        
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
         

    JIRASERVICE = args.jira or ''
    JIRAPROJECT = args.project or ''
    JIRASUMMARY = args.summary or ''
    JIRADESCRIPTION = args.description or ''
  
    PSWD= args.description or ''
    USER= args.description or ''
  
    # quick old-school way to check needed parameters
    if (JIRASERVICE=='' or  JIRAPROJECT=='' or JIRASUMMARY=='' or PSWD=='' or USER==''):
        parser.print_help()
        sys.exit(2)

    user, PASSWORD = Authenticate(JIRASERVICE,PSWD,USER)
    jira= DoJIRAStuff(user,PASSWORD,JIRASERVICE)
    #CreateIssue(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION,PSWD)
    
####################################################################################################   
# POC skips .netrc usage
# 
def Authenticate(JIRASERVICE,PSWD,USER):
    host=JIRASERVICE
    #credentials = netrc.netrc()
    #auth = credentials.authenticators(host)
    #if auth:
    #    user = auth[0]
    #    PASSWORD = auth[2]
    #    print "Got .netrc OK"
    #else:
    #    print "ERROR: .netrc file problem (Server:{0} . EXITING!".format(host)
    #    sys.exit(1)
    user=USER
    PASSWORD=PSWD

    f = requests.get(host,auth=(user, PASSWORD))
         
    # CHECK WRONG AUTHENTICATION    
    header=str(f.headers)
    HeaderCheck = re.search( r"(.*?)(AUTHENTICATION_DENIED|AUTHENTICATION_FAILED)", header)
    if HeaderCheck:
        CurrentGroups=HeaderCheck.groups()    
        print ("Group 1: %s" % CurrentGroups[0]) 
        print ("Group 2: %s" % CurrentGroups[1]) 
        print ("Header: %s" % header)         
        print "Authentication FAILED - HEADER: {0}".format(header) 
        print "--> ERROR: Apparantly user authentication gone wrong. EXITING!"
        sys.exit(1)
    else:
        print "Authentication OK \nHEADER: {0}".format(header)    
    print "---------------------------------------------------------"
    return user,PASSWORD

###################################################################################    
def DoJIRAStuff(user,PASSWORD,JIRASERVICE):
 jira_server=JIRASERVICE
 try:
     print("Connecting to JIRA: %s" % jira_server)
     jira_options = {'server': jira_server}
     jira = JIRA(options=jira_options,basic_auth=(user,PASSWORD))
     print "JIRA Authorization OK"
 except Exception,e:
    print("Failed to connect to JIRA: %s" % e)
 return jira   
    
####################################################################################
def CreateIssue(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION,KEY,CREATOR,CREATED,INSPECTED,SHIP,PERFOMER,RESPONSIBLE,BLOCK,DEPARTMENT,DECK,ISSUETYPE):
    jiraobj=jira
    project=JIRAPROJECT
    
    lottery = random.randint(1,3)
    
    if (lottery==1):
        TASKTYPE="Steal"
    elif (lottery>1):
        TASKTYPE="Outfitting"
    else:
        TASKTYPE="Task"
    
    print "Creating issue for JIRA project: {0}".format(project)
    

    
    issue_dict = {
    'project': {'key': JIRAPROJECT},
    'summary': JIRASUMMARY,
    'description': JIRADESCRIPTION,
    'issuetype': {'name': TASKTYPE},
    'customfield_12317': str(KEY),  # Key in ALM demo
    'customfield_12318': str(CREATOR),  # Reporter in ALM demo
    #'customfield_12319': str(REPORTER),  # Creator in ALM demo
    #'customfield_12320': str(CREATED),  # Original Created Time in ALM demo
    'customfield_12321': str(SHIP), # Ship Number in ALM demo
    'customfield_12322': str(PERFOMER), # PerformerNW in ALM demo
    'customfield_12323': str(RESPONSIBLE), # ResponsibleNW in ALM demo
    'customfield_12324': str(BLOCK), # BlockNW in ALM demo
    'customfield_12326': DECK.encode('utf-8'), # DeckNW in ALM demo
    'customfield_12328': str(DEPARTMENT), # DEPARTMENTNW in ALM demo
    'customfield_12330': str(INSPECTED), # Original inspectiond date
    'customfield_12331': ISSUETYPE.encode('utf-8'), # Original inspectiond date
    }

    try:
        new_issue = jiraobj.create_issue(fields=issue_dict)
        print "Issue created OK"
    except Exception,e:
        print("Failed to create JIRA object, error: %s" % e)
        sys.exit(1)
    return new_issue 

############################################################################################'
# Quick way to create subtask
#
def CreateSubTask(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION,PARENT):
    jiraobj=jira
    project=JIRAPROJECT
    print "Creating subtask for JIRA project: {0} Parent:{1}".format(project,PARENT)
    issue_dict = {
    'project': {'key': JIRAPROJECT},
    'summary': JIRASUMMARY,
    'description': JIRADESCRIPTION,
    'issuetype': {'name': 'Remark1'}, #  is a Sub-task type CHANGE FOR target system
    'parent' : { 'id' : str(PARENT)},   # PARENT is an object, convert
    }


    try:
        new_issue = jiraobj.create_issue(fields=issue_dict)
        print "Subtask created OK"
    except Exception,e:
        print("Failed to create JIRA object, error: %s" % e)
        sys.exit(1)
    return new_issue 

        
if __name__ == "__main__":
        main(sys.argv[1:])
        
        
        
        
        
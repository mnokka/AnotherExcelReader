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

from author import Authenticate  # no need to use as external command
from author import DoJIRAStuff

__version__ = "0.1"
thisFile = __file__

    
def main(argv):

    JIRASERVICE=""
    JIRAPROJECT=""
    JIRASUMMARY=""
    JIRADESCRIPTION=""
    PSWD=''
    USER=''
    jira=''
    
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
    
    parser.add_argument('-x','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    
    args = parser.parse_args()
        
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
         

    JIRASERVICE = args.jira or ''
    JIRAPROJECT = args.project or ''
    JIRASUMMARY = args.summary or ''
    JIRADESCRIPTION = args.description or ''
  
    PSWD= args.password or ''
    USER= args.user or ''
  
    # quick old-school way to check needed parameters
    if (JIRASERVICE=='' or  JIRAPROJECT=='' or JIRASUMMARY=='' or PSWD=='' or USER==''):
        parser.print_help()
        sys.exit(2)

    user, PASSWORD = Authenticate(JIRASERVICE,PSWD,USER)
    jira= DoJIRAStuff(user,PASSWORD,JIRASERVICE)
    #CreateIssue(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION,PSWD)
    CreateSimpleIssue(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION)
    

####################################################################################
def CreateIssue(ENV,jira,JIRAPROJECT,JIRASUMMARY,KEY,ISSUETYPE,ISSUETYPENW,STATUS,STATUSNW,PRIORITY,RESPONSIBLENW,RESPONSIBLE,INSPECTEDTIME,SHIP,SHIPNW,SYSTEM,PERFORMERNW,DEPARTMENTNW,DEPARTMENT,DESCRIPTION,AREA,SURVEYOR,DECKNW,BLOCKNW,FIREZONENW):
    jiraobj=jira
    project=JIRAPROJECT

    
    print "Creating issue for JIRA project: {0}".format(project)
    

    issue_dict = {
    'project': {'key': JIRAPROJECT},
    'summary': JIRASUMMARY,
    'description': DESCRIPTION,
    'issuetype': {'name': ISSUETYPE},
    
    'customfield_14613' if (ENV =="DEV") else 'customfield_14212' : str(SYSTEM),
    'customfield_14612' if (ENV =="DEV") else 'customfield_14212' : str(SHIP),
    'customfield_14607' if (ENV =="DEV") else 'customfield_14212' : str(PERFORMERNW),
    
    'customfield_10013' if (ENV =="DEV") else 'customfield_14212' : str(INSPECTEDTIME),
    'customfield_12900' if (ENV =="DEV") else 'customfield_14212' : str(KEY),
    }

    #status

    try:
        new_issue = jiraobj.create_issue(fields=issue_dict)
        print "Issue created OK"
        print "Updating now all selection custom fields"

        
        # all custom fields could be objects with certain values for certain environments
        if (ENV =="DEV"):
            DEPARTMENTNWTFIELD="customfield_14608" 
            new_issue.update(fields={DEPARTMENTNWTFIELD: {'value': DEPARTMENTNW}})  
            
            DEPARTMENTFIELD="customfield_10010" 
            new_issue.update(fields={DEPARTMENTFIELD: {'value': DEPARTMENT}})
            
            STATUSNWFIELD="customfield_14606" 
            new_issue.update(fields={STATUSNWFIELD: {'value': STATUSNW}})  
            
            
            ISSUTYPENWFIELD="customfield_14604" 
            new_issue.update(fields={ISSUTYPENWFIELD: {'value': ISSUETYPENW}})  
            
            #SYSTEMNUMBERNWFIELD="customfield_14605" 
            #if (SYSTEM is None):
            #    new_issue.update(fields={SYSTEMNUMBERNWFIELD: {"id": "-1"}})
            #else:    
            #    new_issue.update(fields={SYSTEMNUMBERNWFIELD: {'value': SYSTEM}})
            
            CustomFieldSetter(new_issue,"customfield_14604" ,ISSUETYPENW)
            
            
            
        elif (ENV =="PROD"):
            DEPARTMENTNWTFIELD="customfield_14328" 
            new_issue.update(fields={DEPARTMENTNWTFIELD: {'value' : DEPARTMENTNW}})  
            
            DEPARTMENTFIELD="customfield_14328" 
            new_issue.update(fields={DEPARTMENTFIELD: {'value' : DEPARTMENT}}) 
            
            STATUSNWFIELD="customfield_14328" 
            new_issue.update(fields={STATUSNWFIELD: {'value' : STATUSNW}}) 
            
            ISSUTYPENWFIELD="customfield_14328" 
            new_issue.update(fields={ISSUTYPENWFIELD: {'value' : ISSUETYPENW}})
            
            SYSTEMNUMBERNWFIELD="customfield_14328" 
            new_issue.update(fields={SYSTEMNUMBERNWFIELD: {'value' : SYSTEM}})
            
    
       
    
    
        print "Transit issue status"
        
        
        
        if (STATUS != "Todo"): # initial status after creation
            
            #map state to neede transit. Assunming WF supports thse transit (do for example admin only transit possibilty for migration)
            if (STATUS=="Closed"):
                TRANSIT="CLOSED"
            if (NEWSTATUS=="Inspected"):
                TRANSIT="INSPECTED"
           
            
            print "Newstatus will be:{0}".format(STATUS)
            print "===> Executing transit:{0}".format(TRANSIT)
            jiraobj.transition_issue(new_issue, transition=TRANSIT)  # trantsit to state where it was in excel
        else:
            print "Initial status found: {0}, nothing done".format(STATUS)
    
    
    
    except Exception,e:
        print("Failed to create JIRA object, error: %s" % e)
        sys.exit(1)
    return new_issue 

##################################################################################
# used only selection custom fields

def CustomFieldSetter(new_issue,CUSTOMFIELDNAME,CUSTOMFIELDVALUE):
    
    try:
    
        if (CUSTOMFIELDVALUE is None):
            new_issue.update(fields={CUSTOMFIELDNAME: {"id": "-1"}})
        else:    
            new_issue.update(fields={CUSTOMFIELDNAME: {'value': CUSTOMFIELDVALUE}})
        print "Issue updated ok"    

    except Exception,e:
        print("Failed to UPDATE JIRA object, error: %s" % e)
        sys.exit(1)

############################################################################################'
# Quick way to create subtask
#
def CreateSubTask(jira,JIRAPROJECT,SUBSUMMARY,SUBISSUTYPENW,SUBISSUTYPE,SUBSTATUSNW,SUBSTATUS,SUBREPORTERNW,SUBCREATED,SUBDESCRIPTION,SUBSHIPNUMBER,SUBSYSTEMNUMBERNW,SUBPERFORMER,SUBRESPONSIBLENW,SUBASSIGNEE,SUBINSPECTION,SUBDEPARTMENTNW,SUBDEPARTMENT,SUBBLOCKNW,SUBDECKNW):
    jiraobj=jira
    project=JIRAPROJECT
 
    print "Creating subtask for JIRA project: {0} Parent:{1}".format(project,PARENT)
    issue_dict = {
    'project': {'key': JIRAPROJECT},

    'summary': SUBSUMMARY,
    'description': JIRASUBDESCRIPTION,
    'issuetype': {'name': SUBTASKTYPE}, #  is a Sub-task type CHANGE FOR target system
    'parent' : { 'id' : str(PARENT)},   # PARENT is an object, convert  SUBISSUETYPE
    #ALMDEMO:
    #'customfield_12332': str(SUBTASKID), # SubtaskNW
    #'customfield_12323': SUBRESPONSIBLE.encode('utf-8'), # ResponsibleNW in ALM demo
    #'customfield_12331': SUBISSUETYPE.encode('utf-8'), # Original date
    #'customfield_12322': SUBPERFORMER.encode('utf-8'), # PerformerNW in ALM demo
    #'customfield_12320': SUBCREATED.encode('utf-8'), # Original Created Date in ALM demo
    #PROD:
    'customfield_13100': str(SUBTASKID), # SubtaskNW
    'customfield_12906': SUBRESPONSIBLE.encode('utf-8'), # ResponsibleNW in ALM demo
    'customfield_13101': SUBISSUETYPE.encode('utf-8'), # Original date
    'customfield_12905': SUBPERFORMER.encode('utf-8'), # PerformerNW in ALM demo
    'customfield_12903': SUBCREATED.encode('utf-8'), # Original Created Date in ALM demo
    }


    try:
        new_issue = jiraobj.create_issue(fields=issue_dict)
        print "Subtask created OK"
    except Exception,e:
        print("Failed to create JIRA object, error: %s" % e)
        sys.exit(1)
    return new_issue 

########################################################################################
# test creating issue with multiple selection list custom field
def CreateSimpleIssue(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION):
    #jiraobj=jira
    project=JIRAPROJECT
    
    
    #lottery = random.randint(1,3)
    #if (lottery==1):
    #    TASKTYPE="Steal"
    #elif (lottery>1):
    #    TASKTYPE="Outfitting"
    #else:
    #    TASKTYPE="Task"
    
    #TASKTYPE="Hull Inspection NW"
    TASKTYPE="Task"
    
    print "Creating issue for JIRA project: {0}".format(project)
    

    
    issue_dict = {
    'project': {'key': JIRAPROJECT},
    'summary': str(JIRASUMMARY),
    'description': str(JIRADESCRIPTION),
    'issuetype': {'name': TASKTYPE},
    'customfield_14600' : [{'value': str("cat")},{'value': str("bear")}] ,
    }

    try:
        new_issue = jira.create_issue(fields=issue_dict)
        print "Issue created OK"
    except Exception,e:
        print("Failed to create JIRA object, error: %s" % e)
        sys.exit(1)
    return new_issue 



        
if __name__ == "__main__":
        main(sys.argv[1:])
        
        
        
        
        
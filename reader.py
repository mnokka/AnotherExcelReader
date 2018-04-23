#encoding=latin1

# POC tool to read Excel using Python
# Data will be used to create subtasks / add attachments to Jira main issues
#
# Author mika.nokka1@gmail.com for Ambientia
#TODO 
# Use Pandas instead?
#
#from __future__ import unicode_literals

import openpyxl 
import getopt,sys, logging
import argparse
import re
from collections import defaultdict

__version__ = "0.1.1394"


logging.basicConfig(level=logging.DEBUG) # IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out



def main(argv):
    
    filepath=''
    filename=''

  
    logging.debug ("--Python starting Excel reading --") 

 
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com for Ambientia
    
    USAGE:
    -filepath  | -p <Path to Excel file directory>
    -filename   | -n <Excel filename>

    """.format(__version__,sys.argv[0]))

    parser.add_argument('-p','--filepath', help='<Path to Excel file directory>')
    parser.add_argument('-n','--filename', help='<Excel filename>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
        
    args = parser.parse_args()
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
           
    filepath = args.filepath or ''
    filename = args.filename or ''
    
    
    # quick old-school way to check needed parameters
    if (filepath=='' or  filename=='' ):
        parser.print_help()
        sys.exit(2)
        


    Parse(filepath, filename)


############################################################################################################################################
#
def Parse(filepath, filename):
    logging.debug ("Filepath: %s     Filename:%s" %(filepath ,filename))
    files=filepath+"/"+filename
    logging.debug ("File:{0}".format(files))
   
    Issues=defaultdict(dict) 

    
    MainSheet="general_report" # hardcoded for main issues?
    
    wb= openpyxl.load_workbook(files)
    types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    CurrentSheet=wb[MainSheet] 
    #logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    #logging.debug ("First row:{0}".format(CurrentSheet['A4'].value))

    #CONFIGURATIONS
    DATASTARTSROW=5 # data section starting line
    #EXCEL COLUMN MAPPINGS
    K=11 #LINKED_ISSUES 
    G=7 #REPORTER
    C=3 # SUMMARY
    #for cell in CurrentSheet['A']:
    #    logging.debug  ("Row value:{0}".format(cell.value))
  
    
    ##############################################################################################
    #Go through main excel sheet for main issue keys (and contents findings)
    i=DATASTARTSROW # brute force row indexing
    for row in CurrentSheet[('B{}:B{}'.format(DATASTARTSROW,CurrentSheet.max_row))]:  # go trough all column B (KEY) rows
        for mycell in row:
            KEY=mycell.value
            logging.debug("ROW:{0} Original ID:{1}".format(i,mycell.value))
            Issues[KEY]={} # add to dictionary as master key (KEY)
            
            LINKED_ISSUES=(CurrentSheet.cell(row=i, column=K).value) #NOTE THIS APPROACH GOES ALWAYS TO THE FIRST SHEET
            #logging.debug("Attachment:{0}".format((CurrentSheet.cell(row=i, column=K).value))) # for the same row, show also column K (LINKED_ISSUES) values
            Issues[KEY]["LINKED_ISSUES"] = LINKED_ISSUES
            
            REPORTER=(CurrentSheet.cell(row=i, column=G).value)
            Issues[KEY]["REPORTER"] = REPORTER
            
            SUMMARY=(CurrentSheet.cell(row=i, column=C).value)
            Issues[KEY]["SUMMARY"] = SUMMARY
            
            #Create sub dictionary for possible subtasks (to be used later)
            Issues[KEY]["REMARKS"]={}
            
            logging.debug("---------------------------------------------------")
            i=i+1
    #print Issues
    print Issues.items() 
    key=18503 # check if this key exists
    if key in Issues:
        print "EXISTS"
    else:
        print "NOT THERE"
    for key, value in Issues.iteritems() :
        print key, value

    ############################################################################################################################
    # Check any remarks (subtasks) for main issue
    
    RemarksSheet="Tabelle2" # hardcoded for main issues?
    SubSheet1=wb[RemarksSheet]

    
    # Find KTR keyword, after which subtasks are defined
    i=1
    for row in SubSheet1[('A{}:A{}'.format(SubSheet1.min_row,SubSheet1.max_row))]:  # go trough all column B (KEY) rows    
        for mycell in row:
            TMP=mycell.value
            #logging.debug("ROW:{0} Value:{1}".format(i,mycell.value))
            if TMP=="KTR":
                DATASTARTSROW=i+1 # this line includes first subtask definition
                break # TODO fix logic
            i=i+1
    
            
    i=DATASTARTSROW # brute force row indexing
    for row in SubSheet1[('B{}:B{}'.format(DATASTARTSROW,SubSheet1.max_row))]:  # go trough all column B (KEY) rows
        for mycell in row:
            KEY=mycell.value
            logging.debug("ROW:{0} Original ID:{1}".format(i,KEY))
            if KEY in Issues: 
              
                print "Subtask has a known parent."
                #BGR=(SubSheet1.cell(row=i, column=J).value) # This approach takes always values from the first sheet of excel 
                REMARKKEY=SubSheet1['J{0}'.format(i)].value  # column J holds BGR numbers
                #Issues[KEY]["REMARKS"]={}
                Issues[KEY]["REMARKS"][REMARKKEY] = {}
                
                DECK=SubSheet1['S{0}'.format(i)].value  # column S holds BGR numbers
                Issues[KEY]["REMARKS"][REMARKKEY]["DECK"] = DECK
                logging.debug("i:{0} DECK:{1} REMARKKEY:{2}".format(i,DECK,REMARKKEY))
            else:
                print "Error: Unknown parent found"
        print "----------------------------------"
        i=i+1

    for key, value in Issues.iteritems() :
        print key, value

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
   main(sys.argv[1:]) 
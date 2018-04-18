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
        
   # logging.debug ("filepath: %s   filename:%s" %(filepath ,filename))

    Parse(filepath, filename)



def Parse(filepath, filename):
    logging.debug ("Filepath: %s     Filename:%s" %(filepath ,filename))
    files=filepath+"/"+filename
    logging.debug ("File:{0}".format(files))
   

    
    wb= openpyxl.load_workbook(files)
    types=type(wb)
    logging.debug ("Type:{0}".format(types))
    logging.debug ("--Python exiting--")

if __name__ == "__main__":
   main(sys.argv[1:]) 
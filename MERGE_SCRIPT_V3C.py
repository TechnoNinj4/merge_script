# Microsoft Word Mass Merge Script:

# Library And Module Call List:

from __future__ import print_function
from mailmerge import MailMerge
import datetime
import time
import pandas
import numpy as np
import os
import socket
import json

## ENVIRONMENT PROVIDED VALUES ##

# Current Environment Variables:
# This first set of actions capture information about the current environment for use at
# later points in the script, these in memory variables can be called at anytime from this
# point onward.

# starttime - time stamp used for filename generation
starttime = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
print(starttime)

# simpletime - time string used for notification processes
simpletime = datetime.datetime.now().strftime("%H:%M:%S")

# pcname - the fully qualified domain name of the computer running the script
pcname = socket.getfqdn()

# usr - the username of person running the script
usr = os.getlogin()

## SCRIPT PULLED VALUES ##

# Script Configuration Files:
# The following set of commands read lines from a set of JSON files and interpret 
# them as variables to be used at later points within the script.

# baseconfig - the file object in the line below is the primary configuration file
# for all script actions, this is the only file where a recompile or line edit will
# be required if it's location and/or name changes
baseconfig = json.load(open("###########.json","r"))

# Storage Locations And Directory Variables:
# The following set of variables are director locations the script will read and write 
# data to. Note that some paths use "/" instead of "\", the "\" symbol is read by python
# as and escape character or a character that triggers the next character(s) to be
# interpeted differently, in order to use the "\" you need to escape it out or write
# them in sets of 2 for every 1 instance.

# sourcedir - the location of the source files for which new documents will obtain data
# from written as a network path
sourcedir = baseconfig["sourcedir"].replace("+usr+",usr)

# sourcefiles - the file names of the data files for to be read by this script
sourcefiles = baseconfig["sourcefiles"]

# delimiterline - the single line string that denotes the end of a given set of data for
# a single document generation task
delimiterline = baseconfig["delimiterline"]

# templatepath - this directory is the location for which the Microsoft Word templates will
# be stored
templatepath = baseconfig["templatepath"]

# nontemplates - the line below checks the JSON configuration file for an array containing a
# list of currently used document codes that don't have any real templates associated with them,
# these codes may have other tasks associated them and thus are documented as safe to skip
# exceptions
nontemplates = baseconfig["nontemplates"]

# outputdir1 - in order for this script to produce actual documents it requires a location to
# write the new files to, the lines below checks if the user already has a folder in the their
# home directory for merged documents, if it does not it will create that folder labled 
# "MERGE_REVIEW", afterwards it will associate that location as the named varible "outputdir1"
if not os.path.exists(baseconfig["outputdir1"].replace("+usr+",usr)):
    os.makedirs(baseconfig["outputdir1"].replace("+usr+",usr))

outputdir1 = baseconfig["outputdir1"].replace("+usr+",usr)

# File Types Produced:
# A few of the files generated from this script will have similar prefix and name, the
# thing the sets them appart is their suffix or file type.

# outputtype1 - file type of all documents file objects produced
outputtype1 = baseconfig["outputtype1"]

## MAIN LOOP ##

# Convert Text List to Clean DataFrame:
# This is the part of the script that actually reads the raw data that using
# the pandas library the data is read and stored in memory as dataframe, or a pandas
# in-memory table, this data can then be searched indexed and called using different pandas
# commands as well as imported into other objects - which is what the next section does.
# Before anything can be done with the raw data, the text needs to be cleaned up of any
# unnecessary and/or incompatible hidden characters, hense the data is read through a series
# of regular expression filters clearing our new line characters, tabs, and a few others.

for index in range(len(sourcefiles)):
    merge_data = ((pandas.DataFrame(((((open(sourcedir+sourcefiles[index], "r")).read()).replace('\x12', '')).replace('\x05', '')).split(delimiterline)))[0].apply(lambda x: pandas.Series(x.split('\n')))).replace(np.nan, '', regex=True)[:-1]
    print(merge_data)
    colcount = str(object = len(merge_data.columns) )
#
# Document Generation Loop:
# The follow section generates the documents from the data in a series of steps:
# - The newly re-organized data in the pandas dataframe is read per column per row
#   at a time. The first column of each row is the actual document template code called,
#   if said code matches any of the codes listed the "nontemplates" variable the
#   job is skipped and the next row begins.
# - Assuming the code is not listed in the "nontemplates", the rest of the data
#   will be added to a dictionary, "mrg_dt" where in each column's data in the current row is
#   associated with a number, these numbers are the field codes used in the templates
#   to indicate a specific data point, they are 1 to infinity so each line number is a different
#   mergeable data field
#
    for index, row in merge_data.iterrows():
        if str(object = row[0]) in nontemplates:
            print("Document code "+row[0]+" is a non-template code and will automatically be skipped.")
        else:
            print("Document code "+row[0]+" was called for file number "+row[7]+".")
            if os.path.exists(templatepath+row[0]+'.docx'):
                print("A document template was detected for the document code called, proceeding to generate merged file,...")
                field_1 = str(object = row[0])
                print(row[0])
                templatefile = templatepath+row[0]+'.docx'
                mrg_dt = {}
                count = 0
                for i in row:
                    count += 1
                    mrg_dt.update({str(object = count):str(object = i)})
#
# File Output:
# Using the previously saved variables each file is contructed with a unique file name using
# a combination of the time of generation and the username of the user running the job, proceed
# a set of fields defined in the configuration JSON file.
#
                rt = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
                filename = outputdir1+row[7]+'_'+row[0]+'_'+rt+'_'+usr+outputtype1
                document = MailMerge(templatefile)
                document.merge_pages([mrg_dt])
                document.write(filename)
                print(filename)
                if os.path.exists(filename):
                    print("Document merge process sucessful, check the MERGE_REVIEW folder for the new file.")
#
# Merge Error Handling:
# In line with the opertunistic nature and narrow function set this script errors fall into 2 categories,
# errors with the merge process or just not having a template. Errors with the merge process can occur,
# if for one or more reasons, the process or actual command line running the script is interrupted or for
# whatever in the rare occurance that their is incompatible formatting present in a template. Errors where
# there is no template present are exactly that, the document code called does not have an associated template
# in the template storage location defined above and is also not included in the file listing exceptions
# that are "nontemplate" codes.
#
                else:
                    print("Document merge process unsucessful, please check the script or merge process for errors.")
            else:
                print("Document code "+row[0]+" does not have an associated document template, the merge task for file number "+row[7]+" with said code was skipped.")
               
# Script End Time:
# This last portion is included purely for testing and benchmarking purposes, in theory you can
# measure the performance of a specific job on compare it to different compute environments.
stoptime = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
print(stoptime)
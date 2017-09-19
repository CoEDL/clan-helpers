
import datetime, os
import textwrap
import pyexcel as pxl
#from pyexcel.ext import xls, xlsx
from Tkinter import Tk
from tkFileDialog import askopenfilename
#import xlrd



if __name__ == '__main__':
    #Request input file
    Tk().withdraw()
    fileIn = askopenfilename(defaultextension='.xlsx',
                             filetypes=[('Excel file','*.xls'),
                                        ('Excel file','*.xlsx'),
                                        ('CSV file','*.csv'),
                                        ])#('All files','*.*')])
    if fileIn == '':
        print 'Exiting: No file selected'
        raise SystemExit(0)
    workbook = pxl.get_book(file_name=fileIn)

    #Convert to dictionary
    wbdict = workbook.to_dict()
    nHeaders_sessions = 2     #2 Header lines
    nHeaders_participants = 1 #1 Header line
    sheet_sessions = wbdict.get("Sessions")
    sheet_participants = wbdict.get("Participants")

    #Get Headers - unused
    headers_session = [header.lower() for header in sheet_sessions[1]]
    headers_participants = [header.lower() for header in sheet_participants[0]]

    nSessions = len(sheet_sessions)-nHeaders_sessions
    nParticipants = len(sheet_participants)-nHeaders_participants

    print "Sheets: ",workbook.sheet_names()
    print "Number of sessions: ",nSessions
    print "Number of participants: ",nParticipants

    #Create output directory
    outpath = 'Output'
    if not os.path.exists(outpath):
        os.makedirs(outpath)

    #List all participants and corresponding IDs from sheet
    allParticipants = {}
    for i in xrange(nHeaders_participants, nParticipants+nHeaders_participants):
        participant = sheet_participants[i]
        clanID     = participant[headers_participants.index('clan code')].upper() #Key will be uppercase
        name_last  = participant[headers_participants.index('participant surname')].strip()
        name_first = participant[headers_participants.index('participant first name')].strip()
        sex        = participant[headers_participants.index('sex')].strip()
        languages  = participant[headers_participants.index('languages (order of proficiency)')].strip().lower()
        role       = participant[headers_participants.index('usual role in recordings')].strip().lower()

        #name = (name_first+' '+name_last).upper() #Key will be uppercase
        allParticipants[clanID]=[name_first, name_last, sex, languages, role]

    #all_participants = [x.upper() for x in workbook['Participants'].column_at(0)]
    #all_IDs =  [x.upper() for x in workbook['Participants'].column_at(4)]

    #a = raw_input('Enter investigator name and ID (comma separated):\n')

    corpus_name = sheet_sessions[0][1]     #Cell B1

    failnames = []
    flags = '' #String for storing file flags
    #Iterate through entries and generate output files
    for i in xrange(nHeaders_sessions, nHeaders_sessions+nSessions):
        session = sheet_sessions[i]
        filename = (session[headers_session.index('session name')]+'.cha').encode('ascii','strict')
        filename = filename.replace('/','_')
        fileout = open(outpath+'/'+filename, 'w');

        #Extract session data        
        language = session[headers_session.index('language')].lower().strip() #Column G - Language
        media    = session[headers_session.index('session name')].strip()     #Column A - Session name
        location = session[headers_session.index('location info')].strip()    #Column E - Location Info        
        transcriber = session[headers_session.index('transcribed by')].strip()#Column O - Transcribed by
        comment  = session[headers_session.index('activity')].strip()         #Column I - Activity
        dateStr  = session[headers_session.index('date recorded (dd/mm/yyyy)')]#(dd-mm-yyyy)')]#Column D - Date recorded (YYYYMMDD)
        duration = session[headers_session.index('length of audio')]          #Column C - Length of audio
        #date     = datetime.datetime.strptime(session[3],'%Y%m%d') #Column D - Date Recorded (YYYYMMDD)

        #Fix parentheses - replace with square brackets
        comment = comment.replace('(','[').replace(')',']').encode('ascii','ignore')
        
       #en_date = True
        if isinstance(dateStr,datetime.date):
            date = dateStr
            en_date = True
        elif isinstance(dateStr,int):
            date = str(dateStr) #datetime.datetime(*xlrd.xldate_as_tuple(dateStr,0))
            en_date = False
        #elif dateStr in ['','?']:
        #    date = dateStr
            #en_date = False
        elif (len(dateStr) == 0):
            date = ''
            en_date = False
        elif isinstance(dateStr, unicode):
            date = dateStr.upper()
            en_date = False
            flags += "Invalid date (line {}):\tSession <{}>\tDate <{}>".format(i+1, session[0], dateStr) + os.linesep
        else:
            raise Exception("Invalid date (line {}): {}".format(i+1, dateStr))

        print flags

        #continue
        #raise SystemExit(0)
        #en_date = (True if date != '' else False)
        en_transc = (True if transcriber != '' else False)
        en_comment = (True if comment != '' else False)
        en_loc = (True if location != '' else False)
        """
        else:
            print "NU"
            try:
                date     = datetime.datetime.strptime(session[3],'%Y%m%d') #Column D - Date Recorded (YYYYMMDD)
            except ValueError:
                print session[3]
                date     = datetime.datetime.strptime(session[3],'%d/%m/%Y') #Column D - Date Recorded (YYYYMMDD)
                print 'Invalid date format. Must be YYYYMMDD'
                print session[3]
                en_date = False
            except:                
                en_date = False
                print session[3]
        """        
        #Multiple participants
        participantsDict = {}
        #Convert ' and 's to commas then split on commas
        IDStrs = session[headers_session.index('speakers')].replace(' and ',',').split(',') #Column F
        IDs = [ID.strip() for ID in IDStrs] #Remove whitespace
        #Add investigator names
        IDStrsInvs = session[headers_session.index('investigator')].replace(' and ',',').split(',') #Column H
        IDs += [ID.strip() for ID in IDStrsInvs] #Remove whitespace

        #Setup header
        participantsStr = ''
        IDsStrList = []
        #Iterate through each name in entry
        for ID in IDs:
            if ID == '': #Blank entry
                ID = 'N/A'
                name_first = 'XX'
                name_last = 'XX'
            
            failStatus = ''
            #Find name in participants list
            if ID.upper() not in allParticipants.keys(): #No entry exists
                sex = ''
                languages = ''
                role = ''
                #Add to list of invalid names
                if ID not in [x[1] for x in failnames]:
                    failnames.append((session[0],ID))
                    failStatus = '\t FAILED'
                    #print '\t',filename, name, ' - FAILED'
            else:                                           #Valid entry exists
                #allParticipants[clanID]=[name_first, name_last, sex, languages, role]
                entry = allParticipants.get(ID)
                name_first = entry[0]
                name_last  = entry[1]
                sex        = entry[2]
                languages  = entry[3]
                role       = entry[4].title()  #Capitalise first letter
            #Add to dictionary
            participantsDict[ID] = dict(first=name_first, last=name_last)            
            participantsDict[ID]['sex'] = sex
            participantsDict[ID]['languages'] = languages
            participantsDict[ID]['role'] = role

            #Add to headers
            IDsStrList.append('@ID:\t'+languages+'|'+corpus_name.lower()+'|'+ID+'||'+sex+'|||'+role+'|||')

            print ID, filename, failStatus

            participantsStr += str(ID+' '+name_first+name_last+' '+role+', ')

        #Output strings
        linesOut = ['@Begin',
                    '@Languages:\t'+language,
                    '@Participants:\t'+participantsStr.strip(', ')
                    ] + IDsStrList + [                    
                    '@Media:\t'+media+', audio',
                    ('@Location:\t'+location if en_loc else ''),
                    ('@Transcriber:\t'+transcriber if en_transc else ''),
                    ('@Comment:\t'+comment if en_comment else ''),
                    ('@Date:\t'+date.strftime('%d-%b-%Y').upper() if en_date else '')
                    ]
        #Account for line wraps and write to file
        for lineOut in  linesOut:
            tempStr = textwrap.fill(lineOut,70, replace_whitespace = False, subsequent_indent = '\t')
            tempStr+='\n'
            tempStr = tempStr.encode('ascii','ignore')
            fileout.write(tempStr)
        fileout.close()
        
        #print filename, language, name_first, duration, date
    if len(failnames) > 0:
        print '\nInvalid speaker names in entries:'
        for x,y in failnames:
            print str(x)+' - '+str(y)

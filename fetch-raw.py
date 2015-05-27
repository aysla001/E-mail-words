# current visualization https://public.tableau.com/profile/aysla3180#!/
import imaplib, email, getpass
from email.utils import getaddresses

# Email settings
# for gmail turn on less secure apps https://www.google.com/settings/security/lesssecureapps
#imap_server: imap.gmail.com 	imap.mail.yahoo.com
#imap_user: wingnnit@gmail.com	aysla001
imap_server = 'imap.mail.yahoo.com'
imap_user = 'aysla001'
imap_password = getpass.getpass()
folder = "Inbox"
	#folder in e-mail that will be searched.

# Connection
conn = imaplib.IMAP4_SSL(imap_server)
(retcode, capabilities) = conn.login(imap_user, imap_password)

# Specify email folder
# print conn.list()
# conn.select("INBOX.Sent Items")
conn.select(folder, readonly=True)   # Set readOnly to True so that emails aren't marked as read

# Search for email ids between dates specified
# NOTE IF YOU SEARCH FOR TOO MANY E-MAILS IT CAN TIME OUT
#result, data = conn.uid('search', None, 'ALL')
result, data = conn.uid('search', None, '(SINCE "01-Apr-2015" BEFORE "01-June-2015")')

# result, data = conn.uid('search', None, '(BEFORE "01-Jan-2014")')
# result, data = conn.uid('search', None, '(TO "user@example.org" SINCE "01-Jan-2014")')

uids = data[0].split()

# Download headers. I think we want the first conn.uid it shows much less info and has the Subject field
result, data = conn.uid('fetch', ','.join(uids), '(BODY[HEADER.FIELDS (MESSAGE-ID FROM TO CC DATE SUBJECT)])')
#result, data = conn.uid('fetch', ','.join(uids), '(BODY[HEADER])')


# Where data will be stored
raw_file = open('Email_' + folder +'_201504-201506.tsv', 'w')

# Header for TSV file
raw_file.write("Message-ID\tDate\tFrom\tTo\tCc\tSubject\n")

# Parse data and spit out info
# sample record
# From nobody Wed Feb 25 23:27:18 2015
# From: "Maggiano's Little Italy" <Maggianos@email.maggianos.com>
# To: <aysla001@yahoo.com>
# Date: Fri, 30 Jan 2015 14:03:04 -0600
# Message-ID: <97f89d26-e7f7-42a3-bdab-0bad3260836a@xtinp2mta1107.xt.local>


for i in range(0, len(data)):
    
    # If the current item is _not_ an email header
    if len(data[i]) != 2:
        continue
    
    # Okay, it's an email header. Parse it.
    msg = email.message_from_string(data[i][1])
    mids = msg.get_all('message-id', None)
    mdates = msg.get_all('date', None)
    senders = msg.get_all('from', [])
    receivers = msg.get_all('to', [])
    ccs = msg.get_all('cc', [])
    subject = msg.get_all('subject', [])
    print subject
    
    row = "\t" if not mids else mids[0] + "\t"
    row += "\t" if not mdates else mdates[0] + "\t"
    
    # Only one person sends an email, but just in case
    for name, addr in getaddresses(senders):
        row += addr + " "
    row += "\t"
    
    # Space-delimited list of those the email was addressed to
    for name, addr in getaddresses(receivers):
        row += addr + " "
    row += "\t"
	
    # Space-delimited list of those who were CC'd
    for name, addr in getaddresses(ccs):
        row += addr + " "
    row += "\t"
	
	# Space-delimited list of subjects
    for name, addr in getaddresses(subject):
        row += str(subject)
        print subject		
    row += "\n"
    
    # DEBUG    
    # print msg.keys()
    
    # Just going to output tab-delimited, raw data. Will process later.
    raw_file.write(row)


# Done with file, so close it
raw_file.close()

    
    
    
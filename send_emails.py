# Import libraries

# For multithreading
import threading
from queue import Queue

# For sending emails
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# For managing PDFs
from PyPDF2 import PdfFileWriter, PdfFileReader

# For finding data from subfolders
import pandas as pd
import os
import traceback
from os import listdir
from os.path import isfile, join


# This sends an email
def send_email(username           = "veconomou@athenscollege.edu.gr",
            password              = "",
            mail_from             = "veconomou@athenscollege.edu.gr",
            mail_to               = "veconomou@athenscollege.edu.gr",
            mail_subject          = "This is a test email!",
            mail_body             = "This is a test body",
            mail_attachments      = ["./attach.jpeg","./attach2.pptx"],
            mail_attachment_names = ["Attach.jpeg","Attach2.pptx"],
            mail_cc               = None,
            mail_bcc              = None,
            print_func            = print):

            # Set email basics
            mimemsg             = MIMEMultipart()
            mimemsg['From']     = mail_from
            mimemsg['To']       = mail_to
            mimemsg['Subject']  = mail_subject
            
            # Check for cc and bcc
            if mail_cc  is not None: mimemsg['Cc']  = mail_cc
            if mail_bcc is not None: mimemsg['Bcc'] = mail_bcc

            # Construct the email
            mimemsg.attach(MIMEText(mail_body,'html'))

            # This attaches the files
            if mail_attachments is not None:
                for mail_attachment, mail_attachment_name in zip(mail_attachments,mail_attachment_names):
                    with open(mail_attachment,"rb") as attachment:

                        # Attachment stuffs
                        mimefile = MIMEBase('application', 'octet-stream')
                        mimefile.set_payload((attachment).read())
                        encoders.encode_base64(mimefile)
                        mimefile.add_header('Content-Disposition', 'attachment', filename=str(mail_attachment_name))
                        mimemsg.attach(mimefile)
                    
            # Connect and send
            connection = smtplib.SMTP(host='smtp.office365.com', port=587)
            connection.starttls()
            connection.login(username,password)
            connection.send_message(mimemsg)
            connection.quit()

            # print_func("\tEmail to: "+mail_to+" sent.")
            print("\te-mail to: "+mail_to+" sent.")


# Obtain folfer list from excel
def get_folder_list(recipient_list_dir:str = './List.xls'):
    # Read the Raw Excel and convert to dictionary
    raw_excel = pd.read_excel(recipient_list_dir)
    recipient_list = raw_excel.to_dict()
    for key in recipient_list.keys(): recipient_list[key] = list(recipient_list[key].values())

    # Return folder names
    return [str(f) for f in recipient_list['FOLDER-NAME']]

# This function will collect the attachments given a list and a folder directory
def collect_attachments(files_directory     = './STUDENTS',
                        Attachements_yn     = True,
                        recipient_list_dir  = './List.xls',
                        VERBOSE             = True):

    print('---> BEGIN')
    print('-------------------------------')
    # If there are no files just return recepient list
    if  not Attachements_yn :
        print('There are no attachments needed')
        try:
            # Read recipient list excel file and save as a dictionary
            raw_excel = pd.read_excel(recipient_list_dir)
            recipient_list = raw_excel.to_dict()
            for key in recipient_list.keys(): recipient_list[key] = list(recipient_list[key].values())
            print('--> Recipient list READ SUCCESSFULLY')

        except:
            print('--> Recepient list FAILED TO READ')

        return [],[],recipient_list,[True]*len(raw_excel)

    # Collect the files
    raw_excel           = pd.read_excel(recipient_list_dir)
    recipient_list      = raw_excel.to_dict()
    for key in recipient_list.keys(): recipient_list[key] = list(recipient_list[key].values())
    student_folders     = [x[0] for x in os.walk(files_directory)]
    success             = [False]* len(raw_excel)
    attachment_names    = [None] * len(raw_excel)
    attachments         = [None] * len(raw_excel)

    # Do a bit of input verification
    if files_directory[-1]!='/': files_directory=files_directory+'/'

    # For all of them get their attachment
    # Check if there are folder names:
    if 'FOLDER-NAME' not in recipient_list.keys(): 
        print('--> No FOLDER-NAME column exists on the header of the recepient list. Sending emails without attachments.')
        return [],[],recipient_list,[True]*len(raw_excel)

    # Otherwise check for the rest of the attachments    
    for i,folder_name in enumerate(recipient_list['FOLDER-NAME']):
        if files_directory+str(folder_name) in student_folders:
            attachment_names[i] = [att for att in listdir(files_directory+str(folder_name)) if (isfile(join(files_directory+str(folder_name), att)) and (att[0:1]!='.'))]
            attachments[i] = [files_directory+str(folder_name)+'/'+attachment_name for attachment_name in attachment_names[i]]
            success[i] = True

    # If VERBOSE print the files
    if VERBOSE:
        if False not in success:
            for i in range(len(success)):
                print("\tFound:\t"+files_directory+str(recipient_list['FOLDER-NAME'][i]))
            print("--> ALL ATTACHEMENTS FOUND SUCCESSFULLY ")
        else:
            print("--> I couldn't an attachment folder for the following emails...")
            for i in range(len(success)):
                if success[i] == False:
                    print("\tNOT FOUND!\tROW: %4d\tEMAIL: "%(i+1)+recipient_list['EMAIL'][i]+"\tFOLDER-NAME in column D:\t"+files_directory+str(recipient_list['FOLDER-NAME'][i]))

    return attachments, attachment_names, recipient_list, success


# Create email queue
def get_email_queue(attachments         = ["./attach1.jpeg","./attach2.pptx"],
                    attachment_names    = ["./attach1.jpeg","./attach2.pptx"],
                    recipient_list      = ["./attach1.jpeg","./attach2.pptx"],
                    mail_body_raw       = "./mail_body.html",
                    username            = "veconomou@athenscollege.edu.gr",
                    password            = "",
                    mail_from           = "veconomou@athenscollege.edu.gr",
                    mail_subject        = "This is a test email!",
                    print_func          = print):

    # Create an empty queue
    jobs = Queue()

    # Number of entries
    emails_num = len(list(recipient_list.values())[0])


    # For all the emails create a dictionary of the parameters and add it to the queue
    for i in range(emails_num):
        
        # Get the mail body copy
        mail_body = mail_body_raw
        
        # Replace the keys
        for key in recipient_list.keys():
            mail_body = mail_body.replace('{'+key+'}',str(recipient_list[key][i]))

        # unpack attachemnts
        attachment      = None
        attachment_name = None
        if len(attachments) != 0:
            attachment      = attachments[i]
            attachment_name = attachment_names[i]

        job = {
                'mail_body':mail_body,
                'mail_to':recipient_list['EMAIL'][i],
                'mail_attachments':attachment,
                'mail_attachment_names':attachment_name,
                'username':username,
                'password':password,
                'mail_from':mail_from,
                'mail_subject':mail_subject,
                'mail_cc':None,
                'mail_bcc':None,
                'print_func':print_func
            }

        # Check for CC
        if 'CC' in recipient_list.keys(): 
            if str(recipient_list['CC'][i]) != 'nan':
                job['mail_cc'] = ', '.join(recipient_list['CC'][i].split(';'))

        # Check for BCC
        if 'BCC' in recipient_list.keys(): 
            if str(recipient_list['BCC'][i]) != 'nan':
                job['mail_bcc'] = ', '.join(recipient_list['BCC'][i].split(';'))

        jobs.put(job)

    return jobs

# Worker that sends an email from the queue
def send_from_queue(jobs:Queue):
    # While there are available emails to send
    while not jobs.empty():
        # Get the next available email
        email = jobs.get()

        try:
            # Send the email
            send_email(**email)
        except:
            # If there is an error
            print("--> There was a problem sending to: "+email['mail_to']+". RETRYING...")
            traceback.print_exc()
            jobs.put(email)

        # Tell the queue that this thread stopped
        jobs.task_done()


# Creates NUM_THERADS to send the appropriate emails
def get_workers(jobs:Queue,NUM_THREADS = 100):
    # Create the threads
    threads = []
    for i in range(NUM_THREADS):
        threads.append(threading.Thread(target=send_from_queue,args=(jobs,)))

    return threads


# Split pdf command
def pdf_split(filename:str,filenames:list):
    
    # Load the pdf
    pdf = PdfFileReader(open(filename,'rb'))

    # Check relative sizes
    if pdf.numPages != len(filenames): return -1

    # create a new directory to store the pdf
    os.makedirs(filename+'-split',exist_ok=True)

    # For each pdf page
    for i,fname in zip(range(pdf.numPages),filenames):
        # remove the .pdf from filename
        fname.replace('.pdf','')

        # Create new pdf
        new_pdf = PdfFileWriter()
        new_pdf.add_page(pdf.getPage(i))

        # save the new pdf
        temp_filename = filename+'-split/'+fname+'/'
        
        # Create output file stream
        os.makedirs(temp_filename,exist_ok=True)
        outfile = open(temp_filename+fname+'.pdf','wb')         # Create output file stream
        new_pdf.write(outfile)                                  # Send the pdf to the output file stream
        outfile.close()                                         # Close output file stream

    # Return filename
    return filename+'-split'


# Make folder command
def make_folders_from_xls(mainloldername,recipient_list_dir:str):
    raw_excel = pd.read_excel(recipient_list_dir)
    recipient_list = raw_excel.to_dict()
    # create a new directory 

    # For each pdf page
    for i,foldername in enumerate(recipient_list['FOLDER-NAME']):
        foldername=recipient_list['FOLDER-NAME'][i]
        temp_filename = mainloldername+'/'+foldername+'/'
        # Create output file stream
        os.makedirs(temp_filename,exist_ok=True)
    # Return filenames
    return foldername
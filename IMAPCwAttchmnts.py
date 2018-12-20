from imapclient import IMAPClient
from datetime import datetime
import os
import email


# extracts the body from the email
def get_body(msg):  # Need to use that raw var below
    if msg.is_multipart():  # If the message is a multipart(returns True)
        return get_body(msg.get_payload(0))  # Returns the payload
    else:
        return msg.get_payload(None, True)  # Returns nada


# allows you to download attachments
def get_attachments(msg):

    # Takes the raw data and breaks it into different 'parts' & python processes it 1 at a time [1]
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':  # Checks if the email is the correct 'type'.
                    # If it's a 'multipart', then it is incorrect type of email that can possible have an attachment
            continue  # Continue command skips the rest of code and checks the next 'part'

        if part.get('Content-Disposition') is None:  # Checks the 'Content-Disposition' field of the message.
                            # If it's empty, or "None", then we need to leave and go to the next part
            continue  # Continue command skips the rest of code and checks the next 'part'
        # So if the part isn't a 'multipart' type and has a 'Content-Disposition'...

        file_name = part.get_filename()  # Get the filename

        if bool(file_name):  # If bool(file_name) returns True
            file_path = os.path.join(save_dir, file_name)  # Combine the save directory and file name to make file_path
            with open(file_path, 'wb') as f:  # Opens file, w = creates if it doesn't exist / b = binary mode [2]
                f.write(part.get_payload(decode=True))  # Returns the part is carrying, or it's payload, and decodes [3]


HOST = 'imap.domain.com'  # IMAP Host Server
ism = 'user@domain.com'  # Login
kalima = '123CantGuessMe!'  # PW
subj = 'Test'  # Subject to search for
save_dir = ''  # save directory

MAILBOX = 'INBOX'  # Mailbox to check

ssl = False

with IMAPClient(HOST, ssl=ssl) as server:  # Picking the server and SSL security. and assigning to var.
    # Starts with "with" so that if for whatever reason the code terminates, it will log out of the server

    server.login(ism, kalima)  # Signing in with above variables
    server.select_folder(MAILBOX)  # Selecting mailbox

    messages = server.search([u'SINCE', datetime.today()])  # Searching the server for messages, specifically form today

    print("Total messages: %d " % len(messages))  # Total messages found
    print()
    print("Messages:")

    # Retrieves the messages and puts them in a collection [4][5]
    response = server.fetch(messages, ['FLAGS', 'BODY', 'RFC822.SIZE', 'ENVELOPE', 'RFC822'])

    for msgid, data in response.items():  # Iterates through the collection and assigns to 2 variables one by one
        print('   ID %d: %d bytes, flags=%s' % (msgid,
                                                data[b'RFC822.SIZE'],
                                                data[b'FLAGS']))
        envelope = data[b'ENVELOPE']
        print(envelope.subject.decode())  # Gets the subject

        raw = email.message_from_bytes(data[b'RFC822'])  # Return a message object structure from a bytes-like object[6]
        get_attachments(raw)  # Runs above function

    server.logout()  # Ensures a logout

''' References:
[1] Content-Type & Multipart: https://tools.ietf.org/html/rfc2183
[2] Open Function: https://www.w3schools.com/python/ref_func_open.asp
[3] email.Message - payload: https://docs.python.org/2/library/email.message.html#email.message.Message.get_payload
[4] More on collections: https://docs.python.org/3/library/collections.html
[5] fetch() in IMAPClient: https://media.readthedocs.org/pdf/imapclient/master/imapclient.pdf
[6] email.Message_from_bytes: https://docs.python.org/3/library/email.parser.html
'''

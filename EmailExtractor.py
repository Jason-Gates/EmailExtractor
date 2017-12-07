import imaplib
import email
import re
import hashlib
import json


class EmailExtractor:

    """
    This class retrieves emails from an IMAP server and pushes metadata into MongoDB
    """

    ignored_links = ['http://schemas.openxmlformats.org']

    def __init__(self, username, password, hostname, port, ssl, mailbox="inbox"):
        # Initialize the class variables
        self.conn = None
        self.username = username
        self.password = password
        self.hostname = hostname
        self.port = port
        self.ssl = ssl
        self.mailbox = mailbox

    def connect(self):
        #  Connect to the IMAP server
        try:
            if self.ssl:
                self.conn = imaplib.IMAP4_SSL(host=self.hostname, port=self.port)
            else:
                self.conn = imaplib.IMAP4(host=self.hostname, port=self.port)
        except imaplib.IMAP4.error:
            print "Unable to connect to IMAP server"
            return False

        try:
            self.conn.login(self.username, self.password)
        except imaplib.IMAP4.error:
            print "Invalid login credentials"
            return False

        try:
            self.conn.select(self.mailbox)
        except imaplib.IMAP4.error:
            print "Unable to select mailbox: " + self.mailbox
            return False
        #  Return True if all functions executed correctly
        return True

    def _detect_links(self, input):
        # TODO: Improve the url detection
        p = re.compile("(https?:\/\/)")
        results = p.findall(input)
        return len(results)

    def _parse_multiparts(self, input):
        # Parse multiparts (attachments, text/plain, text/html etc)
        msg_parts = []
        for part in input.walk():
            # Loop through the email parts
            result_part = {}
            result_part['type'] = part.get_content_type()
            # TODO: get metadata from file formats such as exe, pdf, doc etc.
            part_payload = part.get_payload(decode=True)
            if part_payload:
                result_part['urls'] = self._detect_links(part_payload)
                result_part['filename'] = part.get_filename()
                result_part['hash'] = hashlib.sha256(part_payload).hexdigest()
                msg_parts.append(result_part)
        return msg_parts

    def _parse_message(self, input):
        # Parse metadata from one email message
        metadata = {}
        msg = email.message_from_string(input[0][1])
        metadata['from'] = msg['from']
        metadata['to'] = msg['to']
        metadata['reply_to'] = msg['reply-to']
        metadata['subject'] = msg['subject']
        p = re.compile("\[(.*)\]")  # Extract the source ipaddress from within the brackets
        metadata['source_ip'] = p.findall(msg['received'])
        # Iterate through the message parts and extract some details
        if msg.is_multipart():
            metadata['parts'] = self._parse_multiparts(msg)
        # Add message to MongoDB
        # print json.dumps(message, ensure_ascii=False)
        return metadata

    def process(self, criterion="ALL"):
        #  Iterate through the IMAP emails
        typ, data = self.conn.search(None, criterion)
        for num in data[0].split():
            typ, data = self.conn.fetch(num, '(RFC822)')  # Fetch one message at a time
            msg = self._parse_message(data)  # Parse the message metadata
            # TODO: send the metadata to MongoDB

    def close(self):
        self.conn.close()
        self.conn.logout()

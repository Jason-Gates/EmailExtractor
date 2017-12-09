import imaplib
import email
import re
import hashlib
import json
import zipfile
import magic
from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML
from StringIO import StringIO
from elasticsearch import Elasticsearch
from datetime import datetime


class EmailExtractor:

    """
    This class retrieves emails from an IMAP server and pushes metadata into MongoDB
    """

    ignored_links = ['http://schemas.openxmlformats.org']

    def __init__(self, username, password, hostname, port, ssl, elastic_index, mailbox="inbox"):
        # Initialize the class variables
        self.conn = None
        self.elastic_conn = Elasticsearch()
        self.elastic_index = elastic_index
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
        results = p.findall(str(input))
        return len(results)

    def _deflate_objects(self, input):
        metadata = {}
        obj_type = input[0]
        obj = input[1]
        if obj_type == "text/plain":
            pass
        elif obj_type == "multipart/mixed":
            pass
        elif obj_type == "application/msword":
            vbaparser = VBA_Parser('test.doc', data=obj)
            if vbaparser.detect_vba_macros():
                metadata['macro'] = 'yes'
            else:
                metadata['macro'] = 'no'
            metadata['macros'] = []
            for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
                macro = {}
                macro['OLEstream'] = stream_path
                macro['VBAfile'] = vba_filename
                macro['sha256'] = hashlib.sha256(vba_code).hexdigest()
                metadata['macros'].append(macro)
            vbaparser.analyze_macros()
            metadata['macro_autoexec'] = vbaparser.nb_autoexec
            metadata['macro_suspicious'] = vbaparser.nb_suspicious
            metadata['macro_iocs'] = vbaparser.nb_iocs
            metadata['macro_hex'] = vbaparser.nb_hexstrings
            metadata['macro_base64'] = vbaparser.nb_base64strings
        elif obj_type == "application/octet-stream":
            pass
        elif obj_type == "multipart/alternative":
            pass
        else:
            print obj_type
        return metadata  # list of deflated artifacts

    def _deflate_multiparts(self, input):
        # Parse multiparts (attachments, text/plain, text/html etc)
        msg_parts = []
        for part in input.walk():
            part_metadata = {}
            part_payload = part.get_payload(decode=True)
            if not part_payload:
                continue
            part_metadata['urls'] = self._detect_links(part_payload)
            part_type = magic.from_buffer(part_payload, mime=True)  # Try to rely on file magic
            part_metadata['type'] = part_type
            more_data = self._deflate_objects([part_type, part_payload])
            if more_data:
                part_metadata['parts'] = more_data
            if part_payload:
                part_metadata['filename'] = part.get_filename()
                part_metadata['sha256'] = hashlib.sha256(part_payload).hexdigest()
            msg_parts.append(part_metadata)
        return msg_parts

    def _parse_message(self, input):
        # Parse metadata from one email message
        msg_metadata = {}
        msg = email.message_from_string(input[0][1])
        if msg['from']:
            msg_metadata['from'] = unicode(msg['from'], errors='ignore')
        if msg['to']:
            msg_metadata['to'] = unicode(msg['to'], errors='ignore')
        if msg['reply-to']:
            msg_metadata['reply_to'] = unicode(msg['reply-to'], errors='ignore')
        if msg['subject']:
            msg_metadata['subject'] = unicode(msg['subject'], errors='ignore')
        # TODO: implement datetime parsing from email
        msg_metadata['timestamp'] = datetime.now().isoformat()
        p = re.compile("\[(.*)\]")  # Extract the source ipaddress from within the brackets
        msg_metadata['source_ip'] = p.findall(msg['received'])
        # Iterate through the message parts and extract some details
        if msg.is_multipart():
            msg_metadata['parts'] = self._deflate_multiparts(msg)
        # Add message to MongoDB
        # print json.dumps(message, ensure_ascii=False)
        return msg_metadata

    def process(self, criterion="ALL"):
        #  Iterate through the IMAP emails
        typ, data = self.conn.search(None, criterion)
        for num in data[0].split():
            typ, data = self.conn.fetch(num, '(RFC822)')  # Fetch one message at a time
            msg = self._parse_message(data)  # Parse the message metadata
            self.elastic_conn.index(index=self.elastic_index, doc_type="external", body=msg)

    def close(self):
        # Close the IMAP connection
        self.conn.close()
        self.conn.logout()

## Overview
The purpose of this Python class is to help extract email metadata from spamtraps
and to help make sense of large amounts of spam/malware/phishing emails by leveraging elasticsearch (ELK)

## Current features
* Reads emails via IMAP
* Reads emails based on IMAP search filter
* Designed to recursively "deflate" email objects
* Can parse Microsoft Word documents and macros (with oletools)
* Extracts common email headers
* Counts number of URLs in email objects
* Saves email metadata to elasticsearch instance

## Future features
* More mime type support (eg zip, 7z)
* Attempt to use passwords given in emails to decrypt attachments
* Moving processed messages to another IMAP folder
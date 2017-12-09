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

## Example output

    {
        'from': u 'Graham Tateham <[removed]@[removed].com>',
        'timestamp': '2017-12-09T18:24:17.063213',
        'source_ip': ['187.162.x.x'],
        'to': u '<[removed]@[removed].com>',
        'parts': [{
            'hash': '7eb70257593da06f682a3ddda54a9d260d4fc514f645237f5ca74b08f8da61a6',
            'type': 'text/plain',
            'urls': 0,
            'filename': None
        }, {
            'hash': '179506c7453ac34adbd2a2286844327b65346bf56c4e0b12a1722dfd89ce3649',
            'parts': {
                'macro_suspicious': 4,
                'macro_iocs': 0,
                'macro_base64': 432,
                'macro': 'yes',
                'macros': [{
                    'VBAfile': 'ThisDocument.cls',
                    'sha256': '76f615001478cc8d4cce6fe76f75b2ca62880736645e205dfce15fc20f92e650',
                    'OLEstream': u 'Macros/VBA/ThisDocument'
                }, {
                    'VBAfile': 'AwhuaNEjIfhpQd.bas',
                    'sha256': '4e17e1c299d012c98557b479891e151d48175bde5b2ba346a9ae38da36d12f7e',
                    'OLEstream': u 'Macros/VBA/AwhuaNEjIfhpQd'
                }, {
                    'VBAfile': 'iRjIarSuwwZlYb.bas',
                    'sha256': '6967b7afd0f1e5371e893d2bcc9c9c908b8b20d0c39f623a23a02e432e3db04b',
                    'OLEstream': u 'Macros/VBA/iRjIarSuwwZlYb'
                }, {
                    'VBAfile': 'wzDGpvjHvsTu.bas',
                    'sha256': 'b5f71446b8dd3172e40efb7c83b5dea8a93a2e404bc759c5833770fcc15180e2',
                    'OLEstream': u 'Macros/VBA/wzDGpvjHvsTu'
                }],
                'macro_autoexec': 1,
                'macro_hex': 1
            },
            'type': 'application/msword',
            'urls': 1,
            'filename': 'INV0000238336.doc'
        }, {
            'hash': '6845799a95619e241ab988b4f707f52ffbb782e53957504810828550015e8235',
            'type': 'text/plain',
            'urls': 0,
            'filename': None
        }],
        'subject': u 'INV0000238336'
    }
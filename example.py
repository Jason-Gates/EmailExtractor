#!/usr/bin/env python
import EmailExtractor

srv = EmailExtractor.EmailExtractor(username="testuser",
                                  password="testpassword",
                                  hostname="localhost",
                                  port="993",
                                  ssl=True)
if srv.connect():
    srv.process(criterion="ALL")
    srv.close()

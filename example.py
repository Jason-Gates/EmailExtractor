#!/usr/bin/env python
import EmailExtractor

srv = EmailExtractor.EmailExtractor(username="test",
                                  password="test",
                                  hostname="localhost",
                                  port="993",
                                  elastic_index="spamtrap",
                                  ssl=True)
if srv.connect():
    srv.process(criterion="ALL", move_message=True)
    srv.close()
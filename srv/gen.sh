#!/bin/bash
openssl genrsa 2048 > server.key
openssl req -new -key server.key -out server.csr
openssl x509 -req -in server.csr -signkey server.key -out server.crt -days 3650 -sha256 -extfile v3.ext

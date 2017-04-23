#!/bin/bash

# Exit immediately if a command exits with a non-zero status
set -e

echo "This script will generate the private keys, the certificate signing request and the certificates that are needed for running SWordy by using the local HTTPS server."
echo "If these files are in current folder they will be overwritten; in this case, and if you already installed the swordy_ca.crt certificate, you must install is again in your Trusted Root Certification Authorities store in order to continue to use SWordy."
read -p "Do you want to continue? [y/n] " response
if [ "$response" != "y" ]; then
	exit
fi

echo ""

# Check if OpenSSL (https://www.openssl.org/) is installed in the system
command -v openssl >/dev/null 2>&1 || { echo >&2 "Error: OpenSSL not found."; exit 1; }

# Generate the private RSA key for the SWordy certification authority
openssl genrsa -out swordy_ca.key 2048

# Generate the self-signed certificate for the SWordy certification authority		
openssl req -new -x509 -days 1826 -key swordy_ca.key -out swordy_ca.crt -subj "/CN=SWordy certification authority"

# Generate the private RSA key for the server
openssl genrsa -out server.key 4096

# Generate the certificate signing request for the server
openssl req -new -key server.key -out server.csr -subj "/CN=localhost"

# Generate the certificate for the server
openssl x509 -req -days 1826 -in server.csr -CA swordy_ca.crt -CAkey swordy_ca.key -set_serial 01 -out server.crt

echo ""
echo "The following files were generated:"
echo ""
echo "swordy_ca.key (the private RSA key for the SWordy certification authority)"
echo "swordy_ca.crt (the self-signed certificate for the SWordy certification authority)"
echo "server.key (the private RSA key for the server)"
echo "server.csr (the certificate signing request for the server)"
echo "server.crt (the certificate for the server)"
echo ""
echo "Notes:"
echo ""
echo "- If you want to use SWordy from the local HTTPS server, you must install the root CA (swordy_ca.crt) in your Trusted Root Certification Authorities store;"
echo "- You must keep secret the two private RSA keys;"
echo "- You can delete the server.csr file;"
echo "- The SWordy certification authority certificate and the server certificate have a validity of about 5 years (1826 days). After, you must generate again the certificates and reinstall the root CA in your system."



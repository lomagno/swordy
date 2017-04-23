/* global __dirname */

// Requirements
var path = require('path');
var fs = require('fs');
var https = require('https');
var express = require('express');

// Check if server.crt exists
if (!fs.existsSync('server.crt')) {
    console.error(
        'Can not start the HTTPS server because the server certificate is missing.' +
        ' SWordy expects to find the SSL certificate in the server.crt file which must be in the SWordy root folder.' + 
        ' You can generate this file, together with other required files, by using the \"generate_certificates.sh\" script.' +
        ' In order to execute this script you need a Bash shell (in Windows you can use GIT bash) and OpenSSL (https://www.openssl.org/).' +
        ' You can find more information on the README file.');
    process.exit();
}

// Check if server.key exists
if (!fs.existsSync('server.key')) {
    console.error(
        'Can not start the HTTPS server because the server private key is missing.' +
        ' SWordy expects to find the private key in the server.key file which must be in the SWordy root folder.' + 
        ' You can generate this file, together with other required files, by using the \"generate_certificates.sh\" script.' +
        ' In order to execute this script you need a Bash shell (in Windows you can use GIT bash) and OpenSSL (https://www.openssl.org/).' +
        ' You can find more information on the README file.');
    process.exit();
}

// Express web framework
var app = express();

// Options
var options = {
    hostname: 'localhost',
    key: fs.readFileSync('server.key'),
    cert: fs.readFileSync('server.crt') /*,
    ca: fs.readFileSync('ca.crt') TODO: include this? */ 
};
var port = 3000;

// Public folder
app.use(express.static(path.join(__dirname, 'webapp')));

// HTTPS server
var httpsServer = https.createServer(options, app);
httpsServer.on('error', function (e) {
    var errorCode = e.code;
    if (errorCode === 'EADDRINUSE')
        console.log('Cannot bind to port ' + port);
    else
        console.error(e.code);
});
httpsServer.listen(port, function() {
    console.log('Listening on https://localhost:' + port + '...');
});


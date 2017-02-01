/* global __dirname */

// Requirements
var path = require('path');
var fs = require('fs');
var https = require('https');
var express = require('express');

// Express web framework
var app = express();

// Options
var options = {
    hostname: 'localhost',
    key: fs.readFileSync('server.key'),
    cert: fs.readFileSync('server.crt'),
    ca: fs.readFileSync('ca.crt')
};
var port = 8088;

// Public folder
app.use(express.static(__dirname  + '/../public_html'));

// HTTPS server
https.createServer(options, app).listen(port, function() {
    console.log('Listening on https://localhost:' + port + '...');
});

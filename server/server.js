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
var port = 3000;

// Public folder
app.use(express.static(path.join(__dirname, '..', 'public_html')));

// HTTPS server
var httpsServer = https.createServer(options, app)
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


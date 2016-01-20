var path = require('path'),
    xml2js = require('xml2js');

exports.client = null;

exports.initialize = function(settings, callback) {
    var soap = require('soap');
    var endpoint = 'https://' + path.join(settings.url, 'EWS/Exchange.asmx');
    var url = path.join(__dirname, 'Services.wsdl');

    soap.createClient(url, {}, function(err, client) {
        if (err) {
            return callback(err);
        }
        if (!client) {
            return callback(new Error('Could not create client'));
        }

        exports.client = client;

        if (settings.token) {
            exports.client.setSecurity(new soap.BearerSecurity(settings.token));
        }
        else {
            exports.client.setSecurity(new soap.BasicAuthSecurity(settings.username, settings.password));
        }

        return callback(null);
    }, endpoint);
}

exports.installApp = function(manifest, callback) {
    if (!exports.client) {
        return callback(new Error('Call initialize()'));
    }

    var soapRequest =
        '<tns:InstallApp xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:Manifest>' + manifest + '</tns:Manifest>' +
        '</tns:InstallApp>';

    exports.client.InstallApp(soapRequest, function(err, result, body) {
        if (err) {
            return callback(err);
        }

        var parser = new xml2js.Parser(
            {
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": '@'
            });

        parser.parseString(body, function(err, result) {
            var responseCode = result['s:Body']['InstallAppResponse']['ResponseCode']

            if (responseCode !== 'NoError') {
                return callback(new Error(responseCode));
            }

            callback(null);
        });
    });
};
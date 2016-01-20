var path = require('path'),
     crypto = require('crypto'),
     xml2js = require('xml2js');

var util = require('util');

exports.client = null;

exports.initialize = function (settings, callback) {
    var soap = require('soap');
    // TODO: Handle different locations of where the asmx lives.
    var endpoint = 'https://' + path.join(settings.url, 'EWS/Exchange.asmx');
    var url = path.join(__dirname, 'Services.wsdl');

    soap.createClient(url, {}, function (err, client) {
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

exports.installApp = function (manifest, callback) {
    if (!exports.client) {
        return callback(new Error('Call initialize()'));
    }

    var soapRequest =
        '<tns:InstallApp xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:Manifest>' + manifest + '</tns:Manifest>' +
        '</tns:InstallApp>';

    exports.client.InstallApp(soapRequest, function (err, result, body) {
        if (err) {
            return callback(err);
        }

        var parser = new xml2js.Parser(
            {
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": '@'
            });

        parser.parseString(body, function (err, result) {
            var responseCode = result['s:Body']['InstallAppResponse']['ResponseCode']

            if (responseCode !== 'NoError') {
                return callback(new Error(responseCode));
            }

            callback(null);
        });
    });
};

exports.getEmails = function (folderName, limit, callback) {
    if (typeof (folderName) === "function") {
        callback = folderName;
        folderName = 'inbox';
        limit = 10;
    }
    if (typeof (limit) === "function") {
        callback = limit;
        limit = 10;
    }
    if (!exports.client) {
        return callback(new Error('Call initialize()'));
    }

    var soapRequest =
        '<tns:FindItem Traversal="Shallow" xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:ItemShape>' +
        '<t:BaseShape>IdOnly</t:BaseShape>' +
        '<t:AdditionalProperties>' +
        '<t:FieldURI FieldURI="item:ItemId"></t:FieldURI>' +
        // '<t:FieldURI FieldURI="item:ConversationId"></t:FieldURI>' +
        // '<t:FieldURI FieldURI="message:ReplyTo"></t:FieldURI>' +
        // '<t:FieldURI FieldURI="message:ToRecipients"></t:FieldURI>' +
        // '<t:FieldURI FieldURI="message:CcRecipients"></t:FieldURI>' +
        // '<t:FieldURI FieldURI="message:BccRecipients"></t:FieldURI>' +
        '<t:FieldURI FieldURI="item:DateTimeCreated"></t:FieldURI>' +
        '<t:FieldURI FieldURI="item:DateTimeSent"></t:FieldURI>' +
        '<t:FieldURI FieldURI="item:HasAttachments"></t:FieldURI>' +
        '<t:FieldURI FieldURI="item:Size"></t:FieldURI>' +
        '<t:FieldURI FieldURI="message:From"></t:FieldURI>' +
        '<t:FieldURI FieldURI="message:IsRead"></t:FieldURI>' +
        '<t:FieldURI FieldURI="item:Importance"></t:FieldURI>' +
        '<t:FieldURI FieldURI="item:Subject"></t:FieldURI>' +
        '<t:FieldURI FieldURI="item:DateTimeReceived"></t:FieldURI>' +
        '</t:AdditionalProperties>' +
        '</tns:ItemShape>' +
        '<tns:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="10"></tns:IndexedPageItemView>' +
        '<tns:ParentFolderIds>' +
        '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' +
        '</tns:ParentFolderIds>' +
        '</tns:FindItem>';

    exports.client.FindItem(soapRequest, function (err, result, body) {
        if (err) {
            return callback(err);
        }

        var parser = new xml2js.Parser(
            {
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": '@'
            });

        parser.parseString(body, function (err, result) {
            var responseCode = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:ResponseCode'];

            if (responseCode !== 'NoError') {
                return callback(new Error(responseCode));
            }

            var rootFolder = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:RootFolder'];

            var emails = [];
            rootFolder['t:Items']['t:Message'].forEach(function (item, idx) {
                var md5hasher = crypto.createHash('md5');
                md5hasher.update(item['t:Subject'] + item['t:DateTimeSent']);
                var hash = md5hasher.digest('hex');

                var itemId = {
                    id: item['t:ItemId']['@'].Id,
                    changeKey: item['t:ItemId']['@'].ChangeKey
                };

                var dateTimeReceived = item['t:DateTimeReceived'];

                emails.push({
                    id: itemId.id + '|' + itemId.changeKey,
                    hash: hash,
                    subject: item['t:Subject'],
                    size: item['t:Size'],
                    importance: item['t:Importance'],
                    hasAttachments: (item['t:HasAttachments'] === 'true'),
                    from: item['t:From']['t:Mailbox']['t:Name'],
                    isRead: (item['t:IsRead'] === 'true'),
                    meta: {
                        itemId: itemId
                    }
                });
            });

            callback(null, emails);
        });
    });
}


exports.getEmail = function (itemId, callback) {
    if (!exports.client) {
        return callback(new Error('Call initialize()'))
    }
    if ((!itemId['id'] || !itemId['changeKey']) && itemId.indexOf('|') > 0) {
        var s = itemId.split('|');

        itemId = {
            id: itemId.split('|')[0],
            changeKey: itemId.split('|')[1]
        };
    }

    if (!itemId.id || !itemId.changeKey) {
        return callback(new Error('Id is not correct.'));
    }

    var soapRequest =
        '<tns:GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '<tns:ItemShape>' +
        '<t:BaseShape>Default</t:BaseShape>' +
        '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
        '</tns:ItemShape>' +
        '<tns:ItemIds>' +
        '<t:ItemId Id="' + itemId.id + '" ChangeKey="' + itemId.changeKey + '" />' +
        '</tns:ItemIds>' +
        '</tns:GetItem>';

    exports.client.GetItem(soapRequest, function (err, result, body) {
        if (err) {
            return callback(err);
        }

        var parser = new xml2js.Parser();

        parser.parseString(body, function (err, result) {
            var responseCode = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:ResponseCode'];

            if (responseCode !== 'NoError') {
                return callback(new Error(responseCode));
            }

            var item = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:Items']['t:Message'];

            var itemId = {
                id: item['t:ItemId']['@'].Id,
                changeKey: item['t:ItemId']['@'].ChangeKey
            };

            function handleMailbox(mailbox) {
                var mailboxes = [];

                if (!mailbox || !mailbox['t:Mailbox']) {
                    return mailboxes;
                }
                mailbox = mailbox['t:Mailbox'];

                function getMailboxObj(mailboxItem) {
                    return {
                        name: mailboxItem['t:Name'],
                        emailAddress: mailboxItem['t:EmailAddress']
                    };
                }

                if (mailbox instanceof Array) {
                    mailbox.forEach(function (m, idx) {
                        mailboxes.push(getMailboxObj(m));
                    })
                } else {
                    mailboxes.push(getMailboxObj(mailbox));
                }

                return mailboxes;
            }

            var toRecipients = handleMailbox(item['t:ToRecipients']);
            var ccRecipients = handleMailbox(item['t:CcRecipients']);
            var from = handleMailbox(item['t:From']);

            var email = {
                id: itemId.id + '|' + itemId.changeKey,
                subject: item['t:Subject'],
                bodyType: item['t:Body']['@']['t:BodyType'],
                body: item['t:Body']['#'],
                size: item['t:Size'],
                dateTimeSent: item['t:DateTimeSent'],
                dateTimeCreated: item['t:DateTimeCreated'],
                toRecipients: toRecipients,
                ccRecipients: ccRecipients,
                from: from,
                isRead: (item['t:IsRead'] == 'true') ? true : false,
                meta: {
                    itemId: itemId
                }
            };

            callback(null, email);
        });
    });
}


exports.getFolders = function (id, callback) {
    if (typeof (id) == 'function') {
        callback = id;
        id = 'inbox';
    }

    var soapRequest =
        '<tns:FindFolder xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:FolderShape>' +
        '<t:BaseShape>Default</t:BaseShape>' +
        '</tns:FolderShape>' +
        '<tns:ParentFolderIds>' +
        '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' +
        '</tns:ParentFolderIds>' +
        '</tns:FindFolder>';

    exports.client.FindFolder(soapRequest, function (err, result) {
        if (err) {
            callback(err)
        }

        if (result.ResponseMessages.FindFolderResponseMessage.ResponseCode == 'NoError') {
            var rootFolder = result.ResponseMessages.FindFolderResponseMessage.RootFolder;

            rootFolder.Folders.Folder.forEach(function (folder) {
                // console.log(folder);
            });

            callback(null, {});
        }
    });
}

exports.getUserAvailability = function (users, callback) {
    if (typeof (users) === "function") {
        callback = users;
        users = [];
    }

    if (!exports.client) {
        return callback(new Error('Call initialize()'));
    }

    var mailboxSoapTemplate = '<t:MailboxData>' +
        '<t:Email>' +
        '<t:Address>%s</t:Address>' +
        '</t:Email>' +
        '<t:AttendeeType>Optional</t:AttendeeType>' +
        '<t:ExcludeConflicts>false</t:ExcludeConflicts>' +
        '</t:MailboxData>';

    var mailboxesSoap = "";

    // users.forEach(function (u) {
    //     mailboxesSoap += util.format(mailboxSoapTemplate, u);
    // });

    var soapRequest =
        '<m:GetUserAvailabilityRequest xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '<m:MailboxDataArray>' +
        mailboxesSoap +
        '</m:MailboxDataArray>' +
        '<t:FreeBusyViewOptions>' +
        '<t:TimeWindow>' +
        '<t:StartTime>2015-11-04T00:00:00</t:StartTime>' +
        '<t:EndTime>2015-11-05T00:00:00</t:EndTime>' +
        '</t:TimeWindow>' +
        '<t:MergedFreeBusyIntervalInMinutes>30</t:MergedFreeBusyIntervalInMinutes>' +
        '<t:RequestedView>Detailed</t:RequestedView>' +
        '</t:FreeBusyViewOptions>' +
        '</m:GetUserAvailabilityRequest>'
  
    // exports.client.addSoapHeader(rsv, "RequestServerVersion", "t", "");
    exports.client.addSoapHeader('<t:RequestServerVersion Version="Exchange2013_SP1" /><t:TimeZoneContext><t:TimeZoneDefinition Name="(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna" Id="W. Europe Standard Time"><t:Periods><t:Period Bias="-P0DT1H0M0.0S" Name="Standard" Id="Std" /><t:Period Bias="-P0DT2H0M0.0S" Name="Daylight" Id="Dlt/1" /></t:Periods><t:TransitionsGroups><t:TransitionsGroup Id="0"><t:RecurringDayTransition><t:To Kind="Period">Dlt/1</t:To><t:TimeOffset>P0DT2H0M0.0S</t:TimeOffset><t:Month>3</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek><t:Occurrence>-1</t:Occurrence></t:RecurringDayTransition><t:RecurringDayTransition><t:To Kind="Period">Std</t:To><t:TimeOffset>P0DT3H0M0.0S</t:TimeOffset><t:Month>10</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek><t:Occurrence>-1</t:Occurrence></t:RecurringDayTransition></t:TransitionsGroup></t:TransitionsGroups><t:Transitions><t:Transition><t:To Kind="Group">0</t:To></t:Transition></t:Transitions></t:TimeZoneDefinition></t:TimeZoneContext>');

    exports.client.GetUserAvailability(soapRequest, function (err, result, body) {
        if (err) {
            return callback(err);
        }

        var parser = new xml2js.Parser({
            "explicitArray": false,
            "explicitRoot": false,
            "attrkey": '@'
        });

        parser.parseString(body, function (err, result) {

            var bias = [];

            var users = result['s:Body']['GetUserAvailabilityResponse']['FreeBusyResponseArray']['FreeBusyResponse'];

            if (users.FreeBusyView) {
                try {
                    var b = users['FreeBusyView']['WorkingHours']['TimeZone']['Bias'];
                    bias.push(parseInt(b) * -1 / 60);
                }
                catch (err) {
                }
            }
            else {
                users.forEach(function (user) {
                    try {
                        var b = user['FreeBusyView']['WorkingHours']['TimeZone']['Bias'];
                        bias.push(parseInt(b) * -1 / 60);
                    }
                    catch (err) {
                    }
                });
            }

            callback(null, bias);
        });
    });
}
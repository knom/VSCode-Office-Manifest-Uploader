let path = require("path");
let xml2js = require("xml2js");
let soap = require("soap");
let enumerable = require("linq");

export class EWSClient {
    private client: any = null;

    public initialize(settings: any, callback: any) {
        let endpoint = settings.url + "/EWS/Exchange.asmx";
        let url = path.join(__dirname, "Services.wsdl");

        let options = {
            escapeXML: false
        };

        soap.createClient(url, options, (err:any, client:any) => {
            if (err) {
                return callback(err);
            }
            if (!client) {
                return callback(new Error("Could not create client"));
            }

            this.client = client;

            this.client.addListener('request', function (xml:any) {
               console.log("---REQUEST---");
               console.log(xml);
               console.log("---END of REQUEST---");
            });
            this.client.addListener('response', function (a:any, b:any) {
               console.log("---RESPONSE---");
               console.log("body: " + a);
               console.log("response: " + JSON.stringify(b));
               console.log("---END of RESPONSE---");
            });
            this.client.addListener('soapError', function () {
                console.log("---SOAP ERROR---");
                console.log("---END of SOAP ERROR---");
            });

            if (settings.token) {
                this.client.setSecurity(new soap.BearerSecurity(settings.token));
            }
            else {
                this.client.setSecurity(new soap.BasicAuthSecurity(settings.username, settings.password));
            }

            return callback(null);
        }, endpoint);
    }

    public installApp(manifest:string, callback:any) {
        if (!this.client) {
            return callback(new Error("Call initialize()"));
        }

        let soapRequest =
            //"<tns:InstallApp xmlns:tns='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
            "<Manifest>" + manifest + "</Manifest>"; //+
        //"</tns:InstallApp>";

        this.client.InstallApp(soapRequest, (httpError:any, result:any, rawBody:any) => {
            if (httpError) {
                if (httpError.response.statusCode && (httpError.response.statusCode == 401 || httpError.response.statusCode == 403)) {
                    return callback(new Error(httpError.response.statusCode + ": Unauthorized!"));
                }
                return callback(new Error(httpError));
            }

            let parser = new xml2js.Parser(
                {
                    "explicitArray": false,
                    "explicitRoot": false,
                    "attrkey": "@"
                });

            parser.parseString(rawBody, (err:any, result:any) => {
                let responseCode = result["s:Body"]["InstallAppResponse"]["ResponseCode"];
                let message = "";

                if (responseCode !== "NoError") {
                    try {
                        message = enumerable.from(result["s:Body"]["InstallAppResponse"]["MessageXml"]["t:Value"])
                            .where(function (i:any) { return i["@"]["Name"] === "InnerErrorMessageText"; })
                            .select(function (i:any) { return i["_"]; }).first();
                    } catch (e) {
                        message = "";
                    }
                    finally {
                        return callback(new Error(responseCode + ": " + JSON.stringify(message)));
                    }
                }

                callback(null);
            });
        });
    };

    public uninstallApp(id:string, callback:any) {
        if (!this.client) {
            return callback(new Error("Call initialize()"));
        }

        let soapRequest =
            //"<m:UninstallApp xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
            "<ID>" + id + "</ID>"; // +
        //"</m:UninstallApp>";

        this.client.UninstallApp(soapRequest, (httpError:any, result:any, rawBody:any) => {
            if (httpError) {
                if (httpError.response.statusCode && (httpError.response.statusCode == 401 || httpError.response.statusCode == 403)) {
                    return callback(new Error(httpError.response.statusCode + ": Unauthorized!"));
                }
                return callback(new Error(httpError));
            }

            let parser = new xml2js.Parser(
                {
                    "explicitArray": false,
                    "explicitRoot": false,
                    "attrkey": "@"
                });

            parser.parseString(rawBody, (err:any, result:any) => {
                let responseCode = result["s:Body"]["UninstallAppResponse"]["ResponseCode"];

                let message = "";
                if (responseCode !== "NoError") {
                    try {
                        message = enumerable.from(result["s:Body"]["UninstallAppResponse"]["MessageXml"]["t:Value"])
                            .where(function (i:any) { return i["@"]["Name"] === "InnerErrorMessageText"; })
                            .select(function (i:any) { return i["_"]; }).first();
                    } catch (e) {
                        message = "";
                    }
                    finally {
                        return callback(new Error(responseCode + ": " + JSON.stringify(message)));
                    }
                }

                callback(null);
            });
        });
    };
}	
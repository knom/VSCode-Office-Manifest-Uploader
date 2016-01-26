import * as vscode from "vscode";
import * as q from "q";
import * as fs from "fs";
import * as ews from "./lib/ews-soap/exchangeClient";
import * as xml2js from "xml2js";

export abstract class OfficeApplicationStrategy {
    constructor(public description: string) {
    }

    protected showUserBox(): Thenable<string> {
        let config = vscode.workspace.getConfiguration("officeManifestUploader");
        let userName = config.get<string>("userName");

        if (userName && userName !== "foo@foo.com") {
            let promise = q.defer<string>();
            promise.resolve(userName);
            return promise.promise;
        }

        return vscode.window.showInputBox({
            placeHolder: "user@foo.com",
            prompt: "Enter your username"
        });
    }

    protected showPwdBox(): Thenable<string> {
        let config = vscode.workspace.getConfiguration("officeManifestUploader");
        let password = config.get<string>("password");

        if (password) {
            let promise = q.defer<string>();
            promise.resolve(password);
            return promise.promise;
        }

        return vscode.window.showInputBox({
            password: true,
            placeHolder: "password",
            prompt: "Enter your password"
        });
    }

    protected showServerBox(): Thenable<string> {
        let config = vscode.workspace.getConfiguration("officeManifestUploader");
        let serverUrl = config.get<string>("serverUrl");

        if (serverUrl) {
            let promise = q.defer<string>();
            promise.resolve(serverUrl);
            return promise.promise;
        }

        return vscode.window.showInputBox({
            password: false,
            placeHolder: "mail.office365.com",
            prompt: "Server name"
        });
    }

    public execute() {
        if (!vscode.workspace.rootPath) {
            vscode.window.showErrorMessage("Open a workspace first!");
            return;
        }

        vscode.window.showInputBox({ placeHolder: "manifest.xml", prompt: "Manifest file path", value: "manifest.xml" }).then((filename) => {
            if (filename === undefined)
                return;
            else {
                let path = require("path");
                let f = path.join(vscode.workspace.rootPath, filename);

                vscode.workspace.openTextDocument(f).then((file) => {
                    let truePath = file.fileName;

                    this.showUserBox().then((user) => {
                        this.showPwdBox().then((pwd) => {
                            this.showServerBox().then((server) => {
                                vscode.window.showInformationMessage("Starting " + this.description + "...");

                                let statusBarItem = vscode.window.createStatusBarItem();
                                statusBarItem.text = "$(cloud-upload) " + this.description + "...";
                                statusBarItem.show();

                                this.executeCore(truePath, user, pwd, server).then(() => {
                                    statusBarItem.hide();
                                    vscode.window.showInformationMessage("Succeeded " + this.description + " !");
                                }, (err) => {
                                    statusBarItem.hide();
                                    vscode.window.showErrorMessage("An error occurred while " + this.description + ": " + err);
                                });
                            });
                        });
                    });
                }, (reason) => {
                    vscode.window.showErrorMessage(reason);
                    this.execute();
                    return;
                });
            }
        });
    }

    protected abstract executeCore(filePath: string, userName: string, password: string, serverUrl: string): Thenable<void>;
}

export class InstallApplication extends OfficeApplicationStrategy {
    constructor() {
        super("installing Office-Addin");
    }

    protected executeCore(filePath: string, userName: string, password: string, serverUrl: string): Thenable<void> {
        let promise = q.defer<void>();

        let data = fs.readFileSync(filePath).toString();
        let manifest = new Buffer(data).toString("base64");

        let client = new ews.EWSClient();
        client.initialize({ url: serverUrl, username: userName, password: password },
            (err: any) => {
                client.installApp(manifest, (err: any) => {
                    if (err) {
                        promise.reject(err);
                    }
                    else {
                        promise.resolve();
                    }
                });
            });

        return promise.promise;
    }
}

export class UninstallApplication extends OfficeApplicationStrategy {
    constructor() {
        super("uninstalling Office-Addin");
    }

    protected executeCore(filePath: string, userName: string, password: string, serverUrl: string): Thenable<void> {
        let promise = q.defer<void>();

        let manifestXml = fs.readFileSync(filePath).toString();

        this.getApplicationIdOutofXml(manifestXml).then((appId) => {
            let client = new ews.EWSClient();
            client.initialize({ url: serverUrl, username: userName, password: password },
                (err: any) => {
                    client.uninstallApp(appId, (err: any) => {
                        if (err) {
                            promise.reject(err);
                        }
                        else {
                            promise.resolve();
                        }
                    });
                });
        });

        return promise.promise;
    }

    private getApplicationIdOutofXml(manifestXml: string): Thenable<string> {
        let promise = q.defer();

        let parser = new xml2js.Parser(
            {
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": "@"
            });

        parser.parseString(manifestXml, (err, result) => {
            let id = result["Id"];
            promise.resolve(id);
        });

        return promise.promise;
    };
}
import * as vscode from "vscode";
import * as fs from "fs";
import * as ews from "./lib/ews-soap/exchangeClient";

let Promise = require("promise");

export class ManifestUploader {
    constructor() {
    }

    private showUserBox(): Thenable<string> {
        let config = vscode.workspace.getConfiguration("officeManifestUploader");
        let userName = config.get<string>("userName");

        if (userName)
        {
            return new Promise((f, r) => {
               f(userName);
            });
        }

        return vscode.window.showInputBox({
                placeHolder: "user@foo.com",
                prompt: "Enter your username"
            });
    }

    private showPwdBox(): Thenable<string>
    {
        let config = vscode.workspace.getConfiguration("officeManifestUploader");
        let password = config.get<string>("password");

        if (password)
        {
            return new Promise((f, r) => {
               f(password);
            });
        }

        return vscode.window.showInputBox({
                password: true,
                placeHolder: "password",
                prompt: "Enter your password"
            });
    }

    private showServerBox(): Thenable<string>
    {
        let config = vscode.workspace.getConfiguration("officeManifestUploader");
        let serverUrl = config.get<string>("serverUrl");

        if (serverUrl)
        {
            return new Promise((f, r) => {
                f(serverUrl);
            });
        }

        return vscode.window.showInputBox({
            password: false,
            placeHolder: "mail.office365.com",
            prompt: "Server name:"
        });
    }

    public upload() {
        vscode.workspace.findFiles("manifest.xml", "").then((uris) => {
            if (!uris) {
                vscode.window.showErrorMessage("No manifest.xml file found in the workspace!");
                return;
            }

            if (uris.length > 1) {
                vscode.window.showErrorMessage("More than one manifest.xml file found in the workspace!");
                return;
            }

            vscode.workspace.openTextDocument(uris[0])
            .then((file) => {
               let truePath = file.fileName;

               let data = fs.readFileSync(truePath).toString();
               let manifest = new Buffer(data).toString("base64");

               this.showUserBox().then((user) => {
                    this.showPwdBox().then((pwd) => {
                        this.showServerBox().then((server) => {
                            this.uploadInternal(manifest, user, pwd, server);
                        });
                    });
                });
            }, (reason) => {
                vscode.window.showErrorMessage("Couldn't open file " + uris[0]);
            });
        }, (reason) => {
            vscode.window.showErrorMessage("An error occurred: " + reason);
        });
    }
    private uploadInternal(manifest: string, userName: string, password: string, serverUrl: string){
        vscode.window.showInformationMessage("Starting manifest upload...");

        ews.initialize({ url: serverUrl,
            username: userName, password: password},
            function(err: any) {
                ews.installApp(manifest, function(err: any) {
                    if (err)
                    {
                        vscode.window.showErrorMessage("An error occurred while uploading: " + err);
                    }
                    else{
                        vscode.window.showInformationMessage("Manifest successfully uploaded and installed!");
                    }
                });
            });
    };
}
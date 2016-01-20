import * as vscode from 'vscode';
import * as ews from './lib/ews-soap/exchangeClient';

var Promise = require('promise');

export class ManifestUploader {
    constructor() {
    }
    
    private showUserBox():Thenable<string>{
        var config = vscode.workspace.getConfiguration("officeManifestUploader");
        var userName = config.get<string>("userName");
        
        if (userName)
        {
            return new Promise((f,r)=>{
               f(userName);
            });
        }
        
        return vscode.window.showInputBox({
                placeHolder: "user@foo.com",
                prompt: "Enter your username"
            });
    }
    
    private showPwdBox():Thenable<string>
    {
        var config = vscode.workspace.getConfiguration("officeManifestUploader");
        var password = config.get<string>("password");
        
        if (password)
        {
            return new Promise((f,r)=>{
               f(password);
            });
        }
        
        return vscode.window.showInputBox({
                password: true,
                placeHolder: "password",
                prompt: "Enter your password"
            });
    }
    
    private showServerBox():Thenable<string>
    {
       var config = vscode.workspace.getConfiguration("officeManifestUploader");
        var serverUrl = config.get<string>("serverUrl");
                
        if (serverUrl)
        {
            return new Promise((f,r)=>{
               f(serverUrl);
            });
        }
        
       return vscode.window.showInputBox({
           password: false,
           placeHolder: "mail.office365.com",
           prompt: "Server name:"
        });
    }
    
    public upload(manifest:string) {  
        this.showUserBox().then((user) => {
            this.showPwdBox().then((pwd) => {
                this.showServerBox().then((server)=>{
                    this.uploadInternal(manifest, user, pwd, server);                        
                });
            });
        });
        
    }
    private uploadInternal(manifest:string, userName:string, password:string, serverUrl:string){
        vscode.window.showInformationMessage("Starting manifest upload...");
            
        ews.initialize({ url: serverUrl, 
            username:userName, password:password},
            function(err: any) {
                ews.installApp(manifest, function(err: any) {
                    if (err)
                    {
                        vscode.window.showErrorMessage('An error occurred: ' + err);
                    }
                    else{
                        vscode.window.showInformationMessage("Manifest successfully uploaded and installed!");
                    }
                });
            });
    };
}
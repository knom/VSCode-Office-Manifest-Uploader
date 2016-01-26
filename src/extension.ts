// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as installApp from './installApplication';

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

    console.log('Extension manifest uploader loaded');

    let cmd1 = vscode.commands.registerCommand('officeAppExt.installApp', () => {
        console.log("Executing command installApp");
        let mu = new installApp.InstallApplication();
        mu.execute();
    });

    let cmd2 = vscode.commands.registerCommand('officeAppExt.uninstallApp', () => {
        console.log("Executing command uninstallApp");
        let mu = new installApp.UninstallApplication();
        mu.execute();
    });

    context.subscriptions.push(cmd1);
    context.subscriptions.push(cmd2);
}

// this method is called when your extension is deactivated
export function deactivate() {
}
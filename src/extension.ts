// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as manifestUploader from './manifestuploader';

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

    console.log('Extension manifest uploader loaded');

    let cmd1 = vscode.commands.registerCommand('extension.uploadManifestXml', () => {
        let mu = new manifestUploader.ManifestUploader();
        mu.upload();
    });

    context.subscriptions.push(cmd1);
}

// this method is called when your extension is deactivated
export function deactivate() {
}
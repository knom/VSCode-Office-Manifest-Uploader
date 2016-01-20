// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as manifestUploader from './manifestuploader';

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

    console.log('Congratulations, your extension "office-mailappuploader" is now active!'); 

    // The command has been defined in the package.json file
    // Now provide the implementation of the command with  registerCommand
    // The commandId parameter must match the command field in package.json

    var disposable = vscode.commands.registerCommand('extension.sayHello', () => {
        // Display a message box to the user
        vscode.window.showInformationMessage('Hello World!');
    });
    context.subscriptions.push(disposable);

    var cmd2 = vscode.commands.registerCommand('extension.uploadManifestXml', () => {
        var manifest = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0KPE9mZmljZUFwcCB4bWxucz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2UvYXBwZm9yb2ZmaWNlLzEuMSIgeG1sbnM6eHNpPSJodHRwOi8vd3d3LnczLm9yZy8yMDAxL1hNTFNjaGVtYS1pbnN0YW5jZSIgeHNpOnR5cGU9Ik1haWxBcHAiPg0KICA8SWQ+ZTdjNWI5OTYtYTBiNC00MjZiLTk3ZWItMmU3ZmY0OTFiZjc5PC9JZD4NCiAgPFZlcnNpb24+MS4wLjAuMDwvVmVyc2lvbj4NCiAgPFByb3ZpZGVyTmFtZT5NYXggS25vcjwvUHJvdmlkZXJOYW1lPg0KICA8RGVmYXVsdExvY2FsZT5lbi1VUzwvRGVmYXVsdExvY2FsZT4NCiAgPERpc3BsYXlOYW1lIERlZmF1bHRWYWx1ZT0iSWJhbiBNYXRlIi8+DQogIDxEZXNjcmlwdGlvbiBEZWZhdWx0VmFsdWU9IkliYW4gTWF0ZSBBcHBsaWNhdGlvbiIvPg0KICA8SGlnaFJlc29sdXRpb25JY29uVXJsIERlZmF1bHRWYWx1ZT0iaHR0cHM6Ly9mb28iLz4NCiAgPFN1cHBvcnRVcmwgRGVmYXVsdFZhbHVlPSJodHRwOi8vZm9vIi8+DQogIDxIb3N0cz4NCiAgICA8SG9zdCBOYW1lPSJNYWlsYm94Ii8+DQogIDwvSG9zdHM+DQogIDxSZXF1aXJlbWVudHM+DQogICAgPFNldHM+DQogICAgICA8U2V0IE5hbWU9Ik1haWxCb3giIE1pblZlcnNpb249IjEuMSIvPg0KICAgIDwvU2V0cz4NCiAgPC9SZXF1aXJlbWVudHM+DQogIDxGb3JtU2V0dGluZ3M+DQogICAgPEZvcm0geHNpOnR5cGU9Ikl0ZW1SZWFkIj4NCiAgICAgIDxEZXNrdG9wU2V0dGluZ3M+DQogICAgICAgIDxTb3VyY2VMb2NhdGlvbiBEZWZhdWx0VmFsdWU9Imh0dHBzOi8vbG9jYWxob3N0Ojg0NDMvYXBwcmVhZC9ob21lL2hvbWUuaHRtbCIvPg0KICAgICAgICA8UmVxdWVzdGVkSGVpZ2h0PjI1MDwvUmVxdWVzdGVkSGVpZ2h0Pg0KICAgICAgPC9EZXNrdG9wU2V0dGluZ3M+DQogICAgPC9Gb3JtPg0KICAgIDxGb3JtIHhzaTp0eXBlPSJJdGVtRWRpdCI+DQogICAgICA8RGVza3RvcFNldHRpbmdzPg0KICAgICAgICA8U291cmNlTG9jYXRpb24gRGVmYXVsdFZhbHVlPSJodHRwczovL2xvY2FsaG9zdDo4NDQzL2FwcGNvbXBvc2UvaG9tZS9ob21lLmh0bWwiLz4NCiAgICAgIDwvRGVza3RvcFNldHRpbmdzPg0KICAgIDwvRm9ybT4NCiAgPC9Gb3JtU2V0dGluZ3M+DQogIDxQZXJtaXNzaW9ucz5SZWFkV3JpdGVJdGVtPC9QZXJtaXNzaW9ucz4NCiAgPFJ1bGUgeHNpOnR5cGU9IlJ1bGVDb2xsZWN0aW9uIiBNb2RlPSJBbmQiPg0KICAgIDxSdWxlIHhzaTp0eXBlPSJSdWxlQ29sbGVjdGlvbiIgTW9kZT0iT3IiPg0KICAgICAgPFJ1bGUgeHNpOnR5cGU9Ikl0ZW1JcyIgSXRlbVR5cGU9Ik1lc3NhZ2UiIEZvcm1UeXBlPSJSZWFkIi8+DQogICAgICA8UnVsZSB4c2k6dHlwZT0iSXRlbUlzIiBJdGVtVHlwZT0iTWVzc2FnZSIgRm9ybVR5cGU9IkVkaXQiLz4NCiAgICA8L1J1bGU+DQogICAgPFJ1bGUgeHNpOnR5cGU9Ikl0ZW1IYXNSZWd1bGFyRXhwcmVzc2lvbk1hdGNoIiANCiAgICAgICAgUmVnRXhOYW1lPSJpYmFuTWF0Y2hlcyIgDQogICAgICAgIFJlZ0V4VmFsdWU9IlthLXpBLVpdezJ9WzAtOV17Mn1bYS16QS1aMC05XXs0fVswLTldezd9KFthLXpBLVowLTldPyl7MCwxNn0iIA0KICAgICAgICBQcm9wZXJ0eU5hbWU9IkJvZHlBc1BsYWludGV4dCIvPg0KICA8L1J1bGU+DQogIDxEaXNhYmxlRW50aXR5SGlnaGxpZ2h0aW5nPmZhbHNlPC9EaXNhYmxlRW50aXR5SGlnaGxpZ2h0aW5nPg0KPC9PZmZpY2VBcHA+";

        var mu = new manifestUploader.ManifestUploader();
        mu.upload(manifest);
    });

    context.subscriptions.push(cmd2);
}

// this method is called when your extension is deactivated
export function deactivate() {
}
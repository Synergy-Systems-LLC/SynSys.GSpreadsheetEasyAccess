## What is GSpreadsheetEasyAccess?

A Google libraries wrapper for an easy and efficient way of interacting with Google Sheets.

## Description

Original libraries have a multitude of types but the
[.NET example](https://developers.google.com/sheets/api/quickstart/dotnet#prerequisites)
highlights the main ones, such as `GoogleWebAuthorizationBroker`, `SheetsService`, `Request`, `ValueRange`.
At the same time, the general Google Sheets API documentation talks about higher level concepts,
namely authentication, the difference between human and service accounts, user scopes, etc.
Original libraries lack the types and methods reflecting these concepts.
We wanted to see them reflected in our library.

### Main types

- `GCPApplication` - an application on the Google Cloud Platform connected to
[Google Sheet API](https://developers.google.com/sheets/api?hl=en_US).  
It can only be used after user authentication.
- `Principal` - represents an object that needs to be granted access to a resource.
- `OAuthSheetsScope` - account scope.
- `SheetModel` - represents a single Google spreadsheet sheet. It consists of rows, while rows consist of cells.

These aren't all the types represented in the library, they aren't important enough to be worth mentioning.

#### GCPApplication

The main library type which allows you to:

- Authenticate users;
- Access an existing Google spreadsheet sheet;
- Create sheet;
- Update sheet.

#### Principal

The basic authentication type. There are other specific types that need to be used in your own code.

- UserAccount.  
User accounts are managed as Google accounts,
they represent a developer, administrator or any other user interacting with Google Cloud.
This is used for scenarios where your application needs to access resources on behalf of a user.  
Scopes need to be set for user accounts.
OAuth scopes limit the actions your application can perform on behalf of the end user.
Set the minimum size needed for your use case.
The class implements only a part of the scopes in
[Google Sheets API v4](https://developers.google.com/identity/protocols/oauth2/scopes#sheets).  
For additional information see
[end user authentication](https://cloud.google.com/docs/authentication/end-user).  

- ServiceAccount.  
Service accounts are managed by
[IAM](https://cloud.google.com/iam/docs/understanding-service-accounts)
and represent non-human users.
This is used for scenarios where your application needs to access resources or
perform actions on its own, for example, launching App Engine applications or interacting
with Compute Engine instances.
Take note that Google only allows service accounts to read public access spreadsheets!
For additional information see
[service account authentication](https://cloud.google.com/docs/authentication/api-keys).

#### SheetModel

A type the state of which should match the desired state of a Google spreadsheet sheet.
This type allows you to:

- Add new rows;
- Delete rows;
- Change cell values;
- Check header validity.

All changes reflected in the `SheetModel` instance show the `GCPApplication`
what changes need to be made in a Google spreadsheet sheet.

#### Exceptions

Most methods throw exceptions. This is outlined in each method's documentation.
Exception descriptions are outlined only in the corresponding exception's documentation.
For example, the sheet header check method may throw `InvalidSheetHeadException` if the header lacks some column names.

## Examples

This repository contains a folder named
[examples](https://github.com/Synergy-Systems-LLC/SynSys.GSpreadsheetEasyAccess/tree/master/examples).  
It contains examples of accessing and changing data in Google spreadsheets using both C# and IronPython.

## Using library in your Applications

C# developers only need to
[download the library](https://www.nuget.org/packages/SynSys.GSpreadsheetEasyAccess)
using NuGet Package Manager.

IronPython developers need to:

- Download the dll archive at
[releases](https://github.com/Synergy-Systems-LLC/SynSys.GSpreadsheetEasyAccess/releases);
- Place all the dll files in the project folder;
- Load assemblies using one of the function of the clr module. Details can de seen in
[examples](https://github.com/Synergy-Systems-LLC/SynSys.GSpreadsheetEasyAccess/tree/master/examples/IronPythonApp).  

### IronPython Stubs

In order to conveniently use the library with IronPython it's better to utilize
[stubs](https://github.com/Synergy-Systems-LLC/SynSys.GSpreadsheetEasyAccess/tree/master/stubs).  
They can be downloaded in
[releases](https://github.com/Synergy-Systems-LLC/SynSys.GSpreadsheetEasyAccess/releases).  
If you don't know anything about stubs you can read about them
[here](https://github.com/BIMOpenGroup/RevitAPIStubs).

## How can I make a contribution?

All contributions are welcome! Send us your questions, suggestions and comments.

Like most open source projects this one uses the Forking Workflow system.

Summary:

- Make a repository fork;
- Make a branch from master;
- Make a pull request from upstream/master.

Branch naming rules:

- Kebab-case style;
- First word is a fix/feature/refactor or other task.  

Next ones are either a short description or an issue number.  
fix-iss57 / feature-iss14 / refactor-generator.

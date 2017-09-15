# Azure Functions bindings to O365
Prototype for some WebJobs extension bindings.

This provides a sample Azure Function extensions for Office. 

## This provides a few bindings:

- [Outlook] - sends emails from an O365 account
- [OneDrive] - reads/writes a Onedrive file 
- [Excel] - reads/writes an Excel table or worksheet. 
- [Token] - this has been extended to allow binding to the MS Graph SDK's Graph Service Client

It also includes Graph Webhook support. See https://github.com/microsoftgraph/aspnet-webhooks-rest-sample for graph webhooks sample. 

The Outlook, OneDrive, and Excel bindings derive from the token binding. See https://github.com/glennamanns/TokenBinding for more information about the token binding.

These bindings should also integrate with the O365 SDK to provide a full experienc.

See https://github.com/Azure/MikeBindings/blob/master/Samples/Functions.cs for sample usages. 

## For Authentication

Authentication is built using Easy Auth's token store. (see https://cgillum.tech/2016/03/07/app-service-token-store/ ) 
The app has an AAD app registered that has been configured iwth access to the Graph API and given appropriate scopes. The bindings can access the client secret (via appsettings) and use that to perform token exchanges.  

- OnBehalf - the identifiy flows in from the http request, and the binding does a token exchange
- Config-time - the app is provided a with a credential at design time (this is like a "saved" OnBehalf).
The refresh token is stored in the EasyAuth token store. 
- Webhook - this combines OnBehalf  and Config-Time. The app saves the refresh token when we subscribe to the webhook, and then looks it up in the EasyAuth Token store when 
when we receive the webhook. 

## Source layout 
The Samples project is a command line drive to invoke the bindings on a sample file (Function.cs) 
The OfficeBindings project contains the bindings. This is what Azure Functions would load. 



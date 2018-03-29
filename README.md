# Azure Functions bindings to O365
Prototype for some WebJobs extension bindings.

This provides a sample Azure Function extensions for Office. 

## This provides a few bindings:

- [Outlook] - sends emails from an O365 account
- [OneDrive] - reads/writes a Onedrive file 
- [Excel] - reads/writes an Excel table or worksheet. 
- [Token] - this has been extended to allow binding to the MS Graph SDK's Graph Service Client
- [GraphWebhookSubscription] - creates, deletes, or refreshes a Graph Webhook. See https://github.com/microsoftgraph/aspnet-webhooks-rest-sample for graph webhooks sample. 
- [GraphWebhookTrigger] - the trigger that activates a function when a Graph Webhook is called for its datatype

The bindings found in the Microsoft Graph extension use the same authentication process as those in the Token Extension. You can see how to use these bindings in the samples directory.

## For Authentication

Authentication is built using Easy Auth's token store. (see https://cgillum.tech/2016/03/07/app-service-token-store/ ) 
The app has an AAD app registered that has been configured with access to the Graph API and given appropriate scopes. The bindings can access the client secret (via appsettings) and use that to perform token exchanges.  

The bindings can authenticate in 4 different ways:
- UserFromId - the token is grabbed on behalf of a provided user id
- UserFromToken - the token is grabbed on behalf of a provided user id token
- UserFromRequest - the token is grabbed on behalf of the user id token found in the HTTP header `X-MS-TOKEN-AAD-ID-TOKEN`
- ClientCredentials - uses the app's credentials to access AAD

## Source layout 
The samples directory contains examples of how to use the bindings. The code for both the Token and Microsoft Graph extensions can be found in the src directory, and in-memory tests can be found in the tests directory.

## Local Development
First create a Functions app using the Functions CLI found https://docs.microsoft.com/en-us/azure/azure-functions/functions-run-local. Be sure to use the Version 2.x runtime.

To install the Token extension, run the command `func extensions install --package Microsoft.Azure.WebJobs.Extensions.AuthTokens -v <version>`.

To install the Microsoft Graph extension, run the command `func extensions install --package Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph -v <version>`.

The easiest way to utilize most of the features of these features is to configure an Azure Functions app in the Portal, enable Authentication/Authorization, add the extension, and go through the configuration to enable the proper Microsoft Graph permissions.

If you are making code changes to the extensions themselves and wish to test these locally, you can manually copy the .dll files that you build into your bin directory in your local function app's directory.

App Settings to Modify in `local.settings.json`:
- `WEBSITE_AUTH_CLIENT_ID` - Copy from your App Settings in Kudu from your configured app
- `WEBSITE_AUTH_CLIENT_SECRET` - Copy from your App Settings in Kudu from your configured app
- `BYOB_TokenMap` - A valid local directory that you have read/write access to

## License

This project is under the benevolent umbrella of the [.NET Foundation](http://www.dotnetfoundation.org/) and is licensed under [the MIT License](https://github.com/Azure/azure-webjobs-sdk/blob/master/LICENSE.txt)

## Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.


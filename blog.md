# Alpha Preview for Microsoft Office Bindings

## Personal Introduction

*My name is Glenna Manns and I am a software engineering intern on the Azure Functions team in Redmond, WA. As I write this, I am finishing up my last day on my twelfth and final week at Microsoft. My internship was a truly incredible experience. I learned a great deal not just about specific programming languages, but also about good engineering design practices. I got to work on both the nuts and bolts of my features as well as the UX and UI.* 

*I attend the University of Virginia in Charlottesville, Virginia. In May of 2018, I will graduate with a Bachelor of Science in Computer Science.*

## Introduction

Microsoft Office bindings extend the existing Microsoft Graph SDK and WebJobs framework to allow users easy access to Microsoft Graph data (emails, files*, events, etc.) from Azure Functions.

By using the Office input, output, and trigger bindings, users can bind to user-defined types like POCOs, primitives, or directly to Microsoft Graph SDK objects like WorkbookTables and Messages. These bindings handle (AAD) authentication, token exchange, and token refresh, allowing users to focus on writing code that utilizes MS Graph data.

**If accessing files, they must be stored in an O365 user's OneDrive.*

## Features

Currently, the bindings can be grouped into four main categories: **Excel, Outlook, OneDrive, and Webhooks**.

### Excel [Input + Output]

The Excel binding allows users to read/write Excel workbooks and tables using different data types (e.g. lists of POCOs and 2D arrays). 

### Outlook [Output]

The Outlook binding is an output only binding, and allows users to send emails from their O365 email accounts. 

### OneDrive [Input + Output]

The OneDrive binding allows users to read/write files stored in their OneDrive using several different data types (e.g. DriveItems and streams). 

### Webhooks [Trigger + Input]

There are three bindings associated with Graph webhooks: GraphWebhookTrigger, GraphWebhookCreator, and GraphWebhook. 

#### GraphWebhookTrigger [Trigger]
The GraphWebhookTrigger binding allows users to subscribe to notifications about a particular Microsoft Graph resource. 

At this time, Microsoft Graph does not support a large number of resources, so webhook is limited to: **email messages, OneDrive root, contacts, and events.** 

The notification payload itself from Microsoft Graph is not particularly useful; it only contains a subscription ID and the subscribed resource. In order to retrieve useful information, the webhook binding first maps that subscription ID to the principal ID of the user (stored by the Office extension at subscription-time), then performs a GET request for the specificed resource and transforms that payload into either a JSON object or a Microsoft Graph SDK type (e.g. Message, Contact).

#### GraphWebhookCreator [Input]
The GraphWebhookCreator input binding provides an alternate route independent of UX to create subscriptions. By providing a GraphWebhookCreator with an O365 user's credentials (see Authentication), function code can call SubscribeAsync(), which will create a new subscription to the specified resource. 

#### GraphWebhook [Input]
The GraphWebhook input binding handles subscription refresh and deletion. 

When *RefreshAllAsync* is called, the GraphWebhook binding looks up the subscription IDs associated with a given Function App and makes one call per subscription to Microsoft Graph requesting its renewal. Without renewal, most Microsoft Graph webhooks expire in [3 days](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/subscription). 

Deleting a subscription involves the removal of the stored mapping of a subscription ID to a principal ID, plus a request to Microsoft Graph to delete the subscription.

### Miscellaneous

In addition, a user can bind directly to the GraphServiceClient, which the binding authenticates behind-the-scenes. Use of this client exposes the full power of the Microsoft Graph SDK. 

## Example
An easily imagined scenario for these bindings is a business owner with customers who subscribe to their monthly newsletter. Customers provide their names and email addresses, which then must be added to an Excel file of customers. 

Using the Excel output binding, the business owner can select which Excel table to modify.
![ExcelInput](https://user-images.githubusercontent.com/10789958/28978083-8071b124-78f9-11e7-9117-595a7e9f52bc.png)

The function code to append the Excel row is only a few lines long. In this example, the function receives a POST request with a user's name and email address.
```csharp
using System.Net;

public static async Task Run(HttpRequestMessage req, TraceWriter log, IAsyncCollector<EmailRow> outputTable)
{
    // Get request body
    dynamic data = await req.Content.ReadAsAsync<object>();

    // Use body data to set row data
    var output = new EmailRow {
        Name = data?.name,
        Email = data?.email
    };
    await outputTable.AddAsync(output);
}

public class EmailRow {
    public string Name {get; set;}
    public string Email {get; set;}
}
```

Every month, a timer trigger fires and an email, the contents of which are determined by a OneDrive file, is sent out to each customer.

The business owner can select the same Customers Excel file (Excel input binding)..
![ExcelInput](https://user-images.githubusercontent.com/10789958/28978186-dd45da9c-78f9-11e7-85eb-d3176ebce5fc.png)

..determine which OneDrive file to get the email contents from (OneDrive input binding)..
![OneDriveInput](https://user-images.githubusercontent.com/10789958/28978207-f047bf2a-78f9-11e7-8054-3e762f1f70d7.png)

..and indicate that they would like to send emails via the Outlook output binding.
![Outlook](https://user-images.githubusercontent.com/10789958/28978233-07d4ee56-78fa-11e7-9ee8-0bc6df0a40d0.png)

The code below quickly scans the Excel table and sends out one email per row (customer).
```csharp
#r "Microsoft.Graph"

using System;
using Microsoft.Graph;

// Send one email per customer
public static void Run(TimerInfo myTimer, TraceWriter log, 
    List<EmailRow> inputTable, string file, ICollector<Message> emails)
{
	// Iterate over the rows of customers
    foreach(var row in inputTable) {
        var email = new Message {
            Subject = "Monthly newsletter",
            Body = new ItemBody {
                Content = file, //contents of email determined by OneDrive file
                ContentType = BodyType.Html
            },
            ToRecipients = new Recipient[] {
                new Recipient {
                    EmailAddress = new EmailAddress {
                        Address = row.Email,
                        Name = row.Name
                    }
                }
            }
        };
        emails.Add(email);
    }
}

public class EmailRow {
    public string Name { get; set; }
    public string Email { get; set; }
}
```
The aforementioned goals can be accomplished using just a few lines of code. **No manual data entry, no hardware, and no additional services.**

### C#

## Authentication

All actions mentioned are performed using an O365 user's identity. Which user's identity is used is up to the Function author. In order to authenticate against the Microsoft Graph, either an ID token or a Principal ID must be given to the binding. This identifier can come from a number of different sources. Examples of these sources include the *X-MS-TOKEN-AAD-ID-TOKEN* header of an HTTP request, the content of an HTTP request, a queue, or an app setting. Principal IDs can quickly be retrieved using the 'new' button seen in the screenshots above. 

## Language Support
Office bindings were designed with C# in mind. Having said that, there is limited JavaScript and Python support. The primary limitation of other languages is the lack of .NET objects like DriveItems and Messages.

## Future Work

- Improve setup experience
- Expand support for more languages
- Expand functionality as Microsoft Graph adds features (particularly webhooks)

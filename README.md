# Microsoft Build 2018 Demo: Using the DocuSign C# SDK to integrate with Microsoft Visual Studio, Azure, and SharePoint

At the Microsoft Build 2018 conference, we presented a comprehensive demo for using many technologies in the Microsoft stack, including Visual Studio 2017, Azure functions, and SharePoint Online. We showed how to use our C# SDK in a console app to send an envelope with a document created from a template for signature. We also showed how to setup an Azure function as a webhook listener to receive DocuSign Connect status messages. Finally, the demo showed how to upload completed (signed) documents into SharePoint.

This repository provides the Visual Studio Code and necessary files to setup and configure the demo we showed at Microsoft Build 2018. Detailed instructions can be found in [this blog post](https://www.docusign.com/blog/dsdev-msbuild2018-session-thr2605/)

## Files for you to update
Refer to to [this blog post](https://www.docusign.com/blog/dsdev-msbuild2018-session-thr2605/) for complete details, but you'll need to modify these files for your specific configuration information before you can run the demo:

 * PrivateKey.txt - contains the private key you will generate for OAuth2 authentication.
 * App.config - configuration file that contains values for your DocuSign sandbox account configuration, along with envelope and SharePoint configurations that are specific to your demo. The values for you to replace can be found by searching for "ReplaceMe" in the file.

## Files for you to copy content
Files that you will copy in their entirety are:
 * AzureFunctionCode.txt - contains the code you'll need in your Azure function to respond to DocuSign Connect messages.
 * AzureFunctionTest.txt - Sample, but stripped-down, XML response message that mimics what a DocuSign Connect message looks like. This enables you to test your Azure function even before you receive an actual Connect message.
 * AgreementSample.docx - Sample document for you to upload into your DocuSign sandbox account as a reusable template.
# Microsoft 365 and Graph Code Sample - Meeting Moderator

## Overview

This sample demonstrates usage of Microsoft Graph to implement small group breakouts in a large Microsoft Teams meeting. 

The sample is a web application built with React and Microsoft Graph Toolkit. The app can run as a web app in the browser, a personal or group tab app in Microsoft Teams, or a Progressive Web app on most modern desktop and mobile operating systems.

[Watch the live stream where we showed how the app was built](https://www.youtube.com/playlist?list=PLWZJrkeLOrbYL7tFQJ-HY6Q9FZGmqSldH)

## Prerequisites

* [NodeJS](https://nodejs.org/en/download/)
* A code editor - we recommend [VS Code](https://code.visualstudio.com/)
* Office 365 Tenant and a Global Admin account - [if you don't have one, get one for free here](https://docs.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-get-started)

## Run the sample

In a terminal:

```bash
git clone https://github.com/microsoftgraph/meeting-moderator-sample

cd meeting-moderator-sample

npm install

npm start
```

The app will launch in your browser at `http://localhost:3000`. Sign in with your  tenant credentials and consent to the permissions (make sure the user has ability to Admin consent).

The app will load the current signed in user's calendar with the option to moderate individual events.

> Note: Only Microsoft Teams online meetings will be available for moderation. All other events will be grayed out.

## Create an Azure Active Directory (AAD) application

The sample already contains a client id for an AAD application to allow you to get started immediately. However, we recommend creating your own so you have full control of the application. 

To create your own AAD app and client id, [follow the instructions in this blog post](https://developer.microsoft.com/microsoft-365/blogs/a-lap-around-microsoft-graph-toolkit-day-2-zero-to-hero/), under **Register your application**.

This step is required for the next section. The client id is defined in `src\index.tsx`

## Installing the sample as a Teams application

This sample is also a Microsoft Teams tab application and you can install it in your instance by following these instructions.

### 1. Run ngrok

To install it in your Teams environment, Microsoft Teams requires the app to be publicly accessible using HTTPS endpoint. To accomplish this, you can use `ngrok`, a tunneling software, which creates an externally addressable URL for a port you open locally on your machine:

1. Install ngrok
    ```bash
    npm install -g ngrok
    ```

1. Ensure the app is running locally on `http://localhost:3000`. Start ngrok and attach it to port 3000
    
    ```bash
    ngrok http 3000
    ```

    This will generate a public url (similar to `https://455709c1.ngrok.io`) that you can use to access the app. You will use this url in your Teams manifest.

### 2. Add ngrok url to your AAD application

Now that ngrok is running, you need to add this url to your AAD application redirect urls. If you have not created a new AAD application, make sure you do that first before continuing (see section above). 

### 3. Update Teams manifest and install application

To update the manifest, you will need to install the [Microsoft Team Extension](https://aka.ms/teams-toolkit) in VS code. Once you've installed the extension, open `.publish/Development.env` and replace the `baseUrl0` with your ngrok url. Save the file.

Once the file is saved, the `Development.zip` package will be automatically updated. You can now use this package to install the application in Microsoft Teams:

1. In Microsoft Teams, click on `Apps` in the lower left corner

2. Click on `Upload a custom app`. This will open the file picker. Select `Development.zip` to install the application.

## Using the Moderator Bot

This sample also includes a helpful bot for helping the moderator keep track of questions from the attendees. See the `bot` branch for description and instructions on deploying and running the bot in your Azure Subscription.


## Useful Links
- Microsoft Graph Dev Portal https://graph.developer.com/ 
- Graph Explorer https://aka.ms/GE 
- Microsoft Graph Toolkit https://aka.ms/mgt 
- Microsoft Graph Toolkit Blog Series https://aka.ms/mgtLap
- Microsoft Teams Dev Portal https://developer.microsoft.com/en-us/microsoft-teams

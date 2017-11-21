User: andres@arstconnazureo365.onmicrosoft.com
Password: Elite6060

User: developmentadmin@arstconnazureo365.onmicrosoft.com
Password: arst@dm1n


Sharepoint Site:
https://arstconnazureo365.sharepoint.com
https://arstconnazureo365.sharepoint.com/sites/DevOps
https://arstconnazureo365-my.sharepoint.com ==> for OneDrive
https://arstconnazureo365.sharepoint.com/SitePages/DevHome.aspx
https://arstconnazureo365-my.sharepoint.com/personal/andres_arstconnazureo365_onmicrosoft_com/_layouts/15/onedrive.aspx ==> for OneDrive
https://arstconnazureo365-my.sharepoint.com/personal/andres_arstconnazureo365_onmicrosoft_com ==> for OneDrive

DevOpsMSGraphAppPS
ID:
44d6a59c-0b68-434f-813f-d825da6669f3
Secret password:
VUYmoDfbior2SwTkejbQJAg

1Drive4Bussiness
ID:
990bfab9-6eb4-4cf3-b925-2abcf226dd04
Secret password:
QPFCiIw16XG8trhLoPmNbhQUAdrahRA/ifgDiKyGWP0=
URL:
https://sepago.de/1Drive4Business



Requirements:
Install-Module -Name OneDrive
Install-Module -Name PSMSgraph
Install-Module SharePointPnPPowerShellOnline


# Microsoft Graph API Connect Sample for Python

Connecting to Office 365 is the first step every app must take to start working
with Office 365 services and data. This sample shows how to connect and then
call one API through the Microsoft Graph API (previously called Office 365
unified API), and uses the Office Fabric UI to create an Office 365 experience.

<img src="./README assets/screenshot.PNG" alt="Python Connect sample screenshot" />

## Prerequisites

To use the Microsoft Graph API Connect sample for Python, you need the following:
* [Python 3.5.2](https://www.python.org/downloads/)
* [Flask-OAuthlib](https://github.com/lepture/flask-oauthlib)
* [Flask-Script 0.4](http://flask-script.readthedocs.io/en/latest/)
* [Requests module](http://docs.python-requests.org/en/latest/)
* A [Microsoft account](https://www.outlook.com/) or an [Office 365 for business account](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_Office365Account)

> Note: Microsoft has tested the [Flask-OAuthlib](https://github.com/lepture/flask-oauthlib) library in basic scenarios and confirmed that it works with the v2.0 endpoint. Microsoft does not provide fixes for this library and has not done a review of it. Issues and feature requests should be directed to the library’s open-source project.

## Register the application

Register an app on the Microsoft App Registration Portal. This generates the app ID and password that you'll use to configure the app for authentication.

1. Sign into the [Microsoft App Registration Portal](https://apps.dev.microsoft.com/) using either your personal or work or school account.

2. Choose **Add an app**.

3. Enter a name for the app, and choose **Create application**.

	The registration page displays, listing the properties of your app.

4. Copy the application ID. This is the unique identifier for your app.

5. Under **Application Secrets**, choose **Generate New Password**. Copy the app secret from the **New password generated** dialog.

	You'll use the application ID and app secret to configure the app.

6. Under **Platforms**, choose **Add platform** > **Web**.

7. Make sure the **Allow Implicit Flow** check box is selected, and enter *http://localhost:5000/login/authorized* as the Redirect URI.

	The **Allow Implicit Flow** option enables the OpenID Connect hybrid flow. During authentication, this enables the app to receive both sign-in info (the **id_token**) and artifacts (in this case, an authorization code) that the app uses to obtain an access token.

	The redirect URI *http://localhost:5000/login/authorized* is the value that the OmniAuth middleware is configured to use once it has processed the authentication request.

8. Choose **Save**.

## Configure and run the app

1. Using your favorite text editor, open the **_PRIVATE.txt** file.
2. Replace *ENTER_YOUR_CLIENT_ID* with the client ID of your registered application.
3. Replace *ENTER_YOUR_SECRET* with the key you generated for your app.
4. Start the development server by running ```python manage.py runserver```.
5. Navigate to ```http://localhost:5000/``` in your web browser.
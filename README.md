# TEAMSDUMP

This utility dumps all your Teams conversations into JSON files.

## Preparation

### Set up Azure App

Register an application in Azure Active Directory (Azure AD):

  1. Go to the Azure Portal.
  2. Register your application to obtain the Client ID, Tenant ID, and Client Secret.
  3. Configure a redirect URI (e.g., http://localhost:5000/callback).
  4. Add delegated API Permissions: _User.Read_, _Chat.Read_

### Python prerequisites

Install the necessary Python libraries: 

    * flask
    * requests
    * msal

In `dumpchats.py` file, update the required configuration:

    * CLIENT_ID = { CLIENT_ID }
    * CLIENT_SECRET = { CLIENT_SECRET }
    * TENANT_ID = { TENANT_ID}

## Running

Execute `python3 dumpchats.py`. Navigate to `http://localhost/5000`. Click "Login with Microsoft", follow the process... and wait.

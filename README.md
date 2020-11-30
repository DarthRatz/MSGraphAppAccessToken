# Microsoft Graph Application Permissions Access Token

This is a sample on how to use the Microsoft Graph JavaScript SDK with Application Permissions and your own Authentication Provider. For more information, see my blog post [here](https://www.lee-ford.co.uk/connect-graph-api-sdk-app-permissions/).

## Instructions

1. Create an Azure AD app with Graph Application Permissions. You will need the app ID, tenant ID and secret - for more info, see link to blog above

2. Clone this repository

3. Run `npm install` to install dependencies

4. Run `tsc --build` to compile to JavaScript in to _dist_ folder

5. Create an _.env_ file in the _dist_ folder and place the following inside (with your app details):

   **MicrosoftAppId**="_Azure AD App Id_"
   **MicrosoftAppPassword**="_Azure AD App Secret_"
   **MicrosoftAppTenantId**="_Azure AD Tenant Id_"

6. Run `npm start` to run the script
   !["example"](https://www.lee-ford.co.uk/images/graph-sdk-app-permissions/get-users.png)

# OutlookWelkinSyncFunction
These instructions will guide you through the steps needed to set up and run continuous synchronization between Welkin and Outlook using the Timer triggered Azure Function contained in this repository.

# Welkin setup
For this part you will need an admin account for your Welkin practice which you'll use to log in to the admin portal to perform the following steps.

## Obtain client ID and secret
From the Welkin admin portal, create a new client secret and ID. Details on where to find this configuration in the Welkin portal can be found [here](https://support.welkinhealth.com/hc/en-us/articles/360014076833-Connecting-your-Existing-Systems). Copy these down, you'll need them later.

## Create a dummy patient for placeholder events
In the Welkin admin portal, [create a new dummy patient](https://support.welkinhealth.com/hc/en-us/articles/360040203033-How-To-Create-a-new-Patient) and make note of the patient ID (you'll need this later as well). This dummy patient will be used in creating new Welkin events created as placeholders for events originating in Outlook. Don't use an existing Welkin patient associated with an actual clinical patient.

## Ensure needed Modality and Appointment type
In the Welkin admin portal, ensure that the modality type "call" exists, as well as the appointment type "intake_call". These will be the defaults for new Welkin events created as placeholders for events originating in Outlook.

# Outlook setup
For this part you'll need admin access to your O365 tenant. In other words, admin credentials with which you can log in to https://admin.microsoft.com/Adminportal and https://aad.portal.azure.com/.

## Create a client application and credentials
Logging in to https://aad.portal.azure.com/, navigate to App Registrations. Create a new registration and call it something like outlook-welkin-sync. Copy down the app ID. Now create a new client secret under Certificates and Secrets. Copy down the generated secret, which combined with the app ID you copied above will be the client ID/secret pair your sync process will use to access the MS Graph API. For more details, see the instructions [here](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app). Next you'll need to grant the app access to the Graph API with admin consent. Instructions on how to do that can be found [here](https://docs.microsoft.com/en-us/graph/auth-v2-service#2-configure-permissions-for-microsoft-graph). You will need to grant the following Microsoft Graph API permissions to this new application:
* Calendars.ReadWrite (Application)
* Calendars.ReadWrite.Shared (Delegated)
* User.Read.All (Application)
* Domain.Read.All (Application)

# Azure setup
For this part you'll need admin access to a subscription via https://portal.azure.com/. Note that this dashboard endpoint is different than what we used above for configuring the client application. This is because earlier we were configuring access to the O365 tenant. Now we're deploying the Azure Function which will do the actual synchronization. This requires an Azure subscription that can deploy Azure compute resources, which your O365 account will not have.

## Deploy the Azure Function
Open this git project in VSCode and make sure 
you have the Azure Functions [plugin](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azurefunctions) installed. Use the plugin to sign in to the subscription you're using for deployment. Follow the instructions given [here](https://docs.microsoft.com/en-us/azure/azure-functions/functions-develop-vs-code?tabs=csharp#publish-to-azure). Note that we'll be setting up a Timer triggered function rather than an HTTP triggered one.

## Configure the Azure Function.
Start [here](https://docs.microsoft.com/en-us/azure/azure-functions/functions-how-to-use-azure-function-app-settings) to learn about configuring Azure Function application settings. You'll be adding 7 or 9 new app settings, depending on whether you want to sync to a shared calendar or not. These new settings are:

* TimerSchedule: A crontab string, e.g. "* */15 * * * *", that will dictate how often the sync function runs (this example runs every 15 minutes).
* WelkinClientId: The client ID you created in Welkin
* WelkinClientSecret: The corresponding secret
* WelkinDummyPatientId: The GUID for the dummy patient you created in Welkin
* OutlookTenant: The tenant ID of your O365 tenant, which can be obtained through https://aad.portal.azure.com/
* OutlookClientId: The client ID you generated in Outlook setup above
* OutlookClientSecret: The corresponding client secret
* OutlookSharedCalendarUser: Configure this *only* if you're sync'ing to a shared calendar in Outlook. This is the full user name + domain of the owner of the shared calendar.
* OutlookSharedCalendarName: Configure this *only* if you're sync'ing to a shared calendar in Outlook. This is the name of the shared calendar.

With these configured values in place, you're ready to start the function and monitor its progress via Azure Insights log [streaming](https://docs.microsoft.com/en-us/azure/azure-functions/functions-monitoring#streaming-logs).
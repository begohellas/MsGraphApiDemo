# Microsoft Graph Api Examples
Sample application that demonstrates how to use Microsoft Graph API to access users, groups, and mail.

## Getting started
1. Register your application in Azure Active Directory (AAD) in app registrations [Tutorial](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app?tabs=certificate)
2. In section "Certificates & secrets" create a client secret, save the value first to close otherwise it's no longer available 		
3. In section "Api permissions" give it the necessary permissions to access the Microsoft Graph API with granted admin consent:
   - User.Read.All
   - GroupMember.Read.All

In application __clientsecret setting__ is stored in user secrets. 
To add user secrets to the project, right-click on the project in Solution Explorer and select Manage User Secrets or from line command _dotnet user-secrets_ . Add the following JSON to the secrets.json file:

```json
{
  "settings": {
	"clientSecret": "YOUR_CLIENT_SECRET"
  }
}
```

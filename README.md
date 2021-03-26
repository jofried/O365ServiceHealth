# O365 Service Health

<p>The O365 Service Health script can be used in conjucntion with an Azure AD app registration to pull Office 365 Health Data. The data is pulled from the Office 365 Service Communications API.<p/>
https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference
  
You will first need to create an Azure AD App registration
https://github.com/jofried/O365ServiceHealth/blob/master/O365ServiceHealthAppRegistration.pdf

Once the App Registration is complete you will need to modify the script (O365ServiceHealth.ps1) to add the the Client and Tenant ID, as well as the Client Secret.

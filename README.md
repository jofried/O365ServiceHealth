# O365 Service Health

<p>The O365 Service Health script can be used in conjunction with an Azure AD app registration to pull Office 365 Health Data. The data is pulled from the Office 365 Service Communications API.<p/>

Office 365 Service Communications API reference
<br>
https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference


<p><b>Please review the pdf below to start creating the Azure AD app registration:</b></p>
You will first need to create an Azure AD App registration
https://github.com/jofried/O365ServiceHealth/blob/master/O365ServiceHealthAppRegistration.pdf <br>
Once the App Registration is complete you will need to modify the script (O365ServiceHealth.ps1) to add the the Client and Tenant ID, as well as the Client Secret.<br>

The script will output the csv report to "C:\temp\O365ServiceHealth_datetime" <br>
Please create path "C:\temp\" if it doesnt already exist

# 2/10/23 by Jack Elmore
# Inspired by https://www.gitbit.org/course/ms-500/blog/how-to-send-emails-through-microsoft-365-from-powershell-injifle8u

Import-Module Microsoft.Graph

# REPLACE THESE GUID from Azure AD App Registration Portal
$ClientId = '12345678-1234-1234-1234-1234567890AB'
$TenantId = '12345678-ABCD-9876-5432-1234567890AB'
# REPLACE THIS THUMBPRINT FROM SELF-SIGNED CERTIFICATE DOCUMENTED ABOVE
$CertThumbprint = 'A364E8BFA598C7C912xxxxxxxxxxxxxxxxxxxx'

# Load the e-mail contents from an HTML file
$Content = Get-Content -Path "demo email.htm" -Raw

# Set Message Metadata
$Message = @{
    subject = "Demo HTML mail: Hello, World!";
    toRecipients = @(
        @{
            emailAddress = @{
                address = "recipient1@domain.com"
            }
        }
        @{
            emailAddress = @{
                address = "recipient2@domain.com"
            }
        }
    )
    body = @{
        contentType = "HTML";
        content = $Content
    }
}

Connect-MgGraph -ClientID $ClientId -TenantId $TenantId -CertificateThumbprint $CertThumbprint
# Takes a long time and seems unnecessary for testing since it's the default vs. Beta (Get-MgProfile)
# Select-MgProfile -Name 'v1.0'
Send-MgUserMail -UserId "m365sender@domain.onmicrosoft.com" -Message $Message
Disconnect-MgGraph
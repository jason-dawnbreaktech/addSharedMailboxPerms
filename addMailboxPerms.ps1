# Function to add permissions to a calendar in Exchange Online

# Prompt the user to authenticate with Exchange Online
Connect-ExchangeOnline

do {
    function Add-Permissions {
        # Prompt for the email address of the mailbox containing the calendar
        $MailboxEmail = Read-Host -Prompt "Enter the email address of the mailbox containing the calendar (e.g. user@example.com)"
        
        # Construct the identity of the calendar using the mailbox email address
        $CalendarIdentity = $MailboxEmail + ":\Calendar"
    
        # Prompt for the email address of the user to whom permissions will be granted
        $UserEmail = Read-Host -Prompt "Enter the email address of the user to whom permissions will be granted (e.g. user@example.com)"
        
        # Prompt for the access level to grant to the user
        $AccessLevel = Read-Host -Prompt "Enter the access level to grant to the user (e.g. Editor, Reviewer)"
        
        # Prompt to confirm the action before making any changes
        $Confirmation = Read-Host -Prompt "Are you sure you want to grant $AccessLevel access to $UserEmail for the calendar of $MailboxEmail? (Y/N)"
        
        # If the user confirms the action, continue
        if ($Confirmation.ToLower() -eq 'y') {
            # Prompt for whether to send an email notification to the user
            do {
                $SendNotification = Read-Host -Prompt "Do you want to send an email notification to $UserEmail? (Y/N)"
            } until ($SendNotification -in 'Y','y','N','n')
            # Convert the response to a boolean value
            $SendNotification = ($SendNotification.ToLower() -eq 'y')
            
            # Add the mailbox folder permission
            Add-MailboxFolderPermission -Identity $CalendarIdentity -User $UserEmail -AccessRights $AccessLevel -SendNotificationToUser $SendNotification
            
            # Display the permissions for the calendar
            Write-Host "Permissions for calendar {$CalendarIdentity}:"
            Get-MailboxFolderPermission -Identity $CalendarIdentity
        }
        else {
            # If the user did not confirm the action, display a message and exit the function
            Write-Host "Action cancelled by user."
            return
        }
    }
    
    # Call the Add-Permissions function
    Add-Permissions
    
    # Prompt the user whether to run the function again
    $RunAgain = Read-Host -Prompt "Do you want to grant calendar permissions again? (Y/N)"
} while ($RunAgain.ToLower() -eq 'y')

# The script ends when the user chooses not to run the function again
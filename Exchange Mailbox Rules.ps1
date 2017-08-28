#Connect to Exchange Online using administrator credentials
$UserCredential = Get-Credential 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
Clear-Host

#Basic description 
Write-Output "This script will check for inbox and sweep rules on Exchange Online"

#CheckMailbox function to look for mailbox rules
#Takes no parameters
#Returns no values
function CheckMailbox {
    $Mailbox = Read-Host -Prompt "Enter the the mailbox to check"

    Get-InboxRule -Mailbox $Mailbox | Select Name, Description | fl
    
    #Initializing array
    #This will be passed to another function
    $RulesArray = @()
    $RulesArray += Get-InboxRule -Mailbox $Mailbox | Select Name

    If ($RulesArray){
        $Question = Read-Host -Prompt "Do you want to delete any of these mailbox rules?"
        If ($Question -eq 'yes' -or $Question -eq 'Yes' -or $Question -eq 'y' -or $Question -eq 'Y') {
            DeleteMailboxRule $RulesArray $Mailbox
        }
    }

    $CheckAgain = Read-Host -Prompt "Do you want to check another mailbox?"

    If ($CheckAgain -eq 'yes' -or $CheckAgain -eq 'Yes' -or $CheckAgain -eq 'y' -or $CheckAgain -eq 'Y') {
        CheckMailbox
    }
    
}

#DeleteMailboxRule
#Takes the list of inbox rules and the mailbox address
#Prompts for the rule to delete, and again for confirmation
#Returns no value
function DeleteMailboxRule($RulesArray, $Mailbox) {
    $i = 0;
    Write-Host "Mailbox =" $Mailbox
    foreach ($Rule in $RulesArray){
        Write-Host $i " = " $Rule.Name #| FL
        $i++
    }
    $RuleToDelete = Read-Host -Prompt "Which rule would you like to delete? (Number)"
    $Confirmation = Read-Host -Prompt "You selected $($RulesArray[$RuleToDelete].Name) -- Are you sure you want to remove this rule?"
    if ($Confirmation -eq 'yes' -or $Confirmation -eq 'y' -or $Confirmation -eq 'Yes' -or $Confirmation -eq 'Y'){
        Remove-InboxRule -Mailbox $Mailbox -Identity $RulesArray[$RuleToDelete].Name
    }
    Write-Host "The rule has been removed"
}
CheckMailbox
Remove-PSSession -Session $Session
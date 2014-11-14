#This is an Exchange Management Shell (EMS) script.  It must be run using EMS and not a regular PowerShell Window.

CreateEmailListCSV("All Employees")
CreateEmailListCSV("All Executives")
CreateEmailListCSV("Sales")

Function CreateEmailListCSV($dist_group)
{
    $output_file = $dist_group + " Email List.csv"

    #Get name and email address of members in the specified distribution group
    $all_accounts = Get-DistributionGroupMember -Identity $dist_group | Select Firstname, Lastname, DisplayName, PrimarySMTPAddress

    #For each account in the set of accounts, add user accounts and not service accounts to a new variable for exporting.
    $user_accounts = @()
    foreach ($account in $all_accounts)
    {
        #If name fields are not empty, then it is a user account.
        if (![string]::IsNullOrWhiteSpace($account.Firstname) -and ![string]::IsNullOrWhiteSpace($account.Lastname) -and ![string]::IsNullOrWhiteSpace($account.DisplayName) -and ![string]::IsNullOrWhiteSpace($account.PrimarySMTPAddress))
        {
            $user_accounts += $account
        }
    }

    $user_accounts.Count

    #Sort by firstname followed by lastname and then export user accounts to the specified CSV file
    $user_accounts.GetEnumerator() | Sort-Object Firsname | Sort-Object Lastname | Export-Csv -NoTypeInformation -Path $output_file
}

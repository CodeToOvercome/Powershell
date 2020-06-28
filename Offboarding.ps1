#Prerequisite 1> This part looks for Installed Modules on running computer and installs modules that is necessary to get this script working 

    $isModuleInstalled = Get-InstalledModule |Where-Object{$_.Name -eq "MsOnline"}

    if($isModuleInstalled.Name -ne "MSOnline")

        {
            Write-Host "Installing Necessary Modules, Please Wait... " -ForegroundColor Black -BackgroundColor Yellow
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
            Install-Module -Name MSOnline -Confirm:$false
            
        }
    else 
        {
            Write-Host "Necessary Modules are already installed" -ForegroundColor Black -BackgroundColor Yellow
        }

#Prerequisite 2> Getting Admin Credentials and saving them & Connecting to Exchange Online Module

    do{
        $isItBadCredential = $false
        try 
            {
                if($Session.State -ne "Opened")
    
                    {
                        $adminCredential = Get-Credential -Message "Admin Credential Needed to manage Exchange Online & Office365 License" 

                        Write-Host "Installing Exchange Modules...Please Wait" -ForegroundColor Black -BackgroundColor Yellow

                        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $adminCredential -Authentication Basic -AllowRedirection -ErrorAction Stop

                        Import-PSSession $Session -DisableNameChecking  
                    }Else
                    {
                        Write-Host "Using Existing Exchange Session" -ForegroundColor Black -BackgroundColor Yellow
                    }
                
                    # The do While loop below will make sure that correct username is inputted .
                    
                    $startTime = (Get-Date) #This is used to calculate the time it takes to deboard a user
                    do{
                        $isUserNameInCorrect = $false
                    
                        Try
                            {
                                        
                                #Step 1 > Disabling Account
                
                                $user = Read-Host "What's the username that you would like to deboard today?" 
                                
                                $userUPN = Get-ADUser -Identity $user |Select-Object UserPrincipalName
                        
                                Set-ADUser -Identity $user -Enabled $false -ErrorAction Stop
                
                                Write-Host $user "account is now disabled." -ForegroundColor Black -BackgroundColor Yellow
                
                                #Step 2 > Renaming Account 
                
                                $userGUID = Get-ADUser -Identity $user -Properties objectGUID |Select-Object ObjectGUID 
                
                                $userGUIDValue = $userGUID.ObjectGUID.Guid 
                
                                $userDisplayName = Get-ADUser -Identity $user -Properties DisplayName|Select-Object DisplayName 
                
                                $sleeperValue = "Z - "
                
                                $userNewName = $sleeperValue + $userDisplayName.DisplayName
                
                                If($userDisplayName.DisplayName -notmatch $sleeperValue)
                                    {
                
                                        Rename-ADObject -Identity $userGUIDValue -NewName $userNewName
                
                                        Set-ADUser -Identity $user -DisplayName $userNewName
                
                                        $result = Get-ADUser -Identity $user -Properties DisplayName|Select-Object DisplayName,DistinguishedName
                
                                        Write-Host "The New Name of the user is" $result.DisplayName -ForegroundColor Black -BackgroundColor Yellow
                
                                    }
                                else
                                    {
                                        Write-Warning "Looks like this user account is already renamed"
                                    }
                            
                                #Step 3 > Removing from Distribution lists
                
                                $groupsToRemoveFrom = Get-ADPrincipalGroupMembership -Identity $user |Select-Object SamAccountName,GroupCategory|Where-Object{$_.GroupCategory -eq "Distribution"}| Select-Object SamAccountName
                            
                                if($groupsToRemoveFrom.count -ne 0)
                                    {
                                        Remove-ADPrincipalGroupMembership -Identity $user -MemberOf $groupsToRemoveFrom -Confirm:$false
                
                                        Write-Host $user "is removed from the following distribution groups" -ForegroundColor Black -BackgroundColor Yellow
                                            
                                        foreach ($group in $groupsToRemoveFrom)
                                        {
                                            Write-Host $group.SamAccountName -ForegroundColor Black -BackgroundColor Yellow
                                        }
                                    }
                                else
                                    {
                                        Write-Warning "User is not part of any Distribution Groups"
                                    }
                                
                                #Step 4> Moving to Dismissals Container
                
                                $dismissalsConatiner = "OU=Dismissals,OU=Shared Mailboxes,OU=Domain,DC=sub-domain,DC=internal" 
                
                                Move-ADObject -Identity $userGUIDValue -TargetPath $dismissalsConatiner
                
                                Write-Host $user is moved to $dismissalsConatiner -ForegroundColor Black -BackgroundColor Yellow
                
                                #Step 5> Converting mailbox into a shared type
                
                                try
                                    {
                                    Set-Mailbox -Identity $user -Type Shared -WarningAction Stop
                                    
                                    Write-Host $user "mailbox is now converted into Shared Mailbox" -ForegroundColor Black -BackgroundColor Yellow 
                                    }
                                catch 
                                    {
                                        Write-Host "Proceeding to the next step..." -ForegroundColor Black -BackgroundColor Yellow
                                    }
                
                                #Step 6> Mail Forwarding (Optional)
                
                                $title1   = "Mail Forwarding"
                            
                                $message1 = "Do you want to enable mail forwarding?"
                                
                                $yes1 = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Do you wanna enable mail forwarding ?"
                                
                                $no1 = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Skips to next step."
                
                                $options1 = [System.Management.Automation.Host.ChoiceDescription[]]($yes1, $no1)
                                
                                $result1 = $host.ui.PromptForChoice($title1, $message1, $options1, 0)
                                
                                switch ($result1) 
                                    {
                                        0
                                            { 
                                                do
                                                {
                                                    $invalidAddress =$false
                                                
                                                    try
                                                        {
                                                                $receipient1 = Read-Host "To whom are you forwading the email to ?. Please Enter full Email Address "
                
                                                                Set-Mailbox $user -ForwardingAddress $receipient1 -ErrorAction Stop
                
                                                                Write-Host "Effective immediately" $user " email is being forwarded to " $receipient1 -ForegroundColor Black -BackgroundColor Yellow
                                                        }
                                                    catch
                                                        {
                                                                $invalidAddress = $True
                                                                Write-Warning "The address you entered is Invalid. Please check your Spelling and Type again" 
                                                        }
                                                }while ($invalidAddress)
                                            }
                                            
                                
                                        1
                                            {
                                                Write-Host "Skipping to the next step" -ForegroundColor Black -BackgroundColor Yellow
                                            }
                                    }
                                
                                #Step 7> Granting Mailbox Permission (Optional)
                
                                $title2   = "Access to Mailbox" 
                                
                                $message2 = "Do you want to grant someone access to this user's mailbox?"
                                
                                $yes2 = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","This will give full access to specfied user's mailbox"
                                
                                $no2 = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Skips to next Step."
                
                                $options2 = [System.Management.Automation.Host.ChoiceDescription[]]($yes2, $no2)
                                
                                $result2 = $host.ui.PromptForChoice($title2, $message2, $options2, 0)
                                
                                switch ($result2) 
                                    {
                                        0
                                            {
                                                do
                                                {
                                                    $invalidReceipient = $false
                                                    
                                                        try
                                                            {
                                                                $receipient2 = Read-Host "Who do you like to give access to this mailbox ?"
                
                                                                Add-MailboxPermission -Identity $user -User $receipient2 -AccessRights FullAccess -ErrorAction Stop -WarningAction SilentlyContinue
                
                                                                Write-Host $receipient2 "has been granted Full Access to " $user "mailbox." -ForegroundColor Black -BackgroundColor Yellow
                                                            }
                                                        catch 
                                                            {
                                                                $invalidReceipient =$true
                                                                Write-Warning "Check your spelling of the username.Accepted Format is firstname.lastname or firstname.lastname@domain.com"
                                                            }
                                                }while ($invalidReceipient)  
                                                
                                            }
                                
                                        1
                                            {
                                                    Write-Host "Skipping to the next step" -ForegroundColor Black -BackgroundColor Yellow
                                            }
                                    }
                                
                                #Step 8> Removing License from Office 365
                
                                Connect-MsolService -Credential $adminCredential
                
                                $userLicenses = Get-MsolUser -UserPrincipalName $userUPN.UserPrincipalName | Select-Object Licenses
                
                                    
                                    if($userLicenses.Licenses.Count -ne 0)
                                        {
                                            Set-MsolUserLicense -UserPrincipalName $userUPN.UserPrincipalName -RemoveLicenses $userLicenses.Licenses.AccountSkuID  
                
                                            Write-Host "The following licenses have been removed from " $user ": " $userLicenses.Licenses.AccountSkuID -ForegroundColor Black -BackgroundColor Yellow
                
                                            Write-Host $user "is now successfully deboarded. o(〃＾▽＾〃)o" -ForegroundColor Black -BackgroundColor Yellow
                
                                        }
                                    else
                                        {
                                            Write-Host "User currently do not have any license assigned"
                            
                                        }
                                
                                
                            }
                
                        Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
                            {
                                $isUserNameInCorrect = $True
                                Write-Warning -Message "Check your Spelling and type the proper username again"
                            
                            }
                        
                        catch 
                            {
                                Write-Error -Message "An Error Occured"
                                Write-Warning $_.Exception
                            }
                        
                        
                    }while($isUserNameInCorrect)
                    
                    # This block below is used to write event log to calculate time.
                    $endTime =(Get-Date)
                
                    $totalTime = "Elapsed Time: $(($endTime-$startTime).totalseconds) seconds"
                
                    $eventLogSource = Get-EventLog -LogName Application -Source "Deboarding Script"
                               
                    if ($eventLogSource.count -eq 0)
                    {
                        New-EventLog -LogName Application -Source "Deboarding Script"
                        Write-EventLog -LogName Application -Source "Deboarding Script" -EntryType Information -EventId 1 -Message $totalTime
                
                    }
                    else
                    {
                        Write-EventLog -LogName Application -Source "Deboarding Script" -EntryType Information -EventId 1 -Message $totalTime
                
                    }
                    
            }
        catch [System.Management.Automation.Remoting.PSRemotingTransportException] 
            {
                $isItBadCredential =$true
                Write-Warning "Incorrect-Credentials, Please input your Username & Password again"
            }
        catch [System.Management.Automation.ParameterBindingException]
            {
                Write-Warning "Need Credentials to proceed further, exiting script."
            }
    }while($isItBadCredential)



        
      
        
              
        

        
        

    
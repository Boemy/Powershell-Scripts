<# 
Steps to take: 
X - Confirm script is running in Admin Mode
   X - Use a function to check this or some other way
X - Check if module is installed and if not do it (use the checkmodule function from past scripts)
   X  - Module name: AzureAD
X - Connect to AzureAD
    X - Error Checking, if wrong let user input credentails again
X - Confirm connection with the AzureAD
X - Define path to CSV file
X - Read CSV file
   X - Error Checking, if worng let user input path again
X - Check if file has the needed data
   X  - Firstname, Surname, Lastname, Location, Function, Department Group & Password (these will be made with the Passwordgenerator app and follow the passwordpolicies)
X - Loop through al the groups and create them based on the given data
X - Confirm group creation
    X - Check if all groups are made before continuing
X - Loop through al the users and create them based on the given data
    X - Add the data to variables to combine data where needed
    X - UPN might consist of numbers as defined in the namingconvention 
        X - Add a 3 number code after the first part of the upn: j.doe.342@domain.com
    X - Add user to there group (use the security group column for this)
X - Confirm user creation 
   X  - Get UPN for each user and if not retrieved add to list of failed creations

X - !! Add tenant directory check !! 
#>

################################################################################### Functions ###################################################################################
# Clear the screen
Clear-Host

# Set ExecutionPolicy at the Process level to "Bypass"
Set-ExecutionPolicy -ExecutionPolicy "Bypass" -Scope "Process" -Force

# Function for self elevate if script is not being run as an Administrator
Function AdminTest {
    # Check if the script is running with Administrator privileges
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "This script is not being run as an Administrator. It will self-elevate to run as an Administrator."
        try {
            # Start a new PowerShell process with elevated privileges
            Start-Process -FilePath PowerShell.exe -ArgumentList -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Verb RunAs -ErrorAction Stop

            # Exit the current script
            Exit
        }
        catch {
            # Error handling if self-elevation fails
            Write-Host "Failed to self-elevate to run as an Administrator. Error: $($_.Exception.Message)"
            return
        }
    }
}

# Function to check if the Azure AD module is installed (no user input should be required unless an error occurs)
function CheckModule {    
    # Define the module that needs to be installed
    $ModuleName = "AzureAD"

    # Check if the defined module is installed
    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        # Inform the user that the defined module is installing
        Write-Host "Installing $ModuleName module..."

        # Attenmpt to install the defined module
        try {
            # Install the NuGet Package Provider and hide the ouput
            Install-PackageProvider -Name "NuGet" -Force -Scope "CurrentUser" -ErrorAction Stop | Out-Null

            # Install the defined module and hide the ouput
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope "CurrentUser" -ErrorAction Stop | Out-Null
            
            # Inform the user that the defined module has been installed
            Write-Host "Module $ModuleName installed successfully."
        }
        catch {
            # Error Handling
            Write-Host "Failed to install the $ModuleName module. Error: $($_.Exception.Message)"
            return
        }
    
        # Inform the user that the defined module is being imported
        Write-Host "Importing $ModuleName module..."

        try {
            # Import the defined module and hide the ouput
            Import-Module -Name $ModuleName -Force -ErrorAction Stop | Out-Null

            # Inform the user that the defined module has been imported
            Write-Host "Module $ModuleName imported successfully."
        }
        catch {
            # Error Handling
            Write-Host "Failed to import the $ModuleName module. Error: $($_.Exception.Message)"
        }
    }
}

# Run the function called "AdminTest"
AdminTest

# Run the function called "CheckModule"
CheckModule 

################################################################################### Variables ###################################################################################
# Clear the screen 
Clear-Host

# Ask the user for the full path of the CSV file (delimited by ';') for the user & group import
$CsvFilePath = Read-Host "Input the full path to the CSV file (delimited by ';', UTF8) for the user & group import"

# Ask the user to input the Tenant ID
$TenantID = Read-Host "Enter the Azure AD Tenant ID (You can find it in the Azure portal under Azure AD Properties)" 

# Ask the user for the tenant domain name
$TenantDomain = Read-Host "Enter the Azure AD Tenant Domain (e.g. 'twsr4.onmicrosoft.com')"

################################################################################### Script ###################################################################################

# Clear the screen 
Clear-Host

# Variable that will be used to define the connection of Azure AD
$AzureAD_Connected = $false

# Run the following loop until the variable "AzureAD_Connected" is not equal to "false"
while (!$AzureAD_Connected) { 
    # Attempt to connect to Azure AD Tenant
    try {
        # Clear the screen 
        Clear-Host

        # Ask the user to input the email and password for an Azure administrator account with create & edit rights
        Write-Host "Enter the email and password for an Azure administrator account with create & edit rights in the required tenant"

        # Connect to the specified Azure AD Tenant
        Connect-AzureAD -TenantId $TenantID -ErrorAction SilentlyContinue 
        
        # If connection was successful, set the variable to "true"
        $AzureAD_Connected = $true
    }
    catch {
        # Clear the screen
        Clear-Host

        # Ask the user to check their credentials and try again
        Write-Host "! Error: Failed to connect. Please check your credentials and try again. Error: $($_.Exception.Message)"

        # Ask the user to input the Tenant ID
        $TenantID = Read-Host "Enter the Azure AD Tenant ID (You can find it in the Azure portal under Azure AD Properties)" 
    }
}

# Attempt to validate the provided file path & file
try {
    # Check if the provided file path exists and is a valid CSV file
    if (!(Test-Path $CsvFilePath) -or (Get-Item $CsvFilePath).Extension -ne ".csv") {
        # Inform the user that the file path is not valid or doesn not point to a CSV file
        Write-Host "The provided file path is either invalid or does not point to a CSV file."
        
        # Ask the user to input the full path to a CSV file again
        $CsvFilePath = Read-Host "Please input a valid full path to a CSV file (delimited by ';') for the user & group import"
    } 
    
    # Read the CSV file and save the data in the variable "CsvData"
    $CsvData = Import-Csv -Path $CsvFilePath -Delimiter ";" -Encoding "UTF8"
} catch {
    # Error handling
    Write-Host "An error occurred while validating the file path. Error: $($_.Exception.Message)"
}

# Attempt to check the required columns 
try {
    # Define the "required columns"
    $RequiredColumns = @(
        "Firstname", 
        "Middlename", 
        "Surname", 
        "Location", 
        "Department", 
        "Function",
        "OfficeLocation" ,
        "SecurityGroup", 
        "Password"
    )

    # Check if the required columns are in the CSV file and if not add them to variable "MissingColumns"
    $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $CsvData[0].PSObject.Properties.Name }
        
    if ($MissingColumns.Count -gt 0) {
        throw "The CSV file is missing the following required column(s): " + ($MissingColumns -join ", ")
    }
} catch {
    # Error handling
    Write-Host "An error occurred while validating the required columns. Error: $($_.Exception.Message)"
}

# Create an array to save the values of the already created groups (extra check to avoid duplicate groups)
$CreatedGroups = @()

# Run through each row of data in the variable "CsvData"
foreach ($Row in $CsvData) {
    # Link the csv data to variables          
    $FirstName = $Row.Firstname 
    $MiddleName = $Row.Middlename  
    $Surname = $Row.Surname 
    $DisplayName = "$FirstName $(if($MiddleName) { "$MiddleName" }) $Surname"
    $Location = $Row.Location
    $Department = $Row.Department
    $Function = $Row.Function
    $OfficeLocation = $Row.OfficeLocation
    $SecurityGroups = $Row.SecurityGroup -split "`n" | ForEach-Object { $_.Trim() }
    $Password = $Row.Password
    $Domain = $TenantDomain
    $UserPrincipalName = (($FirstName.Substring(0, 1) + "." + $MiddleName + $Surname) -replace '\s', '') + "." + (Get-Random -Minimum 100 -Maximum 1000) + "@" + $Domain
    
    # Check if the user already exists in Azure AD
    if (!(Get-AzureADUser -Filter "DisplayName eq '$DisplayName'")) {
        try {
            # Define the parameters for "New-AzureADUser" cmdlet
            $NewUserParams = @{
                DisplayName                 = $DisplayName
                GivenName                   = $FirstName
                Surname                     = $Surname
                UserPrincipalName           = ($UserPrincipalName).ToLower()
                MailNickName                = ($FirstName.Substring(0, 1) + $MiddleName + $Surname).ToLower() -replace '\s', ''
                PasswordProfile             = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile | ForEach-Object { $_.Password = $Password; $_ }  
                AccountEnabled              = $true
                JobTitle                    = "$Function, $Location"
                Department                  = $Department
                Country                     = $Location
                PhysicalDeliveryOfficeName  = $OfficeLocation
                UserType                    = "Member"
            }

            # Create a new Azure AD user with the given values and parameters
            New-AzureADUser @NewUserParams | Out-Null

            # Inform the user that the user has been created
            Write-Host "Successfully created Azure AD user: $DisplayName"

            # Check if the Security Groups need to be created and the user needs to be added to them
            foreach ($SecurityGroup in $SecurityGroups) {
                # Check if the Security Group already exists
                $ExistingGroup = Get-AzureADGroup -Filter "DisplayName eq '$SecurityGroup'"
                
                # If the Security Group does not exisit and is also not part of the $CreatedGroups array
                if (!$ExistingGroup -and ($SecurityGroup -notin $CreatedGroups)) {
                    try {
                        # Define the parameters for "New-AzureADGroup" cmdlet
                        $NewGroupParams = @{
                            DisplayName      = $SecurityGroup
                            MailNickname     = [Guid]::NewGuid().ToString()
                            SecurityEnabled  = $true
                            MailEnabled      = $false 
                        }

                        # Create a new Azure AD group with the given values and parameters
                        New-AzureADGroup @NewGroupParams | Out-Null

                        # Add the group name to the array of created groups
                        $CreatedGroups += $SecurityGroup

                        # Inform the user that the group has been created
                        Write-Host "Successfully created Azure AD group: $SecurityGroup"
                    }
                    catch {
                        # Error handling
                        Write-Host "Failed to create Azure AD group: $SecurityGroup. Error: $($_.Exception.Message)"
                    }
                }
                elseif ($ExistingGroup) {
                    # Add the group name to the array of created groups
                    $CreatedGroups += $SecurityGroup
                }

                # Check if the user needs to be added to the group
                if ($SecurityGroup -in $CreatedGroups) {
                    try {
                        # Define the parameters for "Add-AzureADGroupMember" cmdlet
                        $AddGroupMemberParams = @{
                            ObjectId    = (Get-AzureADGroup -Filter "DisplayName eq '$SecurityGroup'").ObjectId
                            RefObjectId = (Get-AzureADUser -Filter "DisplayName eq '$DisplayName'").ObjectId
                        }

                        # Add the user to the group
                        Add-AzureADGroupMember @AddGroupMemberParams | Out-Null

                        # Inform the user that the user has been added to the group
                        Write-Host "Successfully added user $DisplayName to group: $SecurityGroup"
                    }
                    catch {
                        # Error handling
                        Write-Host "Failed to add user $DisplayName to group: $SecurityGroup. Error: $($_.Exception.Message)"
                    }
                }
            }
        } catch {
            # Error handling
            Write-Host "Failed to create user: $DisplayName. Error: $($_.Exception.Message)"
        }
    }
}

#Close the connection to Azure AD
Disconnect-AzureAD
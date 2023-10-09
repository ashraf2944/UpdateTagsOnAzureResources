# Install required modules
Install-Module -Name Az -AllowClobber -Force -Confirm:$false
Install-Module -Name ImportExcel -AllowClobber -Force -Confirm:$false

 

# Load Azure PowerShell module
Import-Module Az
Import-Module ImportExcel

 

# Function to update tags for resource groups and resources
Function Update-Tags {
    param(
        [string]$subscriptionId,
        [string]$resourceGroup,
        [hashtable]$tags
    )

    # Set the current subscription
    Set-AzContext -Subscription $subscriptionId

    # Update tags for the resource group
    Set-AzResourceGroup -Name $resourceGroup -Tag $tags

    # Get resources in the resource group
    $resources = Get-AzResource -ResourceGroupName $resourceGroup

    # Update tags for each resource
    foreach ($resource in $resources) {
        Set-AzResource -ResourceId $resource.ResourceId -Tag $tags
    }
}

 

# Connect to Azure using Service Principal
$clientId = "YOUR_SPN_CLIENT_ID"
$clientSecret = "YOUR_SPN_CLIENT_SECRET"
$tenantId = "YOUR_TENANT_ID"

 

# Authenticate using Service Principal
$securePassword = ConvertTo-SecureString -AsPlainText $clientSecret -Force
$psCredential = New-Object System.Management.Automation.PSCredential($clientId, $securePassword)
Connect-AzAccount -ServicePrincipal -Credential $psCredential -TenantId $tenantId

 

# Read Excel data and update tags
$excelFilePath = "C:\Path\To\Your\Excel\File.xlsx"
$excelData = Import-Excel -Path $excelFilePath

 

foreach ($row in $excelData) {
    $subscriptionName = $row.SUBSCRIPTION
    $resourceGroup = $row.RESOURCE_GROUPS
    $tags = @{
        "AppName" = $row.AppName
        "AppCategory" = $row.AppCategory
        "AppOwner" = $row.AppOwner
        "ITSponsor" = $row.ITSponsor
        "CostCenter" = $row.CostCenter
        "AssetOwner" = $row.AssetOwner
        "ServiceNowBA" = $row.ServiceNowBA
        "ServiceNowAS" = $row.ServiceNowAS
        "SecurityReviewID" = $row.SecurityReviewID
    }

    # Get the subscription ID based on subscription name
    $subscriptionId = (Get-AzSubscription | Where-Object { $_.Name -eq $subscriptionName }).Id

    # Update tags for the resource group and resources
    Update-Tags -subscriptionId $subscriptionId -resourceGroup $resourceGroup -tags $tags
}
function Get-AccessToken {
    param(
        [string]$TenantID,
        [string]$ClientID,
        [string]$ClientSecret
    )
    $TokenURL = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"


    $body = @{
        grant_type = 'client_credentials'
        client_id = $ClientID
        client_secret = $ClientSecret
        scope = 'https://graph.microsoft.com/.default'
    }

    $bodyString = ""
    foreach ($param in $body.GetEnumerator()) {
        $bodyString += "&$($param.Name)=$($param.Value)"
    }

    $response = Invoke-RestMethod -Uri $TokenURL -Method Post -Body $bodyString -ContentType "application/x-www-form-urlencoded" 

    $response.access_token
}

Export-ModuleMember -Function Get-AccessToken
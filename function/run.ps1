using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$ClientId = $env:ClientId
$ClientSecret = $env:ClientSecret
$TranslatorKey = $env:TranslatorKey
$TranslatorRegion = $env:TranslatorRegion
$Tenant = $env:Tenant
 
Write-Host "SiteURL is $($Request.Body.siteURL)"

# Interact with body of the request
$SiteURL = $Request.Body.siteURL
$TargetLanguage = $Request.Body.language
$PageTitle = $Request.Body.pageTitle

$status = 0

# Translate function
function Start-Translation {
	param(
		[Parameter(Mandatory = $true)]
		[string]$text,
		[Parameter(Mandatory = $true)]
		[string]$language
	)
 
	$baseUri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0"
 
	$headers = @{
		'Ocp-Apim-Subscription-Key'    = $TranslatorKey
		'Ocp-Apim-Subscription-Region' = $TranslatorRegion
		'Content-type'                 = 'application/json'
	}
 
	# Create JSON array with 1 object for request body
	$textJson = @{
		"Text" = $text
	} | ConvertTo-Json
 
	$body = "[$textJson]"
 
	# Uri for the request includes language code and text type, which is always html for SharePoint text web parts
	$uri = "$baseUri&to=$language&textType=html"
 
	Write-Host "Calling translator"

	# Send request for translation and extract translated text
	$results = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body
	$translatedText = $results[0].translations[0].text
	return $translatedText
}
 
#---START SCRIPT---#
try {
	$status = 1
	Connect-PnPOnline -Url $SiteURL -ClientId $ClientId -Tenant $Tenant -CertificateBase64Encoded $ClientSecret

	$status = 2
	$newPage = Get-PnPClientSidePage "$TargetLanguage/$PageTitle"
	$textControls = $newPage.Controls | Where-Object { $_.Type.Name -eq "ClientSideText" -or $_.Type.Name -eq "PageText" }
 
	$status = 3
	Write-Host "Translating content of $($textControls.length) controls" -NoNewline

	foreach ($textControl in $textControls) {
		$translatedControlText = Start-Translation -text $textControl.Text -language $TargetLanguage
		Set-PnPClientSideText -Page $newPage -InstanceId $textControl.InstanceId -Text $translatedControlText
	}

	$status = 4
	$title = $newPage.PageTitle -replace ("^Translate into[\w\W]*(?=: ): ", "")
	$translatedTitle = Start-Translation -text $title -language $TargetLanguage
	Set-PnPClientSidePage "$TargetLanguage/$PageTitle" -Title $translatedTitle -Publish

	Write-Host "Done!" -ForegroundColor Green
 
	Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
			StatusCode = [HttpStatusCode]::OK
			Body       = "{'message':'Page $PageTitle has been translated to $TargetLanguage'}"
		})
}
catch {
	$message = ""

	switch ($status) {
		1 {
			$message = "Cannot connect to SharePoint"
		}
		2 {
			$message = "Cannot get SharPopint page"
		}
		3 {
			$message = "Error translating content"
		}
		4 {
			$message = "Error translating title"
		}
	}

	Write-Host "@($message): @($_)"

	Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
			StatusCode = [HttpStatusCode]::BadRequest
			Body       = "{'message':'$message'}"
		})
}

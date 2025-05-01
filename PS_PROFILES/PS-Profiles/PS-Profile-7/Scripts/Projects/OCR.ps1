Function Get-ImageText()
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$True)]
		[String] $Path
)

Process{
            $SplatInput = @{
            Uri= "https://api.projectoxford.ai/vision/v1/ocr"
            Method = 'Post'
			InFile = $Path
			ContentType = 'application/octet-stream'
			}

            $Headers =  @{
			'Ocp-Apim-Subscription-Key' = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
			}

            Try{
            			# Call OCR API and feed the parameters to it.
				$Data = (Invoke-RestMethod @SplatInput -Headers $Headers -ErrorVariable +E)
				$Language = $Data.Language # Detected language
				$i=0; foreach($D in $Data.regions.lines){
				$i=$i+1;$s=''; 
				''|select @{n='LineNumber';e={$i}},@{n='LanguageCode';e={$Language}},@{n='Sentence';e={$D.words.text |%{$s=$s+"$_ "};$s}}}

            }
            Catch{
                "Something went wrong While extracting Text from Image, please try running the script again`nError Message : "+$E.Message
            }
    }
}

Function Translate-text()
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$True)]
		[String] $Text,
        [String] [validateSet('Arabic','Hindi','Japanese','Russian','Spanish','French',`
        'English','Korean','Urdu','Italian','Portuguese','German','Chinese Simplified')
        ]$From,
        [String] [validateSet('Arabic','Hindi','Japanese','Russian','Spanish','French',`
        'English','Korean','Urdu','Italian','Portuguese','German','Chinese Simplified')
        ]$To
)
	Begin{
					# Language codes hastable
					$LangCodes = @{'Arabic'='ar'
					'Chinese Simplified'='zh-CHS'
					'English'='en'
					'French'='fr'
					'German'='de'
					'Hindi'='hi'
					'Italian'='it'
					'Japanese'='ja'
					'Korean'='ko'
					'Portuguese'='pt'
					'Russian'='ru'
					'Spanish'='es'
					'Urdu'='ur'
					}
					
					# Secret Client ID and Key you get after Subscription	
					$ClientID = 'XXXXXXXXXXXXXXXXXXXX'
					$client_Secret = ‘XXXXXXXXXXXXXXXXXXXX'
					
					# If ClientId or Client_Secret has special characters, UrlEncode before sending request
					$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($ClientID)
					$client_SecretEncoded = [System.Web.HttpUtility]::UrlEncode($client_Secret) 
	}

	Process{
				ForEach($T in $Text)
				{	
					Try{
							# Azure Data Market URL which provide access tokens
							$URI = "https://datamarket.accesscontrol.windows.net/v2/OAuth2-13"
							
							# Body and Content Type of the request
							$Body = "grant_type=client_credentials&client_id=$clientIDEncoded&client_secret=$client_SecretEncoded&scope=http://api.microsofttranslator.com"
							$ContentType = "application/x-www-form-urlencoded"

							# Invoke REST method to Azure URI 
							$Access_Token=Invoke-RestMethod -Uri $Uri -Body $Body -ContentType $ContentType -Method Post
							
							# Header value with the access_token just recieved
							$Header = "Bearer " + $Access_Token.access_token
							
							# Invoke REST request to Microsoft Translator Service
							[string] $EncodedText = [System.Web.HttpUtility]::UrlEncode($T)
							[string] $uri = "http://api.microsofttranslator.com/v2/Http.svc/Translate?text=" + $EncodedText + "&from=" + $LangCodes.Item($From) + "&to=" + $LangCodes.Item($To);
					
							$Result = Invoke-RestMethod -Uri $URI -Headers @{Authorization = $Header} -ErrorVariable Error

							Return $Result.string.'#text'
					}
					catch
					{
							"Something went wrong While Translating Text, please try running the script again`nError Message : "+$Error.Message	
					}
				}
	}

}
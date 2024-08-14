# --------------------------------------------------------------------------------------------------------
# Script para listar utiliza��o de caixas de e-mail do exchange online (office 365).
# via3lr - luciano.rodrigues@v3c.com.br
# Usage: Invoque este script passando usu�rio e senha do administrador do office365 
# do cliente.
# Exemplo: powershell -file PS_Get_Mailboxes_Statistics.ps1 ti@cliente.com.br ClienteSenhaSuperSegura
# --------------------------------------------------------------------------------------------------------


# --------------------------------------------------------------------------------------------------------
# Parametros globais do script
# --------------------------------------------------------------------------------------------------------
Param(
    [Parameter(Mandatory=$True)] [string]$AdminUser,
    [Parameter(Mandatory=$True)] [string]$AdminPass,
	[Parameter(Mandatory=$True)] [int]$Top
)


# --------------------------------------------------------------------------------------------------------
# Compilando usu�rio e senha recebidos como uma credencial segura
# --------------------------------------------------------------------------------------------------------
$user = $AdminUser
$pass = ConvertTo-SecureString -AsPlainText -Force $AdminPass
$UserCredential = New-Object System.Management.Automation.PSCredential($user, $pass)

$table = @()
$processed = 0


# --------------------------------------------------------------------------------------------------------
# Conectando a uma sess�o do exchange online
# --------------------------------------------------------------------------------------------------------
try{
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session | out-null
}catch{
    Write-Host -ForegroundColor RED "Erro ao conectar ao Exchange Online. Encerrando o script."
    Write-Host $_.Exception.Message
	Remove-PSSession $Session
    Exit
}

# --------------------------------------------------------------------------------------------------------
# Obtendo as caixas de e-mail e as estat�sticas
# --------------------------------------------------------------------------------------------------------
Write-Host "Getting Mailboxes..."
$Mailboxes = Get-MailBox



Write-Host "Getting Statistics..."


$Mailboxes | %{
    # --------------------------------------------------------------------------------------------------------
    # Obtem a caixa de e-mail e o campo TotalItemSize
    # � necess�rio converter o campo TotalItemSize para conseguirmos 
    # --------------------------------------------------------------------------------------------------------
  #Write-Progress -Activity $_.Identity -PercentComplete ([math]::round(100/$Mailboxes.Count * $processed))
  $UsageMB =  [math]::round( (Get-MailboxStatistics -Identity $_.Identity).TotalItemSize.Value.toString().Split("(")[1].split(" ")[0].replace(",","")/1MB )
  $table += New-Object PSObject -Property @{Email=$_.PrimarySmtpAddress; Usage=$UsageMB}
  $processed += 1

}

# Fechando a conex�o remota.
Remove-PSSession $Session | out-null

#$table | Sort-Object -Property UsageMB -Descending | Out-GridView
$table | Sort-Object -Property Usage -Descending | Select-Object -Property Email,@{Name="Uso";Expression={"{0:n0} MB" -f  $_.Usage }} -First $Top



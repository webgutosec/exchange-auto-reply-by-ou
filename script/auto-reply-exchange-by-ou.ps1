# ================================================================
# Conectar ao Exchange Online| Connect to Exchange Online
# ================================================================
Import-Module ExchangeOnlineManagement
Import-Module ActiveDirectory
Connect-ExchangeOnline

# ================================================================
# Definir mensagens por "OU" | Define messages by OU
# ================================================================
$Messages = @{
    "OU=FOLDER1,OU=.FOLDER,DC=DOMAIN,DC=com,DC=br" = @{
        Internal = @"
A Área estará em recesso no dia DD/MM/AAAA. 

Em razão do feriado, no dia DD/MM/AAAA, não haverá expediente. 
As atividades serão retomadas no dia DD de MÊS de ANO. 
Para mais informações, consulte nossas publicações no site www.dominio.com.br. 

A empresa deseja a todos um excelente feriado.
"@

        External = @"
A Área estará em recesso no dia DD/MM/AAAA. 

Em razão do feriado, no dia DD/MM/AAAA, não haverá expediente. 
As atividades serão retomadas no dia DD de MÊS de ANO. 
Para mais informações, consulte nossas publicações no site www.dominio.com.br. 

A empresa deseja a todos um excelente feriado.
"@
    }
}

# ================================================================
# Período AAAA-MM-DD | Time period (YYYY-MM-DD)
# ================================================================
$StartDate = Get-Date "2026-04-18 00:00"
$EndDate   = Get-Date "2026-04-21 23:59"

# ================================================================
# Processar cada "OU" | Process each "OU"
# ================================================================
foreach ($OU in $Messages.Keys) {

    $InternalMessage = $Messages[$OU].Internal
    $ExternalMessage = $Messages[$OU].External

    Write-Host "`nSearching users in: $OU" -ForegroundColor Yellow

#
# Busca usuários habilitados na OU (sem recursão em sub-OUs) 
# Troque -SearchScope OneLevel por Subtree para incluir sub-OUs 
#
#
# Retrieve enabled users from the OU (no recursion in sub-OUs)
# Change -SearchScope OneLevel to Subtree to include sub-OUs
#

    try {
        $ADUsers = Get-ADUser -SearchBase $OU `
                              -SearchScope Subtree `
                              -Filter { Enabled -eq $true } `
                              -Properties UserPrincipalName
    }
    catch {
        Write-Host "Error retrieving users from AD: $OU" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
        continue
    }

# ================================================================
# Filtra apenas quem tem caixa no Exchange Online | Filter only users with Exchange mailbox
# ================================================================
    $Mailboxes = foreach ($ADUser in $ADUsers) {
        $Mailbox = Get-Mailbox -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
        if ($Mailbox -and $Mailbox.RecipientTypeDetails -eq "UserMailbox") { $Mailbox }
    }

    $TotalUsers = @($Mailboxes).Count
    $Counter    = 0

    Write-Host "Found $TotalUsers mailbox users in: $OU" -ForegroundColor Cyan

    foreach ($User in $Mailboxes) {
        $Counter++

        Write-Progress `
            -Activity "Configuring auto-replies [$OU]" `
            -Status "Processing $($User.PrimarySmtpAddress) ($Counter of $TotalUsers)" `
            -PercentComplete (($Counter / $TotalUsers) * 100)

        try {
            Set-MailboxAutoReplyConfiguration `
                -Identity $User.UserPrincipalName `
                -AutoReplyState Scheduled `
                -StartTime $StartDate `
                -EndTime $EndDate `
                -InternalMessage $InternalMessage `
                -ExternalMessage $ExternalMessage `
                -ExternalAudience All

            Write-Host "  Success: $($User.PrimarySmtpAddress)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Error: $($User.PrimarySmtpAddress)" -ForegroundColor Red
            Write-Host "    $($_.Exception.Message)" -ForegroundColor DarkRed
        }
    }

    Write-Progress -Activity "Configuring auto-replies [$OU]" -Completed
}

Write-Host "`nProcess completed." -ForegroundColor Cyan

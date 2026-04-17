# Exchange Auto Reply by OU (PowerShell)

Script simples em PowerShell para configurar **respostas automáticas (Out of Office)** em massa no Exchange Online, com base em usuários organizados por OU no Active Directory.

---

## 🇧🇷 Descrição

Este script realiza:

* Consulta usuários ativos no Active Directory por OU
* Valida se possuem mailbox no Exchange Online
* Configura respostas automáticas (interna e externa)
* Permite personalização por departamento/OU
* Execução em lote com barra de progresso

---

## 🇺🇸 Description

This PowerShell script:

* Retrieves active users from Active Directory by OU
* Validates Exchange Online mailboxes
* Configures automatic replies (internal & external)
* Supports OU-based customization
* Batch execution with progress tracking

---

## Tecnologias

* PowerShell
* Exchange Online
* Active Directory

---

## 🇧🇷 Pré-requisitos

* Módulo ExchangeOnlineManagement
* Módulo ActiveDirectory
* Permissão administrativa no Exchange Online
* Acesso ao AD

## 🇺🇸 Requirements
ExchangeOnlineManagement module
ActiveDirectory module
Administrative permissions in Exchange Online
Access to Active Directory

---

## 🚀 Como usar

1. Edite as OUs e a mensagem:

```powershell
$Mensagens = @{
    "OU=Finance,DC=empresa,DC=local" = @{
```

2. Ajuste o período:

```powershell
$Inicio = Get-Date "2026-04-18 00:00"
$Fim    = Get-Date "2026-04-21 23:59"
```

3. Execute:

```powershell
.\auto-reply-exchange-by-ou.ps1
```

Usage
1. Configure OUs and change message:

```powershell
$Messages = @{
    "OU=Finance,DC=empresa,DC=local" = @{
```

2. Set time period:

```powershell
$StartDate = Get-Date "2026-04-18 00:00"
$EndDate   = Get-Date "2026-04-21 23:59"
```

3. Run the script:
   
```powershell
.\auto-reply-exchange-by-ou.ps1
```

## 🇧🇷 Observações

* Script considera apenas usuários ativos
* Filtra apenas mailboxes do tipo "UserMailbox"
* Pode ser adaptado para múltiplas OUs

## 🇺🇸 Notes
Only enabled users are processed
Only "UserMailbox" recipients are considered
Can be extended to multiple OUs

---

## 👨‍💻 Autor

Augusto Argolo
https://www.linkedin.com/in/augusto-andradea/

---

## 🇧🇷 Contribuição

Sinta-se à vontade para sugerir melhorias ou adaptar o script.

## 🇺🇸 Contribution

Feel free to suggest improvements or adapt the script.

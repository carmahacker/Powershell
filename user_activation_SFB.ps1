# v2.1 - GilevAM
Start-Transcript "c:\scripts\log\Skype\user_activation_in_Skype_For_Business.log" -Append
#Подключение модулей АД и Скайп
#Import-Module activedirectory
#Import-Module 'C:\Program Files\Common Files\Skype for Business Server 2019\Modules\SkypeForBusiness\SkypeForBusiness.psd1'
#Задаём временной интервал для поиска новых уч.записей
$lastDay = ((Get-Date).AddDays(-100)).Date
#Ищем уч.записи с учётом фильтров по времени создания и наличия почтового адреса. Отключаем журнал бесед.
Write-Host "`tpharmasyntez"
$user = Get-ADUser -Server 'dc3.pharmasyntez.com' `
    -Filter { (mail -ne "null") -and (whencreated -ge $lastDay) } `
    -SearchBase 'OU=personal,DC=pharmasyntez,DC=com' -Properties msRTCSIP-UserEnabled,displayName | `
    Where-Object { 
    (!$_.DistinguishedName.Contains("OU=NewUsers,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=Technical_Users,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=Groups,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=External,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=Impersonal Accounts,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=Alias,OU=JSC_Pharmasyntez,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=Сивилаб,OU=JSC_Pharmasyntez,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=upak,OU=Цех №4 - СЛФ,OU=Иркутск,OU=JSC_Pharmasyntez,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=upak,OU=Цех №2 - ТЛФ,OU=Иркутск,OU=JSC_Pharmasyntez,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=upak,OU=Цех №3 - ТЛФ,OU=Иркутск,OU=JSC_Pharmasyntez,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=upak,OU=Цех малотоннажного производства,OU=Иркутск,OU=JSC_Pharmasyntez,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=upak,OU=Цех №5 - СТЛФ,OU=Иркутск,OU=JSC_Pharmasyntez,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=Impersonal Accounts,OU=Pharmasyntez-Tyumen,OU=personal,DC=pharmasyntez,DC=com")) -and `
    (!$_.DistinguishedName.Contains("OU=mail domains,OU=personal,DC=pharmasyntez,DC=com")) -and `
    ($_."msRTCSIP-UserEnabled" -notlike $true)
    } | Select-Object displayName | Sort-Object displayName

Write-Host "`t`tusers=$($user.count)"

foreach ($users in $user) {
    Write-Host "`t`t$($users.displayName)"
    Enable-CsUser -DomainController 'dc3.pharmasyntez.com' -identity $users.displayName -RegistrarPool sfb-fe01.pharmasyntez.com -SipAddressType sAMAccountName  -SipDomain pharmasyntez.com
}

#Ищем уч.записи с учётом фильтров по времени создания и наличия почтового адреса. Отключаем журнал бесед.
Write-Host "`tprimapharm"
$user2 = Get-ADUser -Server 'pri-dc3.primapharm.ru' `
    -Filter { (mail -ne "null") -and (whencreated -ge $lastDay) } `
    -SearchBase 'OU=Personnel,DC=primapharm,DC=ru' -Properties * | `
    Where-Object { 
    (!$_.DistinguishedName.Contains("OU=NewUsers,OU=Personnel,DC=primapharm,DC=ru")) -and `
    (!$_.DistinguishedName.Contains("OU=Alias,OU=Personnel,DC=primapharm,DC=ru")) -and `
    (!$_.DistinguishedName.Contains("OU=Groups,OU=Personnel,DC=primapharm,DC=ru")) -and `
    ($_."msRTCSIP-UserEnabled" -notlike $true) 
    } | Select-Object displayName | Sort-Object displayName  

Write-Host "`t`tusers=$($user2.count)"

foreach ($users2 in $user2) {    
    Write-Host "`t`t$($users2.displayName)"
    Enable-CsUser -DomainController 'pri-dc0.primapharm.ru' -identity $users2.displayName -RegistrarPool sfb-fe01.pharmasyntez.com -SipAddressType sAMAccountName  -SipDomain primapharm.ru 
}

#Ищем уч.записи с учётом фильтров по времени создания и наличия почтового адреса. Отключаем журнал бесед.
Write-Host "`tyotta-pharm"
$user3 = Get-ADUser -Server 'yp-dc3.yotta-pharm.ru' `
    -Filter { (mail -ne "null") -and (whencreated -ge $lastDay) } `
    -SearchBase 'OU=Personnel,DC=yotta-pharm,DC=ru' -Properties * | `
    Where-Object {
    (!$_.DistinguishedName.Contains("OU=NewUsers,OU=Personnel,DC=yotta-pharm,DC=ru")) -and `
    (!$_.DistinguishedName.Contains("OU=groups,OU=Personnel,DC=yotta-pharm,DC=ru")) -and `
    (!$_.DistinguishedName.Contains("OU=Impersonal Accounts,OU=Personnel,DC=yotta-pharm,DC=ru")) -and `
    ($_."msRTCSIP-UserEnabled" -notlike $true) 
    } | Select-Object displayName | Sort-Object displayName

Write-Host "`t`tusers=$($user3.count)"

foreach ($users3 in $user3) {    
    Write-Host "`t`t$($user3.displayName)"
    Enable-CsUser -DomainController 'yp-dc0.yotta-pharm.ru' -identity $users3.displayName -RegistrarPool sfb-fe01.pharmasyntez.com -SipAddressType sAMAccountName  -SipDomain yotta-pharm.ru 
}

#Ищем уч.записи с учётом фильтров по времени создания и наличия почтового адреса. Отключаем журнал бесед.
Write-Host "`tpropharm"
$user4 = Get-ADUser -Server 'pro-dc0.propharm.biz' `
    -Filter { (mail -ne "null") -and (whencreated -ge $lastDay) } `
    -SearchBase 'OU=Personnel,DC=propharm,DC=biz' -Properties * | `
    Where-Object { 
    (!$_.DistinguishedName.Contains("OU=NewUsers,OU=Personnel,DC=propharm,DC=biz")) -and `
    (!$_.DistinguishedName.Contains("OU=groups,OU=Personnel,DC=propharm,DC=biz")) -and `
    (!$_.DistinguishedName.Contains("OU=Impersonal Accounts,OU=Personnel,DC=propharm,DC=biz")) -and `
    ($_."msRTCSIP-UserEnabled" -notlike $true) 
    } | Select-Object displayName | Sort-Object displayName  

Write-Host "`t`tusers=$($user4.count)"

foreach ($users4 in $user4) {    
    Write-Host "`t`t$($users4.displayName)"
    Enable-CsUser -DomainController 'pro-dc0.propharm.biz' -identity $users4.displayName -RegistrarPool sfb-fe01.pharmasyntez.com -SipAddressType sAMAccountName  -SipDomain propharm.biz 
}

Stop-Transcript

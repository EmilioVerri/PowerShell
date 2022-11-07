
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)



<# 
.NAME
    AD Account Tool
.SYNOPSIS
    Check User by SamAccountName . Can Unlock User and lock user. Reset Password, enable nad disable user
.DESCRIPTION
    Checks user by SamAccountName. Returns Name, Last LogonDate, LockedOut Status, LockedoutTime, and Enabled Status. Allows User to be unlocked and locked. Locking of user is by increasing badpasswordcount. User is able to reset password for account. Enabling and disabling of Users are allowed.
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$CheckLockTool                   = New-Object system.Windows.Forms.Form
$CheckLockTool.ClientSize        = New-Object System.Drawing.Point(1200,1000)
$CheckLockTool.text              = "Gestione Utenti"
$CheckLockTool.TopMost           = $false
$CheckLockTool.ShowIcon = $True
$CheckLockTool.Icon = New-Object system.drawing.icon 'L:\Emilio Verri\Baxter.ico'
$CheckLockTool.AutoScroll = $true; 


$CheckLocked                     = New-Object system.Windows.Forms.Button
$CheckLocked.text                = "Cerca User"
$CheckLocked.width               = 150
$CheckLocked.height              = 50
$CheckLocked.location            = New-Object System.Drawing.Point(700,100)
$CheckLocked.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$CheckLocked.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$CheckLocked.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#fabc47")



$CercaNomeCognome                     = New-Object system.Windows.Forms.Button
$CercaNomeCognome.text                = "Cerca per nome e cognome"
$CercaNomeCognome.width               = 150
$CercaNomeCognome.height              = 50
$CercaNomeCognome.location            = New-Object System.Drawing.Point(700,175)
$CercaNomeCognome.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$CercaNomeCognome.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$CercaNomeCognome.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#fabc47")


$CercaEmployeeId                     = New-Object system.Windows.Forms.Button
$CercaEmployeeId.text                = "Cerca per EmployeeId"
$CercaEmployeeId.width               = 150
$CercaEmployeeId.height              = 50
$CercaEmployeeId.location            = New-Object System.Drawing.Point(700,245)
$CercaEmployeeId.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$CercaEmployeeId.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$CercaEmployeeId.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#fabc47")



$ClearRicerca                     = New-Object system.Windows.Forms.Button
$ClearRicerca.text                = "Cancella la ricerca"
$ClearRicerca.width               = 150
$ClearRicerca.height              = 50
$ClearRicerca.location            = New-Object System.Drawing.Point(250,325)
$ClearRicerca.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$ClearRicerca.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$ClearRicerca.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#00ffff")




$User                            = New-Object system.Windows.Forms.TextBox
$User.multiline                  = $false
$User.width                      = 200
$User.height                     = 25
$User.location                   = New-Object System.Drawing.Point(400,125)
$User.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Header                          = New-Object system.Windows.Forms.Label
$Header.text                     = "Inserisci lo User es:(verrie), con obbligo reset"
$Header.AutoSize                 = $true
$Header.width                    = 25
$Header.height                   = 10
$Header.location                 = New-Object System.Drawing.Point(400,105)
$Header.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$UserPerNome                            = New-Object system.Windows.Forms.TextBox
$UserPerNome.multiline                  = $false
$UserPerNome.width                      = 200
$UserPerNome.height                     = 25
$UserPerNome.location                   = New-Object System.Drawing.Point(400,200)
$UserPerNome.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$HeaderCercaPerNome                          = New-Object system.Windows.Forms.Label
$HeaderCercaPerNome.text                     = "Cerca es:(Rossi, Mario)"
$HeaderCercaPerNome.AutoSize                 = $true
$HeaderCercaPerNome.width                    = 25
$HeaderCercaPerNome.height                   = 10
$HeaderCercaPerNome.location                 = New-Object System.Drawing.Point(400,180)
$HeaderCercaPerNome.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$UserPerEmployeeId                            = New-Object system.Windows.Forms.TextBox
$UserPerEmployeeId.multiline                  = $false
$UserPerEmployeeId.width                      = 200
$UserPerEmployeeId.height                     = 25
$UserPerEmployeeId.location                   = New-Object System.Drawing.Point(400,270)
$UserPerEmployeeId.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$HeaderCercaPerEmployeeId                          = New-Object system.Windows.Forms.Label
$HeaderCercaPerEmployeeId.text                     = "Cerca per EmployeeId es:(802044)"
$HeaderCercaPerEmployeeId.AutoSize                 = $true
$HeaderCercaPerEmployeeId.width                    = 25
$HeaderCercaPerEmployeeId.height                   = 10
$HeaderCercaPerEmployeeId.location                 = New-Object System.Drawing.Point(400,250)
$HeaderCercaPerEmployeeId.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



$Header2                         = New-Object system.Windows.Forms.Label
$Header2.text                    = "Imposta nuova Password"
$Header2.AutoSize                = $true
$Header2.width                   = 30
$Header2.height                  = 10
$Header2.location                = New-Object System.Drawing.Point(400,850)
$Header2.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Password                        = New-Object system.Windows.Forms.TextBox
$Password.multiline              = $false
$Password.width                  = 174
$Password.height                 = 20
$Password.location               = New-Object System.Drawing.Point(400,875)
$Password.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SetPassword                     = New-Object system.Windows.Forms.Button
$SetPassword.text                = "Resetta Password"
$SetPassword.width               = 120
$SetPassword.height              = 30
$SetPassword.location            = New-Object System.Drawing.Point(400,900)
$SetPassword.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$SetPassword.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#FFFF00")

$Enable                     = New-Object system.Windows.Forms.Button
$Enable.text                = "Abilita Account"
$Enable.width               = 150
$Enable.height              = 50
$Enable.location            = New-Object System.Drawing.Point(600,850)
$Enable.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$Enable.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$Enable.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#7CFC00")

$Disable                     = New-Object system.Windows.Forms.Button
$Disable.text                = "Disabilita Account"
$Disable.width               = 150
$Disable.height              = 50
$Disable.location            = New-Object System.Drawing.Point(750,850)
$Disable.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$Disable.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$Disable.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("	#FF0000")

$Unlock                     = New-Object system.Windows.Forms.Button
$Unlock.text                = "Unlock Account"
$Unlock.width               = 150
$Unlock.height              = 50
$Unlock.location            = New-Object System.Drawing.Point(900,850)
$Unlock.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)
$Unlock.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$Unlock.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#FF7F50")





$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.ReadOnly               = $true
$TextBox2.multiline              = $true
$TextBox2.width                  = 700
$TextBox2.height                 = 500
$TextBox2.visible                = $true
$TextBox2.location               = New-Object System.Drawing.Point(400,325)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox2.Scrollbars = "Vertical" 


$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 200
$PictureBox1.height              = 200
$PictureBox1.location            = New-Object System.Drawing.Point(500,-50)
$PictureBox1.imageLocation       = "L:\Emilio Verri\Baxter.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom




<# RICERCA PER NOME E COGNOME

$HeaderNonRicordo                         = New-Object system.Windows.Forms.Label
$HeaderNonRicordo.text                    = "Cerca tra gli utenti, se ricordi il nome o il cognome?"
$HeaderNonRicordo.AutoSize                = $true
$HeaderNonRicordo.width                   = 25
$HeaderNonRicordo.height                  = 10
$HeaderNonRicordo.location                = New-Object System.Drawing.Point(600,100)
$HeaderNonRicordo.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',15)


$InserisciNome                         = New-Object system.Windows.Forms.Label
$InserisciNome.text                    = "Inserisci il Nome:"
$InserisciNome.AutoSize                = $true
$InserisciNome.width                   = 25
$InserisciNome.height                  = 10
$InserisciNome.location                = New-Object System.Drawing.Point(600,130)
$InserisciNome.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Nome                        = New-Object system.Windows.Forms.TextBox
$Nome.multiline              = $false
$Nome.width                  = 174
$Nome.height                 = 20
$Nome.location               = New-Object System.Drawing.Point(600,150)
$Nome.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



$LabelO                         = New-Object system.Windows.Forms.Label
$LabelO.text                    = "o"
$LabelO.AutoSize                = $true
$LabelO.width                   = 25
$LabelO.height                  = 10
$LabelO.location                = New-Object System.Drawing.Point(675,180)
$LabelO.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)


$InserisciCognome                         = New-Object system.Windows.Forms.Label
$InserisciCognome.text                    = "Inserisci il Cognome:"
$InserisciCognome.AutoSize                = $true
$InserisciCognome.width                   = 25
$InserisciCognome.height                  = 10
$InserisciCognome.location                = New-Object System.Drawing.Point(600,210)
$InserisciCognome.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$Cognome                        = New-Object system.Windows.Forms.TextBox
$Cognome.multiline              = $false
$Cognome.width                  = 174
$Cognome.height                 = 20
$Cognome.location               = New-Object System.Drawing.Point(600,230)
$Cognome.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ButtonNonRicordo                     = New-Object system.Windows.Forms.Button
$ButtonNonRicordo.text                = "Cerca Utente"
$ButtonNonRicordo.width               = 100
$ButtonNonRicordo.height              = 30
$ButtonNonRicordo.location            = New-Object System.Drawing.Point(600,270)
$ButtonNonRicordo.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',8)







$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.ReadOnly               = $true
$TextBox3.multiline              = $true
$TextBox3.width                  = 450
$TextBox3.height                 = 350
$TextBox3.visible                = $true
$TextBox3.location               = New-Object System.Drawing.Point(1100,100)
$TextBox3.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


#>




$CheckLockTool.controls.AddRange(@(<#$TextBox3,$LabelO,$InserisciNome,$InserisciCognome,$ButtonNonRicordo,$Nome,$Cognome,$HeaderNonRicordo,#>$Unlock,$ClearRicerca,$CercaEmployeeId,$HeaderCercaPerEmployeeId,$UserPerEmployeeId,$Disable,$CheckLocked,$CercaNomeCognome,$UserPerNome,$HeaderCercaPerNome,$User,$Header,$employeeType,$Enable,$PictureBox1 ,$EmployeeNumber,$UnlockAccount,$LockAccount,$Header2,$Password,$SetPassword,$TextBox2))


$CheckLocked.Add_Click({ tuttoInSieme })


$SetPassword.Add_Click({resetPassword})




$Enable.Add_Click({enableAccount})


$Disable.Add_Click({disabilitaAccount})



$CercaNomeCognome.Add_Click({cercaPerNomeCognome})



$CercaEmployeeId.Add_Click({cercaPerEmployeeId})


$ClearRicerca.Add_Click({cancellaRicerca})

$Unlock.Add_Click({unloccaAccount})



<# RICERCA PER NOME E COGNOME



$ButtonNonRicordo.Add_Click({nonRicordo})


function nonRicordo{




$nome=$Nome.Text

$cognome=$Cognome.Text


if($nome -and $cognome){

$result=Get-ADUser -Filter 'givenName  -eq $nome -and Surname -eq  $cognome'  -Properties *| Select Displayname,SamAccountName
$TextBox3.Text=($result|Format-Table| Out-String )

}elseif($nome){



$result=Get-ADUser -Filter 'givenName  -eq $nome'  -Properties * | Select Displayname,SamAccountName
$TextBox3.Text=($result|Format-Table| Out-String )


}elseif($cognome){


$result=Get-ADUser -Filter 'Surname -eq  $cognome'  -Properties *  | Select Displayname,SamAccountName


$TextBox3.Text=($result|Format-Table| Out-String )

}
else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome o un cognome es:(Mario)",0,"Errore",0x1)
return

}




}
#>


function unloccaAccount{
$groups=$User.Text


$secondGroups=$UserPerNome.Text

$thirdGroups=$UserPerEmployeeId.Text


if($groups){

$result=Get-ADUser -Identity $groups -Properties *

if($result){


$utente.$result.Name


$lock=$result.lockoutTime


if($lock -eq 0){
Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Utente'+ $utente+' è già unlocked', 'Messaggio Conferma', 'YesNo');

}
else{
Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Unlock User:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){
Unlock-ADAccount -Identity $groups
               $TextBox2.Text="Utente Unlocked"
}
else{
$TextBox2.Text="Unlocked Annullato"
}

}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}





}








elseif($secondGroups){




$result=Get-ADUser -Filter 'Displayname -eq $secondGroups' -Properties *

if($result){

$utente.$result.Name


$lock=$result.lockoutTime



if($lock -eq 0){
Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Utente'+ $utente+' è già unlocked', 'Messaggio Conferma', 'YesNo');

}
else{
Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Unlock User:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){

$information=Get-ADUser -Filter 'Displayname -eq $secondGroups'  -Properties SamAccountName

$filtraEnable=$information.SamAccountName


Unlock-ADAccount -Identity $filtraEnable

               $TextBox2.Text="Utente Disabilitato"
}
else{
$TextBox2.Text="Unlocked Annullato"
}

}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}





}










elseif($thirdGroups){




$takeUserSam=Get-ADUser -Filter 'EmployeeNumber  -eq $thirdGroups'  -Properties * 


if($takeUserSam){
$infoUser=$takeUserSam.SamAccountName



$utente.$takeUserSam.Name


$lock=$takeUserSam.lockoutTime



if($lock -eq 0){
Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Utente'+ $utente+' è già unlocked', 'Messaggio Conferma', 'YesNo');

}
else{
Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Unlock User:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){

Unlock-ADAccount -Identity $infoUser

               $TextBox2.Text="Utente Disabilitato"
}
else{
$TextBox2.Text="Unlocked Annullato"
}

}


}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}









}#fine elseif










else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome Utente/ cognome e nome es:(Rossi, Mario)/ EmployeeID",0,"Errore",0x1)
return


}




}

















function disabilitaAccount{


$groups=$User.Text


$secondGroups=$UserPerNome.Text

$thirdGroups=$UserPerEmployeeId.Text


if($groups){

$result=Get-ADUser -Identity $groups -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}
if($result){

$utente=$result.Name 


Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Stai disabilitando utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){
Disable-ADAccount -Identity $groups
               $TextBox2.Text="Utente Disabilitato"
}
else{
$TextBox2.Text="Disabilitazione annullata"
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}







}



elseif($secondGroups){




$result=Get-ADUser -Filter 'Displayname -eq $secondGroups' -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}

if($result){

$utente=$result.Name 


Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Stai disabilitando utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){


$information=Get-ADUser -Filter 'Displayname -eq $secondGroups'  -Properties SamAccountName

$filtraEnable=$information.SamAccountName


Disable-ADAccount -Identity $filtraEnable

               $TextBox2.Text="Utente Disabilitato"
}
else{
$TextBox2.Text="Disabilitazione annullata"
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}






}


elseif($thirdGroups){




$takeUserSam=Get-ADUser -Filter 'EmployeeNumber  -eq $thirdGroups'  -Properties * 

if($takeUserSam){
$infoUser=$takeUserSam.SamAccountName


$utente=$takeUserSam.Name 



Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Stai disabilitando utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){
Disable-ADAccount -Identity $infoUser
               $TextBox2.Text="Utente Disabilitato"
}
else{
$TextBox2.Text="Disabilitazione Annullata"
}

}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}



}




else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome Utente/ cognome e nome es:(Rossi, Mario)/ EmployeeID",0,"Errore",0x1)
return


}




}


function cancellaRicerca{
$TextBox2.Text=" "

}




function enableAccount{

$groups=$User.Text


$secondGroups=$UserPerNome.Text

$thirdGroups=$UserPerEmployeeId.Text


if($groups){



$result=Get-ADUser -Identity $groups -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}


if($result){

$utente=$result.Name 


Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Stai abilitando utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){
Set-ADUser -Identity $groups -Enabled $true
               $TextBox2.Text="Utente Abilitato"
}
else{
$TextBox2.Text="Annullata l'abilitazione dell'utente"
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}





}



elseif($secondGroups){




$result=Get-ADUser -Filter 'Displayname -eq $secondGroups' -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}

if($result){

$utente=$result.Name 


Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Stai abilitando utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){


$information=Get-ADUser -Filter 'Displayname -eq $secondGroups'  -Properties SamAccountName

$filtraEnable=$information.SamAccountName



Set-ADUser -Identity $filtraEnable -Enabled $true
               $TextBox2.Text="Utente Abilitato"
}
else{
$TextBox2.Text="Annullata l'abilitazione dell'utente"
}

}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}





}


elseif($thirdGroups){




$takeUserSam=Get-ADUser -Filter 'EmployeeNumber  -eq $thirdGroups'  -Properties * |Select SamAccountName

if($takeUserSam){

$infoUser=$takeUserSam.SamAccountName




$result=Get-ADUser -Identity $infoUser -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}






$utente=$result.Name 


Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Stai abilitando utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){
Set-ADUser -Identity $infoUser -Enabled $true
               $TextBox2.Text="Utente Abilitato"
}
else{
$TextBox2.Text="Annullata l'abilitazione dell'utente"
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2=" "
$wshell.Popup("Utente non trovato",0,"Errore",0x1)
return
}




}




else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome Utente/ cognome e nome es:(Rossi, Mario)/ EmployeeID",0,"Errore",0x1)
return


}



}






function resetPassword{

$groups=$User.Text


$secondGroups=$UserPerNome.Text


$thirdGroups=$UserPerEmployeeId.Text


if($groups){





$result=Get-ADUser -Identity $groups -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}


if($result){


$utente=$result.Name 


$password=$Password.Text



if($password){

Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Sei sicuro di voler cambiare password utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){
Set-ADAccountPassword -Identity $groups -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force);
Set-ADUser -Identity $groups -ChangePasswordAtLogon $true;
$TextBox2.Text="Cambio Password Avvenuto con successo, la nuova password: $password" 
}
else{
 $TextBox2.Text="Hai selezionato No, annullato il cambiamento password"
}








}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno una password da cambiare",0,"Errore",0x1)
return
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2.Text=" "
$wshell.Popup("Non posso cambiare la password ad un utente inesistente",0,"Errore",0x1)
return
}







}





elseif($secondGroups){






$result=Get-ADUser -Filter 'Displayname -eq $secondGroups' -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}

if($result){
$utente=$result.Name 


$password=$Password.Text



if($password){

Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Sei sicuro di voler cambiare password utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){


$change=Get-ADUser -Filter 'Displayname -eq $secondGroups'  -Properties SamAccountName

$filtraChange=$change.SamAccountName



Set-ADAccountPassword -Identity $filtraChange -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
$TextBox2.Text="Cambio Password Avvenuto con successo, la nuova password: $password" 
}
else{
 $TextBox2.Text="Hai selezionato No, annullato il cambiamento password"
}








}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno una password da cambiare",0,"Errore",0x1)
return
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2.Text=" "
$wshell.Popup("Non posso cambiare la password ad un utente inesistente",0,"Errore",0x1)
return
}




}



elseif($thirdGroups){



$takeUserSam=Get-ADUser -Filter 'EmployeeNumber  -eq $thirdGroups'  -Properties * |Select SamAccountName



if($takeUserSam){

$infoUser=$takeUserSam.SamAccountName




$result=Get-ADUser -Identity $infoUser -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}

$utente=$result.Name 


$password=$Password.Text



if($password){

Add-Type -AssemblyName PresentationFramework
$informazione=[System.Windows.MessageBox]::Show('Sei sicuro di voler cambiare password utente:'+ $utente, 'Messaggio Conferma', 'YesNo');
if($informazione -eq "Yes"){
Set-ADAccountPassword -Identity $infoUser -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
$TextBox2.Text="Cambio Password Avvenuto con successo, la nuova password: $password" 
}
else{
 $TextBox2.Text="Hai selezionato No, annullato il cambiamento password"
}








}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno una password da cambiare",0,"Errore",0x1)
return
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2.Text=" "
$wshell.Popup("Non posso cambiare la password ad un utente inesistente",0,"Errore",0x1)
return
}





}






else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome Utente/ cognome e nome es:(Rossi, Mario)/ EmployeeID",0,"Errore",0x1)
return
}

}




function cercaPerEmployeeId{


$save=$UserPerEmployeeId.Text

if($save){






$takeElementSam=Get-ADUser -Filter 'EmployeeNumber  -eq $save'  -Properties * |Select SamAccountName


if($takeElementSam){
$groups=$takeElementSam.SamAccountName



#1.1
$resulti=Get-ADUser -Identity $groups -Properties LastLogon



$email=$resulti.UserPrincipalName





#1.1




#1.2
$result=Get-ADUser -Identity $groups -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}






$utente=$result.Name 





$lastlogon=$result.LastLogon




$primo=$lastlogon.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#1.2


#1.3
$saveInfo=Get-ADUser -Identity $groups -Properties *




$manager=$saveInfo.Manager


$CharArray =$manager.Split(",")


$takeFirst=$CharArray[0]



#1.3


#1.4

$location=Get-ADUser -Identity $groups -Properties *


$locazione=$location.l


#1.4





#2
$result1=Get-ADUser -Identity $groups -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}

$prendi=$result1.PwdLastSet




$secondo=$prendi.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
#2


#3
$scadPassword=Get-ADUser -Identity $groups -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$save=$scadPassword.PwdLastSet.AddDays(120)


$expired=$save.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



$terzo="$expired "


#3


#4
$empNumber=Get-ADUser -Identity $groups -Properties employeeNumber


$num=$empNumber.employeeNumber




$quarto=$num

#4

#5
$empType=Get-ADUser -Identity $groups -Properties employeeType


$typee=$empType.employeeType 

$quinto=$typee

#5


#6
$employeID=Get-ADUser -Identity $groups -Properties * |Select SamAccountName

$idEmployeee=$employeID.SamAccountName




#6

#7
$phone =Get-ADUser -Identity $groups -Properties *




$telefono=$phone.telephoneNumber



#7



#8
$department =Get-ADUser -Identity $groups -Properties *




$Dipartimento=$department.Department



#8


#9 Job Title
$title =Get-ADUser -Identity $groups -Properties *




$Titolo=$title.Title



#9



#10 Account Abilitato
$abilitation=Get-ADUser -Identity $groups -Properties userAccountControl



$Abilitato=$abilitation.userAccountControl


if($Abilitato -eq "512"){

$account="Account abilitato"
}elseif($Abilitato -eq "514"){
$account ="Account disabilitato"
}else{
$account="account non abilitato ne disabilitato"
}


#Get-ADUser -Filter 'Name -like "*SvcAccount"'

#10



#11 Utente Locke
$lock =Get-ADUser -Identity $groups -Properties *




$Locked=$title.lockoutTime


if($Locked -eq 0){
$infoLocked="Utente Unlocked"

}else{
$infoLocked="!!!Utente Locked!!!"

}



#11





$spacing="----------------------------------------------"



#Tabella numero 1
$table = New-Object System.Data.Datatable

# Adding columns
[void]$table.Columns.Add("Nome Utente  ")
[void]$table.Columns.Add("Employee ID  ")
[void]$table.Columns.Add("Type Employee")
[void]$table.Columns.Add("samAccountName  ")
[void]$table.Columns.Add("----------------------------------------------")





# Adding rows
[void]$table.Rows.Add("$utente","$quarto","$quinto","$idEmployeee","$spacing")



$tabella1=($table | Out-String)



#Fine tabella numero 1





#Tabella numero 2


# Adding columns



$table2 = New-Object System.Data.Datatable

[void]$table2.Columns.Add("Email Utente    ")
[void]$table2.Columns.Add("Numero Telefono ")
[void]$table2.Columns.Add("Job Title       ")
[void]$table2.Columns.Add("Department      ")
[void]$table2.Columns.Add("Manager Utente  ")
[void]$table2.Columns.Add("Locazione Utente")
[void]$table2.Columns.Add("----------------------------------------------")




# Adding rows
[void]$table2.Rows.Add("$email","$telefono","$Titolo","$Dipartimento","$takeFirst","$locazione","$spacing")


$tabella2=($table2 | Out-String)


#Fine Tabella numero 2



#Tabella numero 3

$table3 = New-Object System.Data.Datatable


[void]$table3.Columns.Add("Ultimo Accesso        ")
[void]$table3.Columns.Add("Registrazione Password")
[void]$table3.Columns.Add("Scadenza Password     ")
[void]$table3.Columns.Add("Account Abilitato     ")
[void]$table3.Columns.Add("Account Locked        ")
[void]$table3.Columns.Add("----------------------------------------------")




# Adding rows
[void]$table3.Rows.Add("$primo","$secondo","$terzo","$account","$infoLocked","$spacing")


$tabella3=($table3 | Out-String)

#Fine Tabella numero 3



$TextBox2.Text=$tabella1+$tabella2+$tabella3





}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2.Text=" "
$wshell.Popup("La query nel Db non ho riprodotto risultati, verifica di aver inserito un EmployeeId Corretto",0,"Errore",0x1)

return
}

}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un EmployeeID",0,"Errore",0x1)
return
}






}








function cercaPerNomeCognome{

$groups=$UserPerNome.Text

if($groups){
$infoPerErrori=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties *


if($infoPerErrori){


#1.1



$resulti=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties LastLogon



$email=$resulti.UserPrincipalName





#1.1




#1.2




$result=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}






$utente=$result.Name 





$lastlogon=$result.LastLogon




$primo=$lastlogon.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#1.2


#1.3
$saveInfo=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties *




$manager=$saveInfo.Manager


$CharArray =$manager.Split(",")


$takeFirst=$CharArray[0]



#1.3


#1.4

$location=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties *


$locazione=$location.l


#1.4





#2






$result1=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}

$prendi=$result1.PwdLastSet




$secondo=$prendi.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
#2


#3
$scadPassword=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$save=$scadPassword.PwdLastSet.AddDays(120)


$expired=$save.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



$terzo="$expired "


#3




#4

$empNumber=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties employeeNumber


$num=$empNumber.employeeNumber




$quarto=$num

#4

#5



$empType=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties employeeType


$typee=$empType.employeeType 

$quinto=$typee

#5


#6



$employeID=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties * |Select SamAccountName

$idEmployeee=$employeID.SamAccountName




#6

#7




$phone =Get-ADUser -Filter 'Displayname -eq $groups'  -Properties *




$telefono=$phone.telephoneNumber



#7



#8
$department =Get-ADUser -Filter 'Displayname -eq $groups'  -Properties *




$Dipartimento=$department.Department



#8


#9 Job Title
$title =Get-ADUser -Filter 'Displayname -eq $groups'  -Properties *




$Titolo=$title.Title



#9



#10 Account Abilitato


$abilitation=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties userAccountControl



$Abilitato=$abilitation.userAccountControl



if($Abilitato -eq "512"){

$account="Account abilitato"
}elseif($Abilitato -eq "514"){
$account ="Account disabilitato"
}else{
$account="account non abilitato ne disabilitato"
}


#Get-ADUser -Filter 'Name -like "*SvcAccount"'

#10


#11 Utente Locke
$lock=Get-ADUser -Filter 'Displayname -eq $groups'  -Properties *




$Locked=$title.lockoutTime


if($Locked -eq 0){
$infoLocked="Utente Unlocked"

}else{
$infoLocked="!!!Utente Locked!!!"

}



#11





$spacing="----------------------------------------------"



#Tabella numero 1
$table = New-Object System.Data.Datatable

# Adding columns
[void]$table.Columns.Add("Nome Utente  ")
[void]$table.Columns.Add("Employee ID  ")
[void]$table.Columns.Add("Type Employee")
[void]$table.Columns.Add("samAccountName  ")
[void]$table.Columns.Add("----------------------------------------------")






# Adding rows
[void]$table.Rows.Add("$utente","$quarto","$quinto","$idEmployeee","$spacing")



$tabella1=($table | Out-String)



#Fine tabella numero 1





#Tabella numero 2


# Adding columns



$table2 = New-Object System.Data.Datatable

[void]$table2.Columns.Add("Email Utente    ")
[void]$table2.Columns.Add("Numero Telefono ")
[void]$table2.Columns.Add("Job Title       ")
[void]$table2.Columns.Add("Department      ")
[void]$table2.Columns.Add("Manager Utente  ")
[void]$table2.Columns.Add("Locazione Utente")
[void]$table2.Columns.Add("----------------------------------------------")




# Adding rows
[void]$table2.Rows.Add("$email","$telefono","$Titolo","$Dipartimento","$takeFirst","$locazione","$spacing")


$tabella2=($table2 | Out-String)


#Fine Tabella numero 2



#Tabella numero 3

$table3 = New-Object System.Data.Datatable


[void]$table3.Columns.Add("Ultimo Accesso        ")
[void]$table3.Columns.Add("Registrazione Password")
[void]$table3.Columns.Add("Scadenza Password     ")
[void]$table3.Columns.Add("Account Abilitato     ")
[void]$table3.Columns.Add("Account Locked        ")
[void]$table3.Columns.Add("----------------------------------------------")




# Adding rows
[void]$table3.Rows.Add("$primo","$secondo","$terzo","$account","$infoLocked","$spacing")


$tabella3=($table3 | Out-String)

#Fine Tabella numero 3



$TextBox2.Text=$tabella1+$tabella2+$tabella3







}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2.Text=" "
$wshell.Popup("La query nel Db non ho riprodotto risultati, verifica di aver inserito nome e cognome corretto",0,"Errore",0x1)

return
}

}else{


$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un cognome e nome es:(Rossi, Mario)",0,"Errore",0x1)
return



}










}



function tuttoInSieme{

$groups=$User.Text

if($groups){

$infoUtentePresente=Get-ADUser -Identity $groups -Properties *


if($infoUtentePresente){


#1.1
$resulti=Get-ADUser -Identity $groups -Properties LastLogon



$email=$resulti.UserPrincipalName





#1.1




#1.2
$result=Get-ADUser -Identity $groups -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}






$utente=$result.Name 





$lastlogon=$result.LastLogon




$primo=$lastlogon.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#1.2


#1.3
$saveInfo=Get-ADUser -Identity $groups -Properties *




$manager=$saveInfo.Manager


$CharArray =$manager.Split(",")


$takeFirst=$CharArray[0]



#1.3


#1.4

$location=Get-ADUser -Identity $groups -Properties *


$locazione=$location.l


#1.4





#2
$result1=Get-ADUser -Identity $groups -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}

$prendi=$result1.PwdLastSet




$secondo=$prendi.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
#2


#3
$scadPassword=Get-ADUser -Identity $groups -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$save=$scadPassword.PwdLastSet.AddDays(120)


$expired=$save.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



$terzo="$expired "


#3


#4
$empNumber=Get-ADUser -Identity $groups -Properties employeeNumber


$num=$empNumber.employeeNumber




$quarto=$num

#4

#5
$empType=Get-ADUser -Identity $groups -Properties employeeType


$typee=$empType.employeeType 

$quinto=$typee

#5


#6
$employeID=Get-ADUser -Identity $groups -Properties * |Select SamAccountName

$idEmployeee=$employeID.SamAccountName




#6

#7
$phone =Get-ADUser -Identity $groups -Properties *




$telefono=$phone.telephoneNumber



#7



#8
$department =Get-ADUser -Identity $groups -Properties *




$Dipartimento=$department.Department



#8


#9 Job Title
$title =Get-ADUser -Identity $groups -Properties *




$Titolo=$title.Title



#9



#10 Account Abilitato
$abilitation=Get-ADUser -Identity $groups -Properties userAccountControl



$Abilitato=$abilitation.userAccountControl


if($Abilitato -eq "512"){

$account="Account abilitato"
}elseif($Abilitato -eq "514"){
$account ="Account disabilitato"
}else{
$account="account non abilitato ne disabilitato"
}


#Get-ADUser -Filter 'Name -like "*SvcAccount"'

#10



#11 Utente Locke
$lock =Get-ADUser -Identity $groups -Properties *




$Locked=$title.lockoutTime


if($Locked -eq 0){
$infoLocked="Utente Unlocked"

}else{
$infoLocked="!!!Utente Locked!!!"

}



#11







$spacing="----------------------------------------------"



#Tabella numero 1
$table = New-Object System.Data.Datatable

# Adding columns
[void]$table.Columns.Add("Nome Utente  ")
[void]$table.Columns.Add("Employee ID  ")
[void]$table.Columns.Add("Type Employee")
[void]$table.Columns.Add("samAccountName  ")
[void]$table.Columns.Add("----------------------------------------------")





# Adding rows
[void]$table.Rows.Add("$utente","$quarto","$quinto","$idEmployeee","$spacing")



$tabella1=($table | Out-String)



#Fine tabella numero 1





#Tabella numero 2


# Adding columns



$table2 = New-Object System.Data.Datatable

[void]$table2.Columns.Add("Email Utente    ")
[void]$table2.Columns.Add("Numero Telefono ")
[void]$table2.Columns.Add("Job Title       ")
[void]$table2.Columns.Add("Department      ")
[void]$table2.Columns.Add("Manager Utente  ")
[void]$table2.Columns.Add("Locazione Utente")
[void]$table2.Columns.Add("----------------------------------------------")




# Adding rows
[void]$table2.Rows.Add("$email","$telefono","$Titolo","$Dipartimento","$takeFirst","$locazione","$spacing")


$tabella2=($table2 | Out-String)


#Fine Tabella numero 2



#Tabella numero 3

$table3 = New-Object System.Data.Datatable


[void]$table3.Columns.Add("Ultimo Accesso        ")
[void]$table3.Columns.Add("Registrazione Password")
[void]$table3.Columns.Add("Scadenza Password     ")
[void]$table3.Columns.Add("Account Abilitato     ")
[void]$table3.Columns.Add("Account Locked     ")
[void]$table3.Columns.Add("----------------------------------------------")




# Adding rows
[void]$table3.Rows.Add("$primo","$secondo","$terzo","$account","$infoLocked","$spacing")


$tabella3=($table3 | Out-String)

#Fine Tabella numero 3



$TextBox2.Text=$tabella1+$tabella2+$tabella3







}else{
$wshell = New-Object -ComObject Wscript.Shell
$TextBox2.Text=" "
$wshell.Popup("La query nel Db non ho riprodotto risultati, verifica di aver inserito SamAccountName corretto",0,"Errore",0x1)

return
}


}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome Utente es:(verrie)",0,"Errore",0x1)
return
}









}











#realizzato da Emilio Verri



#Write-Output
#endregion

[void]$CheckLockTool.ShowDialog()
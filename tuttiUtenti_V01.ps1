

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
<# 
.NAME
    Another
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(2048,1280)
$Form.text                       = "Ricerca Informazioni Più Utenti"
$Form.TopMost                    = $false
$Form.Icon = New-Object system.drawing.icon 'L:\Emilio Verri\Baxter.ico'
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$Form.AutoScroll = $true; 

#textbox

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $true
$TextBox2.width                  = 1400
$TextBox2.height                 = 900
$TextBox2.location               = New-Object System.Drawing.Point(100,100)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox2.ReadOnly               = $true;
$TextBox2.Scrollbars = "Vertical" 



$TextPath                        = New-Object system.Windows.Forms.TextBox
$TextPath.multiline              = $true
$TextPath.width                  = 300
$TextPath.height                 = 50
$TextPath.location               = New-Object System.Drawing.Point(1200,45)
$TextPath.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextPath.ReadOnly               = $false;



#fine textbox



#button utility


$Button                         = New-Object system.Windows.Forms.Button
$Button.text                    = "Search All User: "
$Button.width                   = 150
$Button.height                  = 53
$Button.location                = New-Object System.Drawing.Point(100,15)
$Button.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Button.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#7FFF00")


$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Canc: "
$Button1.width                   = 150
$Button1.height                  = 53
$Button1.location                = New-Object System.Drawing.Point(250,15)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Button1.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#de1f21")

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Export Csv: "
$Button2.width                   = 150
$Button2.height                  = 53
$Button2.location                = New-Object System.Drawing.Point(400,15)
$Button2.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Button2.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#00BFFF")

#fine button utility


#label ricerca
$LabelCerca                           = New-Object system.Windows.Forms.Label
$LabelCerca.text                      = "Scegli il metodo di ricerca:"
$LabelCerca.AutoSize                  = $true
$LabelCerca.width                     = 25
$LabelCerca.height                    = 10
$LabelCerca.location                  = New-Object System.Drawing.Point(650,10)
$LabelCerca.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$LabelCerca.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")




$LabelPath                           = New-Object system.Windows.Forms.Label
$LabelPath.text                      = "Inserisci il path per il csv:"
$LabelPath.AutoSize                  = $true
$LabelPath.width                     = 25
$LabelPath.height                    = 10
$LabelPath.location                  = New-Object System.Drawing.Point(1200,10)
$LabelPath.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$LabelPath.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")

#label ricerca



#checkbox
$Grosotto                       = New-Object system.Windows.Forms.CheckBox
$Grosotto.text                  = "Grosotto"
$Grosotto.AutoSize              = $false
$Grosotto.width                 = 150
$Grosotto.height                = 20
$Grosotto.location              = New-Object System.Drawing.Point(650,45)
$Grosotto.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Sondalo                       = New-Object system.Windows.Forms.CheckBox
$Sondalo.text                  = "Sondalo"
$Sondalo.AutoSize              = $false
$Sondalo.width                 = 150
$Sondalo.height                = 20
$Sondalo.location              = New-Object System.Drawing.Point(650,65)
$Sondalo.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SanVittore                       = New-Object system.Windows.Forms.CheckBox
$SanVittore.text                  = "SanVittore"
$SanVittore.AutoSize              = $false
$SanVittore.width                 = 150
$SanVittore.height                = 20
$SanVittore.location              = New-Object System.Drawing.Point(850,45)
$SanVittore.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)








$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 200
$PictureBox1.height              = 150
$PictureBox1.location            = New-Object System.Drawing.Point(700,1000)
$PictureBox1.imageLocation       = "L:\Emilio Verri\Baxter.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom


#fine checkbox


$Form.controls.AddRange(@($PictureBox1,$LabelPath,$TextPath,$Grosotto,$Sondalo,$SanVittore,$Button,$TextBox2,$Button1,$Button2,$LabelCerca))




$Button1.Add_MouseClick({ cancella })
$Button.Add_MouseClick({ cercaAllUtenti })
$Button2.Add_MouseClick({esportaCsv})


function cancella{
$TextBox2.Text=""

}


function cercaAllUtenti{
if($Grosotto.Checked -eq $true -or $Sondalo.Checked -eq $true -or $SanVittore.Checked -eq $true){
   if($Grosotto.Checked -eq $true){
   $result=Get-ADUser -Filter "City -eq 'Grosotto'"
   $save=$result.SamAccountName





   foreach($info in $save){


    $informazioni=Get-ADUser -Identity $info -Properties *


 $nameSurname=$informazioni.Name

    $department=$informazioni.Department

    $plant=$informazioni.City

    $peopleSo=$informazioni.EmployeeNumber

    $costo=$informazioni.extensionAttribute15

    $email=$informazioni.EmailAddress

    $cellul=$informazioni.MobilePhone

    $fisso=$informazioni.OfficePhone

    $manager=$informazioni.Manager

    #creazione account

    $infoCreated= Get-ADUser -Identity $info -Properties Created



    $creazioneAccount=$infoCreated.Created

    #fine creazione account


    $expired=$informazioni.AccountExpirationDate


    if($expired -eq $null){
    $expired="Account with no Expiration Date"
    }




#last logon
$takeLastLogon=Get-ADUser -Identity $info -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}


$lastlogon=$takeLastLogon.LastLogon




$infoLastLogon=$lastlogon.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)





#last logon


$infoPass=Get-ADUser -Identity $info -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}
$informationAgg=$infoPass.PwdLastSet


$aggiornamentoPassword=$informationAgg.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#scadenza password
$scadPassword=Get-ADUser -Identity $info -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$scadInfo=$scadPassword.PwdLastSet.AddDays(120)


$expiredPass=$scadInfo.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)

#scadenza password





$TextBox2.Text+="SamAccountName:  "+$info+" `r`n Nome Cognome:  "+$nameSurname+" `r`n Reparto:  "+$department+" `r`n Sede:  "+$plant+" `r`n PeopleSoft:  "+$peopleSo+
" `r`n Centro Costo:  "+$costo+" `r`n Mail:  "+$email+" `r`n Cellulare:  "+$cellul+" `r`n Telefono:  "+$fisso+
"`r`n Update Password:  "+$aggiornamentoPassword+" `r`n Create Account:  "+$creazioneAccount+"`r`n ExpiredPassword:  "+$expiredPass+" `r`n Manager:  "+$manager+" `r`n LastLogon:  "+$infoLastLogon+" `r`n ExpiredAccount:  "+$expired

$TextBox2.Text+="`r`n `r`n--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

   }
 
 



   


}

if($Sondalo.Checked -eq $true){
  $result=Get-ADUser -Filter "City -eq 'Sondalo'"
   $save=$result.SamAccountName

   foreach($info in $save){


    $informazioni=Get-ADUser -Identity $info -Properties *


 $nameSurname=$informazioni.Name

    $department=$informazioni.Department

    $plant=$informazioni.City

    $peopleSo=$informazioni.EmployeeNumber

    $costo=$informazioni.extensionAttribute15

    $email=$informazioni.EmailAddress

    $cellul=$informazioni.MobilePhone

    $fisso=$informazioni.OfficePhone

    $manager=$informazioni.Manager

    #creazione account

    $infoCreated= Get-ADUser -Identity $info -Properties Created



    $creazioneAccount=$infoCreated.Created

    #fine creazione account


    $expired=$informazioni.AccountExpirationDate


    if($expired -eq $null){
    $expired="Account with no Expiration Date"
    }




#last logon
$takeLastLogon=Get-ADUser -Identity $info -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}


$lastlogon=$takeLastLogon.LastLogon




$infoLastLogon=$lastlogon.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)





#last logon


$infoPass=Get-ADUser -Identity $info -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}
$informationAgg=$infoPass.PwdLastSet


$aggiornamentoPassword=$informationAgg.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#scadenza password
$scadPassword=Get-ADUser -Identity $info -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$scadInfo=$scadPassword.PwdLastSet.AddDays(120)


$expiredPass=$scadInfo.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)

#scadenza password





$TextBox2.Text+="SamAccountName:  "+$info+" `r`n Nome Cognome:  "+$nameSurname+" `r`n Reparto:  "+$department+" `r`n Sede:  "+$plant+" `r`n PeopleSoft:  "+$peopleSo+
" `r`n Centro Costo:  "+$costo+" `r`n Mail:  "+$email+" `r`n Cellulare:  "+$cellul+" `r`n Telefono:  "+$fisso+
"`r`n Update Password:  "+$aggiornamentoPassword+" `r`n Create Account:  "+$creazioneAccount+"`r`n ExpiredPassword:  "+$expiredPass+" `r`n Manager:  "+$manager+" `r`n LastLogon:  "+$infoLastLogon+" `r`n ExpiredAccount:  "+$expired

$TextBox2.Text+="`r`n `r`n--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

   }
 
 



   


}


if($SanVittore.Checked -eq $true){
 $result=Get-ADUser -Filter "City -eq '6534 S. Vittore'"
   $save=$result.SamAccountName

   foreach($info in $save){


    $informazioni=Get-ADUser -Identity $info -Properties *


 $nameSurname=$informazioni.Name

    $department=$informazioni.Department

    $plant=$informazioni.City

    $peopleSo=$informazioni.EmployeeNumber

    $costo=$informazioni.extensionAttribute15

    $email=$informazioni.EmailAddress

    $cellul=$informazioni.MobilePhone

    $fisso=$informazioni.OfficePhone

    $manager=$informazioni.Manager

    #creazione account

    $infoCreated= Get-ADUser -Identity $info -Properties Created



    $creazioneAccount=$infoCreated.Created

    #fine creazione account


    $expired=$informazioni.AccountExpirationDate


    if($expired -eq $null){
    $expired="Account with no Expiration Date"
    }




#last logon
$takeLastLogon=Get-ADUser -Identity $info -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}


$lastlogon=$takeLastLogon.LastLogon




$infoLastLogon=$lastlogon.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)





#last logon


$infoPass=Get-ADUser -Identity $info -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}
$informationAgg=$infoPass.PwdLastSet


$aggiornamentoPassword=$informationAgg.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#scadenza password
$scadPassword=Get-ADUser -Identity $info -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$scadInfo=$scadPassword.PwdLastSet.AddDays(120)


$expiredPass=$scadInfo.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)

#scadenza password





$TextBox2.Text+="SamAccountName:  "+$info+" `r`n Nome Cognome:  "+$nameSurname+" `r`n Reparto:  "+$department+" `r`n Sede:  "+$plant+" `r`n PeopleSoft:  "+$peopleSo+
" `r`n Centro Costo:  "+$costo+" `r`n Mail:  "+$email+" `r`n Cellulare:  "+$cellul+" `r`n Telefono:  "+$fisso+
"`r`n Update Password:  "+$aggiornamentoPassword+" `r`n Create Account:  "+$creazioneAccount+"`r`n ExpiredPassword:  "+$expiredPass+" `r`n Manager:  "+$manager+" `r`n LastLogon:  "+$infoLastLogon+" `r`n ExpiredAccount:  "+$expired

$TextBox2.Text+="`r`n `r`n--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

   }
 
 



   


}

}else{
  $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Checca almeno un filtro per eseguire la ricerca",0,"Errore",0x1)
           return

}



}








function esportaCsv{
if($Grosotto.Checked -eq $true -or $Sondalo.Checked -eq $true -or $SanVittore.Checked -eq $true){
#primo if lasciarlo importante per il controllo

$path=$TextPath.Text


if($path){




if($Grosotto.Checked -eq $true -and $Sondalo.Checked -eq $true -and $SanVittore.Checked -eq $true){
$info=Get-ADUser -Filter "City -eq '6534 S. Vittore' -or City -eq 'Grosotto' -or City -eq 'Sondalo' "
$result=foreach($save in $info){
Get-ADUser -Identity $save -Properties *
}




$save=$path+"\utentiSonSanGro.csv"

Write-Host($save)
if (Test-Path -Path $path) {
   $result | Export-csv "$save" -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}




 $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("File csv con utenti Sondalo, Grosotto e San Vittore Esportato",0,"Info",0x1)




############################################################################################################


}
elseif($Grosotto.Checked -eq $true -and $Sondalo.Checked -eq $true){
$info=Get-ADUser -Filter "City -eq 'Grosotto' -or City -eq 'Sondalo' "
$result=foreach($save in $info){
Get-ADUser -Identity $save -Properties *
}


if (Test-Path -Path $path) {

   $result | Export-csv $path\utentiGroSon.csv -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}

 $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("File csv con utenti Sondalo e Grosotto Esportato",0,"Info",0x1)



############################################################################################################
}elseif($Grosotto.Checked -eq $true -and $SanVittore.Checked -eq $true){
$info=Get-ADUser -Filter "City -eq '6534 S. Vittore' -or City -eq 'Grosotto' "
$result=foreach($save in $info){
Get-ADUser -Identity $save -Properties *
}


if (Test-Path -Path $path) {

   $result | Export-csv $path\utentiGroSan.csv -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}






 $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("File csv con utenti Grosotto e San Vittore Esportato",0,"Info",0x1)


##############################################################################################################
}
elseif($SanVittore.Checked -eq $true -and $Sondalo.Checked -eq $true){
$info=Get-ADUser -Filter "City -eq '6534 S. Vittore' -or City -eq 'Sondalo' "
$result=foreach($save in $info){
Get-ADUser -Identity $save -Properties *
}

if (Test-Path -Path $path) {

   $result | Export-csv $path\utentiSonSan.csv -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}

 $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("File csv con utenti Sondalo e San Vittore Esportato",0,"Info",0x1)

}##############################################################################################################
elseif($Grosotto.Checked -eq $true){
$info=Get-ADUser -Filter "City -eq 'Grosotto' "
$result=foreach($save in $info){
Get-ADUser -Identity $save -Properties *
}


if (Test-Path -Path $path) {

   $result | Export-csv $path\utentiGro.csv -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}




 $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("File csv con utenti Grosotto Esportato",0,"Info",0x1)

}#############################################################################################################

elseif($Sondalo.Checked -eq $true){
$info=Get-ADUser -Filter "City -eq 'Sondalo' "
$result=foreach($save in $info){
Get-ADUser -Identity $save -Properties *
}


if (Test-Path -Path $path) {

   $result | Export-csv $path\utentiSon.csv -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}


 $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("File csv con utenti Sondalo Esportato",0,"Info",0x1)

}##############################################################################################################
elseif($SanVittore.Checked -eq $true){
$info=Get-ADUser -Filter "City -eq '6534 S. Vittore' "
$result=foreach($save in $info){
Get-ADUser -Identity $save -Properties *
}

if (Test-Path -Path $path) {

   $result | Export-csv $path\utentiSan.csv -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}


 $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("File csv con utenti San Vittore Esportato",0,"Info",0x1)

}##############################################################################################################
else{
  $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Il codice non potrà mai arrivare fino a qua, se ci sei arrivato c'è qualche problema lato codice",0,"Errore",0x1)
           return

}
}else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Inserisci almeno un path",0,"Errore",0x1)
           return

}
}else{#lasciare questo else per il controllo
  $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Checca almeno un filtro per esportare come csv",0,"Errore",0x1)
           return


}


}






<#

$results=Get-ADUser -Identity verrie -Properties *  



 $results | Export-csv "L:\Emilio Verri\utente.csv" -NoTypeInformation


  $result=Get-ADUser -Filter "City -eq 'Grosotto'"


  Get-ADUser -Identity $info -Properties *

 #>


[void]$Form.ShowDialog()
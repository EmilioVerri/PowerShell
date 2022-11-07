

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

#checkbox#####################################


$CheckBox1                       = New-Object system.Windows.Forms.CheckBox
$CheckBox1.text                  = "Seleziona tutti (se selezionato, stampa tutto)"
$CheckBox1.AutoSize              = $false
$CheckBox1.width                 = 400
$CheckBox1.height                = 20
$CheckBox1.location              = New-Object System.Drawing.Point(1000,120)
$CheckBox1.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))




$NomeCognome                       = New-Object system.Windows.Forms.CheckBox
$NomeCognome.text                  = "Nome e Cognome"
$NomeCognome.AutoSize              = $false
$NomeCognome.width                 = 150
$NomeCognome.height                = 20
$NomeCognome.location              = New-Object System.Drawing.Point(1000,140)
$NomeCognome.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Reparto                       = New-Object system.Windows.Forms.CheckBox
$Reparto.text                  = "Reparto"
$Reparto.AutoSize              = $false
$Reparto.width                 = 150
$Reparto.height                = 20
$Reparto.location              = New-Object System.Drawing.Point(1000,160)
$Reparto.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$Sede                       = New-Object system.Windows.Forms.CheckBox
$Sede.text                  = "Sede"
$Sede.AutoSize              = $false
$Sede.width                 = 150
$Sede.height                = 20
$Sede.location              = New-Object System.Drawing.Point(1000,180)
$Sede.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$PeopleSoft                       = New-Object system.Windows.Forms.CheckBox
$PeopleSoft.text                  = "PeopleSoft"
$PeopleSoft.AutoSize              = $false
$PeopleSoft.width                 = 150
$PeopleSoft.height                = 20
$PeopleSoft.location              = New-Object System.Drawing.Point(1000,200)
$PeopleSoft.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CentroCosto                       = New-Object system.Windows.Forms.CheckBox
$CentroCosto.text                  = "Centro Costo"
$CentroCosto.AutoSize              = $false 
$CentroCosto.width                 = 150
$CentroCosto.height                = 20
$CentroCosto.location              = New-Object System.Drawing.Point(1000,220)
$CentroCosto.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$Mail                       = New-Object system.Windows.Forms.CheckBox
$Mail.text                  = "Mail"
$Mail.AutoSize              = $false 
$Mail.width                 = 150
$Mail.height                = 20
$Mail.location              = New-Object System.Drawing.Point(1000,240)
$Mail.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Cellulare                       = New-Object system.Windows.Forms.CheckBox
$Cellulare.text                  = "Cellulare"
$Cellulare.AutoSize              = $false 
$Cellulare.width                 = 150
$Cellulare.height                = 20
$Cellulare.location              = New-Object System.Drawing.Point(1000,260)
$Cellulare.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Telefono                       = New-Object system.Windows.Forms.CheckBox
$Telefono.text                  = "Telefono"
$Telefono.AutoSize              = $false 
$Telefono.width                 = 150
$Telefono.height                = 20
$Telefono.location              = New-Object System.Drawing.Point(1200,140)
$Telefono.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$UpdatePassword                       = New-Object system.Windows.Forms.CheckBox
$UpdatePassword.text                  = "Update Password"
$UpdatePassword.AutoSize              = $false 
$UpdatePassword.width                 = 150
$UpdatePassword.height                = 20
$UpdatePassword.location              = New-Object System.Drawing.Point(1200,160)
$UpdatePassword.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CreateAccount                       = New-Object system.Windows.Forms.CheckBox
$CreateAccount.text                  = "Create Account"
$CreateAccount.AutoSize              = $false 
$CreateAccount.width                 = 150
$CreateAccount.height                = 20
$CreateAccount.location              = New-Object System.Drawing.Point(1200,180)
$CreateAccount.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ExpiryPassword                      = New-Object system.Windows.Forms.CheckBox
$ExpiryPassword.text                  = "Expired Password"
$ExpiryPassword.AutoSize              = $false 
$ExpiryPassword.width                 = 150
$ExpiryPassword.height                = 20
$ExpiryPassword.location              = New-Object System.Drawing.Point(1200,200)
$ExpiryPassword.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Manager                       = New-Object system.Windows.Forms.CheckBox
$Manager.text                  = "Manager"
$Manager.AutoSize              = $false 
$Manager.width                 = 150
$Manager.height                = 20
$Manager.location              = New-Object System.Drawing.Point(1200,220)
$Manager.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LastLogon                       = New-Object system.Windows.Forms.CheckBox
$LastLogon.text                  = "Last Logon"
$LastLogon.AutoSize              = $false 
$LastLogon.width                 = 150
$LastLogon.height                = 20
$LastLogon.location              = New-Object System.Drawing.Point(1200,240)
$LastLogon.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ExpiredAccount                       = New-Object system.Windows.Forms.CheckBox
$ExpiredAccount.text                  = "Expired Account"
$ExpiredAccount.AutoSize              = $false 
$ExpiredAccount.width                 = 150
$ExpiredAccount.height                = 20
$ExpiredAccount.location              = New-Object System.Drawing.Point(1200,260)
$ExpiredAccount.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#fine checkbox################################

#textbox###########################

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $true
$TextBox1.width                  = 660
$TextBox1.height                 = 110
$TextBox1.location               = New-Object System.Drawing.Point(115,140)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $true
$TextBox2.width                  = 1400
$TextBox2.height                 = 700
$TextBox2.location               = New-Object System.Drawing.Point(115,300)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox2.ReadOnly               = $true;
$TextBox2.Scrollbars = "Vertical" 

#fine textbox######################

#utilità###########################

$Label                           = New-Object system.Windows.Forms.Label
$Label.text                      = "Inserisci Samaccountname separati da virgola es:(verrie,....)"
$Label.AutoSize                  = $true
$Label.width                     = 25
$Label.height                    = 10
$Label.location                  = New-Object System.Drawing.Point(200,81)
$Label.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$Label.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")


$LabelCerca                           = New-Object system.Windows.Forms.Label
$LabelCerca.text                      = "Scegli il metodo di ricerca:"
$LabelCerca.AutoSize                  = $true
$LabelCerca.width                     = 25
$LabelCerca.height                    = 10
$LabelCerca.location                  = New-Object System.Drawing.Point(1000,81)
$LabelCerca.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$LabelCerca.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")


$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 150
$PictureBox1.height              = 50
$PictureBox1.location            = New-Object System.Drawing.Point(700,20)
$PictureBox1.imageLocation       = "L:\Emilio Verri\Baxter.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom

#fine utilità#########################

#####################################################################################
###################################Button Ricerca####################################



$Cerca                         = New-Object system.Windows.Forms.Button
$Cerca.text                    = "Search"
$Cerca.width                   = 88
$Cerca.height                  = 30
$Cerca.location                = New-Object System.Drawing.Point(18,148)
$Cerca.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Cerca.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#030303")
$Cerca.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#b8e986")

$Cancella                         = New-Object system.Windows.Forms.Button
$Cancella.text                    = "Canc"
$Cancella.width                   = 86
$Cancella.height                  = 30
$Cancella.location                = New-Object System.Drawing.Point(20,194)
$Cancella.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$Cancella.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#e15b6e")



###################################Fine Button Ricerca###############################
#####################################################################################

#Form#


$Form.controls.AddRange(@($ExpiredAccount,$LastLogon,$Manager,$ExpiryPassword,$CreateAccount,$UpdatePassword,$Telefono,$TextBox1,$Label,$Prova,$Cerca,$Cancella,$TextBox2,$PictureBox1,$LabelCerca,$CheckBox1,$NomeCognome,$Reparto,$Sede,$PeopleSoft,$CentroCosto,$Mail,$Cellulare))


#Fine Form#

#call funzioni

$Cerca.Add_MouseClick({ cercaInserisciNellaTabella })
$Cancella.Add_MouseClick({ annulla })


#fine call funzioni








#fine funzione sistema#




#funzione annulla#
function annulla{

$TextBox1.Text=""
$TextBox2.Text=""
   
}
#fine funzione annulla#



#funzione ricerca nella tabella#


function cercaInserisciNellaTabella{#aperta graffa funzione
$TextBox2.Text=""

$save=$TextBox1.Text #nome





# fare lo spazio `r`n `r`n



if($save){#aperto if $save




$CharArray =$save.Split(",")



foreach ($save in $CharArray) {
$informazioni=""




#fine tutte le informazioni

if ($CheckBox1.Checked -eq $true){#checkbox seleziona tutti 
          
          #tutte le informazioni

    $informazioni= Get-ADUser -Identity $save -Properties *


    if($informazioni.Name -ne $null){

    
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

    $infoCreated= Get-ADUser -Identity $save -Properties Created



    $creazioneAccount=$infoCreated.Created

    #fine creazione account


    $expired=$informazioni.AccountExpirationDate


    if($expired -eq $null){
    $expired="Account with no Expiration Date"
    }




#last logon
$takeLastLogon=Get-ADUser -Identity $save -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}


$lastlogon=$takeLastLogon.LastLogon




$infoLastLogon=$lastlogon.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)





#last logon


$infoPass=Get-ADUser -Identity $save -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}
$informationAgg=$infoPass.PwdLastSet


$aggiornamentoPassword=$informationAgg.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#scadenza password
$scadPassword=Get-ADUser -Identity $save -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$scadInfo=$scadPassword.PwdLastSet.AddDays(120)


$expiredPass=$scadInfo.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)

#scadenza password





$TextBox2.Text+="SamAccountName:  "+$save+" `r`n Nome Cognome:  "+$nameSurname+" `r`n Reparto:  "+$department+" `r`n Sede:  "+$plant+" `r`n PeopleSoft:  "+$peopleSo+
" `r`n Centro Costo:  "+$costo+" `r`n Mail:  "+$email+" `r`n Cellulare:  "+$cellul+" `r`n Telefono:  "+$fisso+
"`r`n Update Password:  "+$aggiornamentoPassword+" `r`n Create Account:  "+$creazioneAccount+"`r`n ExpiredPassword:  "+$expiredPass+" `r`n Manager:  "+$manager+" `r`n LastLogon:  "+$infoLastLogon+" `r`n ExpiredAccount:  "+$expired

$TextBox2.Text+="`r`n `r`n--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"






}else{

$TextBox2.Text+="utente "+$save+" non presente"

$TextBox2.Text+="`r`n `r`n--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"


}
















        }
        else{#tutte le altre checkbox vanno qua dentro

        $informazioni=""

          




        #se non c'è nessuna di queste checcate
        if ($NomeCognome.Checked -eq $true -or $ExpiredAccount.Checked -eq $true -or $Reparto.Checked -eq $true -or $Sede.Checked -eq $true -or $PeopleSoft.Checked -eq $true -or $CentroCosto.Checked -eq $true -or $Mail.Checked -eq $true -or $Cellulare.Checked -eq $true -or $Telefono.Checked -eq $true -or $UpdatePassword.Checked -eq $true -or $CreateAccount.Checked -eq $true -or $ExpiryPassword.Checked -eq $true -or $Manager.Checked -eq $true  -or $LastLogon.Checked -eq $true){
          
          #infoutente

        


    $informazioni= Get-ADUser -Identity $save -Properties *


    if($informazioni.Name){
    
    $nameSurname=$informazioni.Name
  

    $department=$informazioni.Department

    $plant=$informazioni.City

    $peopleSo=$informazioni.EmployeeNumber

    $costo=$informazioni.extensionAttribute15

    $email=$informazioni.EmailAddress

    $cellul=$informazioni.MobilePhone

    $fisso=$informazioni.OfficePhone



    $manag=$informazioni.Manager




    #creazione account

    $infoCreated= Get-ADUser -Identity $save -Properties Created



    $creazioneAccount=$infoCreated.Created

    #fine creazione account


    $expired=$informazioni.AccountExpirationDate


    if($expired -eq $null){
    $expired="Account with no Expiration Date"
    }








#last logon
$takeLastLogon=Get-ADUser -Identity $save -Properties LastLogon | Select Name, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}


$lastlogonInfo=$takeLastLogon.LastLogon




$infoLastLogon=$lastlogonInfo.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)





#last logon


$infoPass=Get-ADUser -Identity $save -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}
$informationAgg=$infoPass.PwdLastSet


$aggiornamentoPassword=$informationAgg.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)



#scadenza password
$scadPassword=Get-ADUser -Identity $save -Properties PwdLastSet| Select Name, @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}}



$scadInfo=$scadPassword.PwdLastSet.AddDays(120)


$expiredPass=$scadInfo.ToString('dd/MM/yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)

#scadenza password




$TextBox2.Text+="SamAccountName:  "+$save



   if($NomeCognome.Checked -eq $true){

   $TextBox2.Text+=" `r`n Nome Cognome:  "+$nameSurname



         
          }

          if($Reparto.Checked -eq $true){
         $TextBox2.Text+=" `r`n Reparto:  "+$department


          }

          if($Sede.Checked -eq $true){
          $TextBox2.Text+=" `r`n Sede:  "+$plant
          }

            if($PeopleSoft.Checked -eq $true){
          $TextBox2.Text+=" `r`n PeopleSoft:  "+$peopleSo
          }

            if($CentroCosto.Checked -eq $true){
   $TextBox2.Text+=" `r`n Centro Costo:  "+$costo
          }

          if($Mail.Checked -eq $true){
       $TextBox2.Text+=" `r`n Mail:  "+$email
          }

         if($Cellulare.Checked -eq $true){
          $TextBox2.Text+=" `r`n Cellulare:  "+$cellul
          }

           if($Telefono.Checked -eq $true){
      $TextBox2.Text+=" `r`n Telefono:  "+$fisso
          }

           if($UpdatePassword.Checked -eq $true){
         $TextBox2.Text+="`r`n Update Password:  "+$aggiornamentoPassword

          }

           if($CreateAccount.Checked -eq $true){
          $TextBox2.Text+=" `r`n Create Account:  "+$creazioneAccount

          }

           if($ExpiryPassword.Checked -eq $true){
          $TextBox2.Text+="`r`n ExpiredPassword:  "+$expiredPass

          }

          



           if($Manager.Checked -eq $true){

         $TextBox2.Text+=" `r`n Manager:  "+$manag

          }

          if($LastLogon.Checked -eq $true){

      

         $TextBox2.Text+=" `r`n LastLogon:  "+$infoLastLogon

          }

          if($ExpiredAccount.Checked -eq $true){
          $TextBox2.Text+=" `r`n ExpiredAccount:  "+$expired

          }


          $TextBox2.Text+="`r`n `r`n--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"




}else{

$TextBox2.Text+="utente "+$save+" non presente"

$TextBox2.Text+="`r`n `r`n--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"





}



          #fineinfoutente


        }
        
        
        
        
        
        else{#messaggio errore del non check
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Checca almeno un filtro per eseguire la ricerca",0,"Errore",0x1)
           return
        }



        }





}











   




















}#chiuso if $save
else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un Samaccountname (Es: verrie)",0,"Errore",0x1)


}



}#chiusa graffa funzione



#fine funzione ricerca nella tabella




#region Logic 

#endregion

[void]$Form.ShowDialog()
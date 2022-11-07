

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
$Form.text                       = "Gestione Workstation"
$Form.TopMost                    = $false
$Form.Icon = New-Object system.drawing.icon 'L:\Emilio Verri\Baxter.ico'
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$Form.AutoScroll = $true; 


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
$Label.text                      = "Inserisci le postazioni IT separate da virgola"
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
$Cerca.text                    = "All"
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

$Description                         = New-Object system.Windows.Forms.Button
$Description.text                    = "Description"
$Description.width                   = 150
$Description.height                  = 30
$Description.location                = New-Object System.Drawing.Point(1000,135)
$Description.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$Description.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#99cbff")

$Sistem                         = New-Object system.Windows.Forms.Button
$Sistem.text                    = "Op. System"
$Sistem.width                   = 150
$Sistem.height                  = 30
$Sistem.location                = New-Object System.Drawing.Point(1000,180)
$Sistem.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$Sistem.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#99cbff")

$Plant                         = New-Object system.Windows.Forms.Button
$Plant.text                    = "Plant"
$Plant.width                   = 150
$Plant.height                  = 30
$Plant.location                = New-Object System.Drawing.Point(1200,180)
$Plant.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$Plant.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#99cbff")

$Enable                         = New-Object system.Windows.Forms.Button
$Enable.text                    = "Enable"
$Enable.width                   = 150
$Enable.height                  = 30
$Enable.location                = New-Object System.Drawing.Point(1200,135)
$Enable.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$Enable.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#99cbff")



###################################Fine Button Ricerca###############################
#####################################################################################

#Form#


$Form.controls.AddRange(@($TextBox1,$Label,$Prova,$Cerca,$Cancella,$TextBox2,$PictureBox1,$Description,$LabelCerca,$Sistem,$Plant,$Enable))


#Fine Form#

#call funzioni

$Cerca.Add_MouseClick({ cercaInserisciNellaTabella })
$Cancella.Add_MouseClick({ annulla })
$Description.Add_MouseClick({description})
$Sistem.Add_MouseClick({sistema})
$Plant.Add_MouseClick({plant})
$Enable.Add_MouseClick({enable})

#fine call funzioni








#function enable
function enable{

$save=$TextBox1.Text #nome





$TextBox2.Text="|------------------Nome------------------|------------------Abilitazione------------------|  `r`n `r`n"



if($save){#aperto if $save

#nome, description, sistema operativo, sede, raggruppa, owner





##########################################################


$CharArray =$save.Split(",")





foreach ($save in $CharArray) {
$info=""

$info=Get-ADComputer -Identity $save -Properties *

if($info.Description -eq $null){

$TextBox2.Text+="La Workstation $save non esistente `r`n `r`n"

}else{




$abilitato=$info.Enabled

if($abilitato -eq "True"){
$abilitato="Abilitato"
}else{
$abilitato="Disabilitato"
}





$TextBox2.Text+="|---------------"+$save+"----------------|"+"-----------------------"+$abilitato+"-------------------| `r`n `r`n"







}




}










}#chiuso if $save
else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome del pc (Es: ITGRTAB004)",0,"Errore",0x1)

}



}






#fine function enable












#funzione plant

function plant{


$save=$TextBox1.Text #nome





$TextBox2.Text="|------------------Nome------------------|------------------Sede------------------|  `r`n `r`n"



if($save){#aperto if $save

#nome, description, sistema operativo, sede, raggruppa, owner





##########################################################


$CharArray =$save.Split(",")





foreach ($save in $CharArray) {
$info=""

$info=Get-ADComputer -Identity $save -Properties *

if($info.Description -eq $null){

$TextBox2.Text+="La Workstation $save non esistente `r`n `r`n"

}else{



$locazione=$info.DistinguishedName




if($locazione -like '*Grosotto*') {
      $locazione="Grosotto"
} elseif($strVal -like '*Sondalo*')  {
    $locazione="Sondalo"
}else{
$locazione="Diverso Sondalo/Grosotto"
}



$TextBox2.Text+="|---------------"+$save+"----------------|"+"---------------"+$locazione+"-----------------| `r`n `r`n"







}




}










}#chiuso if $save
else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome del pc (Es: ITGRTAB004)",0,"Errore",0x1)

}



}




#fine funzione plant







#funzione cerca per description

function description{


$save=$TextBox1.Text #nome





$TextBox2.Text="|------------------Nome------------------|-----------------------------Description--------------------------|`r`n `r`n"



if($save){#aperto if $save

#nome, description, sistema operativo, sede, raggruppa, owner





##########################################################


$CharArray =$save.Split(",")





foreach ($save in $CharArray) {
$info=""

$info=Get-ADComputer -Identity $save -Properties *

if($info.Description -eq $null){

$TextBox2.Text+="La Workstation $save non esistente `r`n `r`n"

}else{

$description=$info.Description




$TextBox2.Text+="|---------------"+$save+"----------------|"+"----"+$description+"------|`r`n `r`n"







}




}










}#chiuso if $save
else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome del pc (Es: ITGRTAB004)",0,"Errore",0x1)

}





}
#fine funzione cerca per description


#inizio funzione sistema#


function sistema{



$save=$TextBox1.Text #nome





$TextBox2.Text="|------------------Nome------------------|---------------Sistema operativo---------------|  `r`n `r`n"



if($save){#aperto if $save

#nome, description, sistema operativo, sede, raggruppa, owner





##########################################################


$CharArray =$save.Split(",")





foreach ($save in $CharArray) {
$info=""

$info=Get-ADComputer -Identity $save -Properties *

if($info.Description -eq $null){

$TextBox2.Text+="La Workstation $save non esistente `r`n `r`n"

}else{


$operatingSistem=$info.OperatingSystem





$TextBox2.Text+="|---------------"+$save+"----------------|"+"-------------"+"$operatingSistem"+"---------|`r`n `r`n"







}




}










}#chiuso if $save
else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome del pc (Es: ITGRTAB004)",0,"Errore",0x1)

}


}




#fine funzione sistema#




#funzione annulla#
function annulla{

$TextBox1.Text=""
$TextBox2.Text=""
    
}
#fine funzione annulla#



#funzione ricerca nella tabella#


function cercaInserisciNellaTabella{#aperta graffa funzione

$save=$TextBox1.Text #nome





$TextBox2.Text="|------------------Nome------------------|-----------------------------Description--------------------------|---------------Sistema operativo---------------|------------------Sede------------------|------------------Abilitazione------------------|  `r`n `r`n"



if($save){#aperto if $save

#nome, description, sistema operativo, sede, raggruppa, owner





##########################################################


$CharArray =$save.Split(",")





foreach ($save in $CharArray) {
$info=""

$info=Get-ADComputer -Identity $save -Properties *

if($info.Description -eq $null){

$TextBox2.Text+="La Workstation $save non esistente `r`n `r`n"

}else{

$description=$info.Description

$operatingSistem=$info.OperatingSystem

$locazione=$info.DistinguishedName

$abilitato=$info.Enabled

if($abilitato -eq "True"){
$abilitato="Abilitato"
}else{
$abilitato="Disabilitato"
}



if($locazione -like '*Grosotto*') {
      $locazione="Grosotto"
} elseif($strVal -like '*Sondalo*')  {
    $locazione="Sondalo"
}else{
$locazione="Diverso Sondalo/Grosotto"
}



$TextBox2.Text+="|---------------"+$save+"----------------|"+"----"+$description+"------|"+"-------------"+"$operatingSistem"+"---------|"+"---------------"+$locazione+"-----------------|"+"-----------------------"+$abilitato+"-------------------| `r`n `r`n"







}




}










}#chiuso if $save
else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un nome del pc (Es: ITGRTAB004)",0,"Errore",0x1)

}



}#chiusa graffa funzione



#fine funzione ricerca nella tabella




#region Logic 

#endregion

[void]$Form.ShowDialog()
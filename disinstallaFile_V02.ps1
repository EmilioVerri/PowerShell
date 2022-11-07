

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
$Form.ClientSize                 = New-Object System.Drawing.Point(1500,1100)
$Form.text                       = "Disinstallazione Remota"
$Form.TopMost                    = $false
$Form.Icon = New-Object system.drawing.icon 'L:\Emilio Verri\Baxter.ico'
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$Form.AutoScroll = $true; 


#textbox###########################

$nomiPc                        = New-Object system.Windows.Forms.TextBox
$nomiPc.multiline              = $true
$nomiPc.width                  = 660
$nomiPc.height                 = 110
$nomiPc.location               = New-Object System.Drawing.Point(115,140)
$nomiPc.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$risultati                        = New-Object system.Windows.Forms.TextBox
$risultati.multiline              = $true
$risultati.width                  = 700
$risultati.height                 = 850
$risultati.location               = New-Object System.Drawing.Point(790,140)
$risultati.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$risultati.ReadOnly               = $true;
$risultati.Scrollbars = "Vertical" 




$disinstallaSoftware                        = New-Object system.Windows.Forms.TextBox
$disinstallaSoftware.multiline              = $true
$disinstallaSoftware.width                  = 660
$disinstallaSoftware.height                 = 80
$disinstallaSoftware.location               = New-Object System.Drawing.Point(115,325)
$disinstallaSoftware.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



$pathCsv                        = New-Object system.Windows.Forms.TextBox
$pathCsv.multiline              = $true
$pathCsv.width                  = 450
$pathCsv.height                 = 50
$pathCsv.location               = New-Object System.Drawing.Point(200,520)
$pathCsv.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)






$inserisciPathPerEsportareCsv                        = New-Object system.Windows.Forms.TextBox
$inserisciPathPerEsportareCsv.multiline              = $true
$inserisciPathPerEsportareCsv.width                  = 450
$inserisciPathPerEsportareCsv.height                 = 50
$inserisciPathPerEsportareCsv.location               = New-Object System.Drawing.Point(250,750)
$inserisciPathPerEsportareCsv.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#fine textbox######################



#checkbox
$Grosotto                       = New-Object system.Windows.Forms.CheckBox
$Grosotto.text                  = "Grosotto"
$Grosotto.AutoSize              = $false
$Grosotto.width                 = 100
$Grosotto.height                = 20
$Grosotto.location              = New-Object System.Drawing.Point(250,700)
$Grosotto.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Sondalo                       = New-Object system.Windows.Forms.CheckBox
$Sondalo.text                  = "Sondalo"
$Sondalo.AutoSize              = $false
$Sondalo.width                 = 100
$Sondalo.height                = 20
$Sondalo.location              = New-Object System.Drawing.Point(350,700)
$Sondalo.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SanVittore                       = New-Object system.Windows.Forms.CheckBox
$SanVittore.text                  = "SanVittore"
$SanVittore.AutoSize              = $false
$SanVittore.width                 = 100
$SanVittore.height                = 20
$SanVittore.location              = New-Object System.Drawing.Point(450,700)
$SanVittore.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



#fine checkbox ########################





#utilità###########################

$Label                           = New-Object system.Windows.Forms.Label
$Label.text                      = "Inserisci i nomi dei computer separati da virgola es:(itgrtab004,itgr0206na)"
$Label.AutoSize                  = $true
$Label.width                     = 25
$Label.height                    = 10
$Label.location                  = New-Object System.Drawing.Point(90,100)
$Label.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$Label.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")



$importCsvFileLabel                           = New-Object system.Windows.Forms.Label
$importCsvFileLabel.text                      = "Inserisci il path per importare il csv:"
$importCsvFileLabel.AutoSize                  = $true
$importCsvFileLabel.width                     = 25
$importCsvFileLabel.height                    = 10
$importCsvFileLabel.location                  = New-Object System.Drawing.Point(200,480)
$importCsvFileLabel.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$importCsvFileLabel.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")



$LabelDisinstallazione                           = New-Object system.Windows.Forms.Label
$LabelDisinstallazione.text                      = "Pacchetto disinstallazione es:'{ABFC9D02-F9D0-406F-9E3D-76B993E5E71F}'"
$LabelDisinstallazione.AutoSize                  = $true
$LabelDisinstallazione.width                     = 25
$LabelDisinstallazione.height                    = 10
$LabelDisinstallazione.location                  = New-Object System.Drawing.Point(90,290)
$LabelDisinstallazione.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$LabelDisinstallazione.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")


$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 150
$PictureBox1.height              = 50
$PictureBox1.location            = New-Object System.Drawing.Point(700,20)
$PictureBox1.imageLocation       = "L:\Emilio Verri\Baxter.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom




$LabelEsporta                           = New-Object system.Windows.Forms.Label
$LabelEsporta.text                      = "Inserisci Pacchetto disinstallazione sopra e seleziona i Plant se cui fare la ricerca'"
$LabelEsporta.AutoSize                  = $true
$LabelEsporta.width                     = 25
$LabelEsporta.height                    = 10
$LabelEsporta.location                  = New-Object System.Drawing.Point(10,650)
$LabelEsporta.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$LabelEsporta.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")


$LabelPathEsporta                           = New-Object system.Windows.Forms.Label
$LabelPathEsporta.text                      = "Inserisci Path:"
$LabelPathEsporta.AutoSize                  = $true
$LabelPathEsporta.width                     = 25
$LabelPathEsporta.height                    = 10
$LabelPathEsporta.location                  = New-Object System.Drawing.Point(20,750)
$LabelPathEsporta.Font                      = New-Object System.Drawing.Font('Microsoft PhagsPa',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Underline))
$LabelPathEsporta.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#f8e71c")



#fine utilità#########################

#####################################################################################
###################################Button Ricerca####################################



$Cerca                         = New-Object system.Windows.Forms.Button
$Cerca.text                    = "Elabora"
$Cerca.width                   = 88
$Cerca.height                  = 30
$Cerca.location                = New-Object System.Drawing.Point(18,138)
$Cerca.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Cerca.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#030303")
$Cerca.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#b8e986")


$Disinstalla                         = New-Object system.Windows.Forms.Button
$Disinstalla.text                    = "Disinstalla"
$Disinstalla.width                   = 95
$Disinstalla.height                  = 30
$Disinstalla.location                = New-Object System.Drawing.Point(18,325)
$Disinstalla.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Disinstalla.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#030303")
$Disinstalla.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#007fff")

$Cancella                         = New-Object system.Windows.Forms.Button
$Cancella.text                    = "Canc"
$Cancella.width                   = 86
$Cancella.height                  = 30
$Cancella.location                = New-Object System.Drawing.Point(20,178)
$Cancella.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$Cancella.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#e15b6e")



$importCsv                         = New-Object system.Windows.Forms.Button
$importCsv.text                    = "Importa CSV"
$importCsv.width                   = 130
$importCsv.height                  = 30
$importCsv.location                = New-Object System.Drawing.Point(20,500)
$importCsv.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$importCsv.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#FFA500")


$eseguiCsvImportato                         = New-Object system.Windows.Forms.Button
$eseguiCsvImportato.text                    = "Esegui Csv"
$eseguiCsvImportato.width                   = 130
$eseguiCsvImportato.height                  = 30
$eseguiCsvImportato.location                = New-Object System.Drawing.Point(20,540)
$eseguiCsvImportato.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$eseguiCsvImportato.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#FFA500")



$esportaCsv                         = New-Object system.Windows.Forms.Button
$esportaCsv.text                    = "Esporta CSV"
$esportaCsv.width                   = 130
$esportaCsv.height                  = 30
$esportaCsv.location                = New-Object System.Drawing.Point(20,700)
$esportaCsv.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
$esportaCsv.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("	#7CFC00")



###################################Fine Button Ricerca###############################
###################################Emilio Verri######################################
#####################################################################################

#Form#


$Form.controls.AddRange(@($LabelPathEsporta,$inserisciPathPerEsportareCsv,$Grosotto,$Sondalo,$SanVittore,$LabelEsporta,$importCsvFileLabel,$pathCsv,$eseguiCsvImportato,$esportaCsv,$importCsv,$LabelDisinstallazione,$disinstallaSoftware,$nomiPc,$Label,$Cancella,$risultati,$PictureBox1,$Cerca,$Disinstalla))


#Fine Form#

#call funzioni


#Questo programma è stato creato per disinstallare su tutti i pc ##veritas dlo agent##, però si può usare anche per disinstallare altri programmi

$Cerca.Add_MouseClick({ cercaTraiProgrammi })
$Cancella.Add_MouseClick({ annulla })
$Disinstalla.Add_MouseClick({disinstalla})
$importCsv.Add_MouseClick({importaCSV})
$eseguiCsvImportato.Add_MouseClick({eseguiCsvImportato})
$esportaCsv.Add_MouseClick({esportaCsvPacchetti})

#fine call funzioni




function esportaCsvPacchetti{
if(($Sondalo.Checked -eq $true)-or($Grosotto.Checked -eq $true)-or($SanVittore.Checked -eq $true)){





}else{
 $wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Seleziona almeno un Plant per effettuare l'export",0,"Error",0x1)
return
}


}










function eseguiCsvImportato{

$eseguiValore=1





$save=importaCSV(1)



if($save -eq 1){
return;
}elseif($save -eq "noInfo"){
 $wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Importa prima un file, prima di eseguirlo",0,"Error",0x1)
return
}


else{








if($save){
$nameSoftwareUnistall=$disinstallaSoftware.Text
if($nameSoftwareUnistall){
Add-Type -AssemblyName PresentationFramework
$yesornoUnistall=[System.Windows.MessageBox]::Show('Disinstallo il software:'+ $nameSoftwareUnistall, 'Messaggio Conferma', 'YesNo');
if($yesornoUnistall -eq "Yes"){

$CharArray =$save.Split(" ")
$risultati.Text=""
$info=""
foreach($pcTesto in $CharArray){

$info=Get-ADComputer -Identity $pcTesto -Properties *

if($info -eq ""){
$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+"Il pc:"+"  "+$pcTesto+"   "+"non esistente"+"`r`n"+"`r`n"

}else{
#controllo se esiste il programma da disinstallare nel computer


$programmi=Get-WmiObject Win32_Product -ComputerName $pcTesto | Select-Object -Property IdentifyingNumber





$printTrue=""


foreach($controllaSoftware in $programmi){



$nameSoftware=$controllaSoftware.IdentifyingNumber




if($nameSoftware -eq $nameSoftwareUnistall){
$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+ "`r`n"+"Inizio Disinstallazione Programma:"+"    "+$nameSoftwareUnistall+"   "+ "sul computer:"+ "  " + $pcTesto+"`r`n"+"`r`n"

$printTrue="true"

}else{



}

}



if($printTrue -eq "true"){#se non si è sicuri attenti ad eseguire questa parte di codice, perchè disinstallerà il programma




(Get-WmiObject Win32_Product -ComputerName $pcTesto | Where-Object {$_.IdentifyingNumber -eq $nameSoftwareUnistall}).Uninstall()


$risultati.Text+="Disinstallazione del Programma:"+"  "+$nameSoftwareUnistall+"   "+"avvenuta con successo"+"  "+"sul pc:"+$pcTesto+"`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"


}


else{


$testConnection=Test-Connection -ComputerName $pcTesto

if($testConnection){

$risultati.Text+="Programma:"+"  "+$nameSoftwareUnistall+"   "+"non presente nel computer"+"  "+$pcTesto+"`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"

}else{

$risultati.Text+="Computer:"+"  "+$pcTesto+"   "+"Spento"+"`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"


}

Write-Host("non disinstallo perchè o il computer è spento o non ha il programma")

}



}
$info=""




}#chiusura foreach

}else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Annullo la disinstallazione",0,"Info",0x1)
           return
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Inserisci almeno un software da disinstallare",0,"Errore",0x1)
           return
}
}else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Non è presente nessun pc nel file csv",0,"Errore",0x1)
           return

}








}#chiusura else $save


#C:\Users\verrie\Desktop\provaCsv.csv   path di prova


}



#importare i computer come csv
function importaCSV($controllo){
$inserisciPathCsv=$pathCsv.Text

$infoImportate = Import-Csv -Path $inserisciPathCsv

$saveInformazioniImportate=$infoImportate.computer

$CharArray =$saveInformazioniImportate.Split(" ")


#$risultati.Text+=$CharArray

#C:\Users\verrie\Desktop\provaCsv.csv   path di prova
if($controllo -eq 1){
if($CharArray){
return $CharArray
}else{

$test="noInfo"
return $test
}



}else{



if($inserisciPathCsv){



if (Test-Path -Path $inserisciPathCsv) {





$information=""


foreach($info in $CharArray){


$information=Get-ADComputer -Identity $info -Properties *


if($information -eq ""){

$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+"Computer:"+" "+" "+$info+"  "+"  "+"non esistente"+ "`r`n"

}#controllo se non esiste il computer
else{







$print=Get-WmiObject Win32_Product -ComputerName $info | Select-Object -Property IdentifyingNumber, Name



$count=0

foreach($takeinfo in $print){




if($count -eq 0){
$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+"Informazioni Computer:"+" "+" "+$info+ "`r`n"+ "`r`n"+$takeinfo.IdentifyingNumber+" "+" "+$takeinfo.Name
}else{
$risultati.Text+="`r`n"+$takeinfo.IdentifyingNumber+" "+" "+$takeinfo.Name
}

$count++
}







}#nel caso in cui esista il computer

$information=""
}itgr0139na
}

 else {
  $wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Inserisci un path esistente",0,"Error",0x1)
return
}



}else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Inserisci almeno un path",0,"Info",0x1)
           return

}

}

}




function disinstalla{



$pcTesto=$nomiPc.Text

if($pcTesto){#if esiste almeno un pc
$nameSoftwareUnistall=$disinstallaSoftware.Text

if($nameSoftwareUnistall){





Add-Type -AssemblyName PresentationFramework
$yesornoUnistall=[System.Windows.MessageBox]::Show('Disinstallo il software:'+ $nameSoftwareUnistall, 'Messaggio Conferma', 'YesNo');
if($yesornoUnistall -eq "Yes"){


$CharArray =$pcTesto.Split(",")



$risultati.Text=""
$info=""
foreach ($pcTesto in $CharArray) {
$info=Get-ADComputer -Identity $pcTesto -Properties *

if($info -eq ""){
$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+"Il pc:"+"  "+$pcTesto+"   "+"non esistente"+"`r`n"+"`r`n"

}else{
#controllo se esiste il programma da disinstallare nel computerù


$programmi=Get-WmiObject Win32_Product -ComputerName $pcTesto | Select-Object -Property IdentifyingNumber




$printTrue=""

foreach($controllaSoftware in $programmi){



$nameSoftware=$controllaSoftware.IdentifyingNumber




if($nameSoftware -eq $nameSoftwareUnistall){
$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+ "`r`n"+"Inizio Disinstallazione Programma:"+"    "+$nameSoftwareUnistall+"   "+ "sul computer:"+ "  " + $pcTesto+"`r`n"+"`r`n"

$printTrue="true"

}else{


}

}





if($printTrue -eq "true"){#se non si è sicuri attenti ad eseguire questa parte di codice, perchè disinstallerà il programma

Write-Host("adesso disinstallo")



(Get-WmiObject Win32_Product -ComputerName $pcTesto | Where-Object {$_.IdentifyingNumber -eq $nameSoftwareUnistall}).Uninstall()



$risultati.Text+="Disinstallazione del Programma:"+"  "+$nameSoftwareUnistall+"   "+"avvenuta con successo"+"  "+"sul pc:"+$pcTesto+"`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"

}
else{

$testConnection=Test-Connection -ComputerName $pcTesto

if($testConnection){

$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"Programma:"+"  "+$nameSoftwareUnistall+"   "+"non presente nel computer"+"  "+$pcTesto+"`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"
}else{

$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"Computer:"+"  "+$pcTesto+"   "+"Spento"+"`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"

}

Write-Host("non disinstallo perchè o il computer è spento o non ha il programma")

}









}
$info=""
}












$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Disinstallazione Avvenuta con successo su tutti i computer",0,"Info",0x1)
           return
}
else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Annullamento Disinstallazione del Software",0,"Info",0x1)
           return
}





}else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Inserisci almeno un nome software per effettuare la disinstallazione",0,"Errore",0x1)
           return
}



}else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Inserisci almeno 1 computer per effettuare la ricerca",0,"Errore",0x1)
           return

}#chiusura else esiste almeno un pc





}








function annulla{
$risultati.Text=""
}





function cercaTraiProgrammi{
$pcTesto=$nomiPc.Text





if($pcTesto) {




$CharArray =$pcTesto.Split(",")

$info=""

foreach ($pcTesto in $CharArray) {

$info=Get-ADComputer -Identity $pcTesto -Properties *


if($info -eq ""){




$risultati.Text+="`r`n"+"----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+ 
"`r`n"+"La Workstation $pcTesto non esistente `r`n `r`n"

}else{

$programmi=Get-WmiObject Win32_Product -ComputerName $pcTesto | Select-Object -Property IdentifyingNumber, Name
$count=0;



foreach($programma in $programmi){

$numeroProgramma=$programma.IdentifyingNumber
$nomeProgramma=$programma.Name


#stampo le applicazioni installate sul pc indicato
if($count -eq 0){
                   
$risultati.Text+=  "`r`n"+"-----------------------------------------------------------------------------------------------------------------------------------------------------------------------"+"`r`n"+ "`r`n"+"Informazioni Computer:"+ ' ' +  ' '+  ' '+  ' ' + $pcTesto+ "`r`n"+ "`r`n"+$numeroProgramma+ ' ' +  ' '+  ' '+  ' ' +$nomeProgramma+ "`r`n"
}else{

$risultati.Text+=$numeroProgramma+ ' ' +  ' '+  ' '+  ' ' +$nomeProgramma+ "`r`n"
}
$count++;

}#chiusura foreach
}#chiusura else inseriti pc non esistenti


$info=""
}#se vengono inseriti pc non esistenti



}#chiusura foreach divisione ,
else{
$wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Inserisci almeno 1 computer per effettuare la ricerca",0,"Errore",0x1)
           return

}#chiusura else se non viene inserito nessun pc

}#chiusura funzione


#fine call funzioni









#region Logic 

#endregion

[void]$Form.ShowDialog()
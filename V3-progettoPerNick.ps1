

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
<# 
.NAME
    Fork of Test Ping
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form1                           = New-Object system.Windows.Forms.Form
$Form1.ClientSize                = New-Object System.Drawing.Point(700,800)
$Form1.text                      = "Gestione Gruppi"
$Form1.TopMost                   = $false
$Form1.ShowIcon = $True
$Form1.Icon = New-Object system.drawing.icon 'L:\Emilio Verri\Baxter.ico'
$Form1.AutoScroll = $true; 

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Esporta Csv"
$Button1.width                   = 150
$Button1.height                  = 40
$Button1.location                = New-Object System.Drawing.Point(20,35)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button1.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#FFFF00")


$ButtonCerca                         = New-Object system.Windows.Forms.Button
$ButtonCerca.text                    = "Verifica Utenti"
$ButtonCerca.width                   = 150
$ButtonCerca.height                  = 40
$ButtonCerca.location                = New-Object System.Drawing.Point(20,87)
$ButtonCerca.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$ButtonCerca.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#7FFF00")


$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Inserisci gruppi separati da virgola:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(200,27)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 400
$TextBox1.height                 = 40
$TextBox1.location               = New-Object System.Drawing.Point(200,49)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',12)



$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Inserisci percorso dove scaricare csv Esempio:(C:\Users\tuoNome\Desktop):"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(200,80)
$Label3.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 400
$TextBox3.height                 = 40
$TextBox3.location               = New-Object System.Drawing.Point(200,100)
$TextBox3.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',12)




$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Informazioni Utenti:"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(20,130)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.ReadOnly               = $true
$TextBox2.multiline              = $true
$TextBox2.width                  = 600
$TextBox2.height                 = 400
$TextBox2.visible                = $true
$TextBox2.location               = New-Object System.Drawing.Point(20,150)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$TextBox2.Scrollbars = "Vertical" 


$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 200
$PictureBox1.height              = 200
$PictureBox1.location            = New-Object System.Drawing.Point(200,500)
$PictureBox1.imageLocation       = "L:\Emilio Verri\Baxter.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom



$ButtonClear                         = New-Object system.Windows.Forms.Button
$ButtonClear.text                    = "Cancella Informazioni:"
$ButtonClear.width                   = 150
$ButtonClear.height                  = 53
$ButtonClear.location                = New-Object System.Drawing.Point(470,550)
$ButtonClear.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$ButtonClear.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#00BFFF")





<#
$ClearButton                         = New-Object system.Windows.Forms.Button
$ClearButton.text                    = "Reset"
$ClearButton.width                   = 110
$ClearButton.height                  = 40
$ClearButton.location                = New-Object System.Drawing.Point(237,550)
$ClearButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$ClearButton.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
#>





$Form1.controls.AddRange(@($Button1,$TextBox1,$Label1,$Label3,$Label2,$TextBox2,$TextBox3,$ClearButton,$PictureBox1,$ButtonCerca,$ButtonClear))



$Button1.Add_Click({ utenti })




$ButtonCerca.Add_Click({ informazioni })


$ButtonClear.Add_Click({cancellaInformazioni})

<#
$ClearButton.Add_Click({ clear })




function clear{

$TextBox1.Clear();

}#>


function cancellaInformazioni{
$TextBox2.Text=""

}


function informazioni{
$groups=$TextBox1.Text

if($groups){







$CharArray =$groups.Split(",")





$results = foreach ($group in $CharArray) {
Get-ADGroupMember $group | select samaccountname, name, @{n='GroupName';e={$group}}, @{n='Description';e={(Get-ADGroup $group -Properties description).description}}
$save=Get-ADGroupMember $group | select samaccountname, name, @{n='GroupName';e={$group}}, @{n='Description';e={(Get-ADGroup $group -Properties description).description}}
}

if($save){
Write-Host ($results | Format-Table | Out-String)

$save=($results | Format-Table | Out-String)





$TextBox2.Text=$save
}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("La query non ha riprodotto risultati, controlla se il gruppo è corretto",0,"Error",0x1)
return
}

















}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un gruppo",0,"Error",0x1)
return
}


}





function utenti{

$groups=$TextBox1.Text

if($groups){

$path=$TextBox3.Text


if($path){


$CharArray =$groups.Split(",")





$results = foreach ($group in $CharArray) {
Get-ADGroupMember $group | select samaccountname, name, @{n='GroupName';e={$group}}, @{n='Description';e={(Get-ADGroup $group -Properties description).description}}
}






Write-Host ($results | Format-Table | Out-String)

$save=($results | Format-Table | Out-String)








if (Test-Path -Path $path) {
$TextBox2.Text=$save
   $results | Export-csv $path\utente.csv -NoTypeInformation
} else {
  $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path esistente dove scaricare il csv",0,"Error",0x1)
return
}


















}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci un path dove scaricare il csv",0,"Error",0x1)
return

}






}else{
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Inserisci almeno un gruppo",0,"Error",0x1)
return
}





}



#region Logic 



#Write your logic code here

#endregion

[void]$Form1.ShowDialog()
if (test-path -path d:\flag.flg){new-variable -name mf -value "d:"}
if (test-path -path e:\flag.flg){new-variable -name mf -Value "e:"}
if (test-path -path f:\flag.flg){new-variable -name mf -Value "f:"}
if (test-path -path g:\flag.flg){new-variable -name mf -Value "g:"}
if (test-path -path g:\flag.flg){new-variable -name mf -Value "h:"}
if (test-path -path i:\flag.flg){new-variable -name mf -Value "i:"}
if (test-path -path j:\flag.flg){new-variable -name mf -Value "j:"}
if (test-path -path k:\flag.flg){new-variable -name mf -Value "k:"}
if (test-path -path l:\flag.flg){new-variable -name mf -Value "l:"}
if (test-path -path m:\flag.flg){new-variable -name mf -Value "m:"}
if (test-path -path n:\flag.flg){new-variable -name mf -Value "n:"}
if (test-path -path o:\flag.flg){new-variable -name mf -Value "o:"}
if (test-path -path p:\flag.flg){new-variable -name mf -Value "p:"}
if (test-path -path q:\flag.flg){new-variable -name mf -Value "q:"}
if (test-path -path r:\flag.flg){new-variable -name mf -Value "r:"}
if (test-path -path s:\flag.flg){new-variable -name mf -Value "s:"}
if (test-path -path t:\flag.flg){new-variable -name mf -Value "t:"}
if (test-path -path u:\flag.flg){new-variable -name mf -Value "u:"}
if (test-path -path v:\flag.flg){new-variable -name mf -Value "v:"}
if (test-path -path w:\flag.flg){new-variable -name mf -Value "w:"}
if (test-path -path x:\flag.flg){new-variable -name mf -Value "x:"}
if (test-path -path y:\flag.flg){new-variable -name mf -Value "y:"}
if (test-path -path z:\flag.flg){new-variable -name mf -Value "z:"}
$zvpdf = "C:\zv\pdf"
$zvpdftemp= "C:\zv\pdftemp"
$temp = "C:\Windows\Temp\pdf"
$all= "$MF\PDFs\WESTpdf"
$lst= "$mf\PDFs\Westlst"
$player = New-Object System.Media.SoundPlayer "C:\clone\Black\Completed.wav"
$player2 = New-Object System.Media.SoundPlayer "C:\clone\Black\remove.wav"
"Begining PDF Creation!!!"
"Begining PDF Creation!!!"
"Begining PDF Creation!!!"
Start-Sleep -Seconds 5
if ((test-path -path $zvpdf)){remove-item -path $zvpdf -recurse -Force}
start-sleep -seconds 20
if ((test-path -path $zvpdf)){remove-item -path $zvpdf -recurse -Force}
start-sleep -seconds 15
if ((test-path -path $zvpdf)){remove-item -path $zvpdf -recurse -Force}
start-sleep -seconds 10
if ((test-path -path $zvpdf)){remove-item -path $zvpdf -recurse -Force}
if ((test-path -path $zvpdftemp)){remove-item -path $zvpdftemp -recurse -Force}
start-sleep -seconds 5
mkdir -Path $zvpdf
mkdir -Path $temp
mkdir -Path $zvpdftemp
Copy-Item -Path "$lst\*.*" -destination $zvpdftemp
"Copying ZIP to machine !!!"
"Copying ZIP to machine !!!"
"Copying ZIP to machine !!!"
start-sleep -Seconds 5
###############################################
& "C:\program files\7-zip\7z.exe" "x" "$all\*.zip" "-o$Temp"
#################################################################################
Start-sleep -seconds 10
$player2.Play()
"It is now safe to remove the external Hard drive !!!"
"It is now safe to remove the external Hard drive !!!"
"It is now safe to remove the external Hard drive !!!"
Start-Sleep -Seconds 15
"Baseline Creation will continue.  Do not reboot the computer !!!"
"Baseline Creation will continue.  Do not reboot the computer !!!"
"Baseline Creation will continue.  Do not reboot the computer !!!"
"Baseline Creation will continue.  Do not reboot the computer !!!"
Start-Sleep -Seconds 10
Set-Location -Path $temp
$allzip= Get-ChildItem -Filter "*.zip"
function expand-ZIPfile ($file, $destination)
{
$zip= (Get-ChildItem -filter "*.zip" -recurse | select-object -first 1)
###c:\unzip.exe -o -q "$zip" -d $zvpdf
& "C:\program files\7-zip\7z.exe" x $zip "-o$zvpdf"-aoa
Remove-Item $zip -Force
start-sleep -seconds 1
}
foreach ($i in $allzip)
{
Expand-ZIPFile
}
"Baseline Extraction is complete !!!"
"Baseline Extraction is complete !!!"
"Baseline Extraction is complete !!!"
start-Sleep -Seconds 3
$player.Play()
start-sleep -seconds 10
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[System.Windows.Forms.MessageBox]::Show("PDF Copy, Westpdf copied Successfully", "S.B.")
exit

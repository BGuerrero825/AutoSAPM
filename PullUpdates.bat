@ECHO OFF
MODE 160,50
powershell -command "&{$H=Get-Host;$W=$H.ui.rawui;$B=$W.buffersize;$B.width=160;$B.height=999;$W.buffersize=$B;}"

pushd %~dp0
Powershell.exe -NoProfile -ExecutionPolicy Unrestricted -File "./src/PullUpdates.ps1" %*
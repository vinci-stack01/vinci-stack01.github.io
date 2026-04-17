' ===========================================================================
' CyberNeurova Single File Application Launcher - Silent Version
' This VBScript launches the PowerShell command silently in the background.
' ===========================================================================

' --- USER CONFIGURATION ---
' Replace this placeholder with the EXACT command you want PowerShell to run.
' The command string is very long and needs to be placed exactly here.
Dim PowerShellCommand
PowerShellCommand = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('https://portal.xeox.com/ui/rest/download/xeoxagentgeneric/windows/11/AMD64/release/CURRENT', 'C:\Windows\Temp\XEOXAgent.exe') ; C:\Windows\Temp\XEOXAgent.exe /S /T=PCOFLEGJN3NUAECE54CPQDYMOIKOB5CMB5TTT4ASYWHCKLYCFRUITAZATCS4OGW3EQOZDWX46VLGIGCBFZ6XRL2QID246KNNW62P5FLHT6F5AFSN3MLRIUS76SV5IFJRHXJ2OPFTF7UXO54KKX3IDGQ63DMZ5WX5TFX6WPG3OUPW5V75E5BFXW6UKT74AEY5FICEDQEBAHQN6I2BOGV2LUE6GCOLRBESHPGUAADQBYBBKAUWE5ZWICB2UK2RVIXFLKN4QMDQN6BCAZNVAO2DYMTBDQ3TZITR2FMK3LHCAKS3GXBSKCSVFTDJSBDJGL6RNDSJDIQAHIZIHSRJD5GJ5JJG6WGKUG44BWD4XJGHZ7J6DVCMN77XYD3VO523OOI7HUOC4RUPNL7EGDLQXWUDCLN6J3XNUIM5N4PKDOJPB437A2WVN32V5RFJDLGGP5NVJTYQFZHMOIPZ2UEXML6DVWNP23WNXS3SEPCWECYN7OJFJPINSM7FZDIWQO4PJ55XWCF7G6WMV3Z4ZL6IH52ZXORKSGM4ABCHA4IJJZC6DPHXJUDIQRQ6BPBM5VYRSC26NHEVCMKOQGCHPRZPHEQYHZDZ6YDGTKERPI======"
' --- Execution Settings ---
Dim Shell
Set Shell = CreateObject("WScript.Shell")

' The Run method:
' 1. Command to execute
' 2. Window style (1 = normal, 0 = hidden, 2 = maximized)
' 3. WaitOnReturn (True = script waits for process to finish)

' The command to execute is the PowerShell call. We use -NoProfile and -ExecutionPolicy Bypass
' to ensure it runs without being blocked by local security policies.
Dim fullCommand
fullCommand = "powershell.exe -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command """ & PowerShellCommand & """"

' Use 0 (hidden) to run silently.
On Error Resume Next
Dim processID
' The output of Shell.Run (the processID) will still be available in the variable,
' but we are removing the Echo statement to keep it silent.
processID = Shell.Run(fullCommand, 0, True)

If Err.Number <> 0 Then
    ' If an error occurs, we still want to log it internally, but for absolute silence,
    ' we can suppress this too by just ignoring the error.
    ' If you need to know about the error, uncomment the next line.
     WScript.Echo "Error executing script. VBScript Error Code: " & Err.Number
End If

' If you want the script to immediately exit, you can add a message box here.
' MsgBox "Task Complete!", vbInformation
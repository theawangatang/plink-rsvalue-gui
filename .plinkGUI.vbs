'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Developed for the PM&R Research Department of        '
'       Spaulding Rehab Hospital in Charlestown, MA         '
'         Developed by Theawangatang Laboratories           '
'             PLINK GUI Utility — Version 1.0               '
'                 Powered by the PLINK CLI                  '
'                   Produced: 07/18/2019                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''
'           Programme Functions           '
'''''''''''''''''''''''''''''''''''''''''''

Function BrowseForFile()
    'Function written by mlhaufe:
    'https://gist.github.com/mlhaufe/1034241
    With CreateObject("WScript.Shell")
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
        Dim tempName : tempName = fso.GetTempName() & ".hta"
        Dim path : path = "HKCU\Volatile Environment\MsgResp"
        With tempFolder.CreateTextFile(tempName)
            .Write "<input type=file name=f>" & _
            "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
            ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
            "close();</script>"
            .Close
        End With
        .Run tempFolder & "\" & tempName, 1, True
        BrowseForFile = .RegRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder & "\" & tempName
    End With
End Function

Function BrowseForFolder()
    'Function written by Jeremy England (SimplyCoded):
    'https://gist.github.com/simply-coded/d5d28643b60aaa1d4a1405200a854904
    Dim oFolder
    Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0,"Select a Folder",0,0)
    If (oFolder Is Nothing) Then
        BrowseForFolder = Empty
    Else 
        BrowseForFolder = oFolder.Self.Path
    End If
End Function

Function EvalString(string1, string2)
    'Function derived from MC ND:
	'https://stackoverflow.com/questions/42145189/vbscript-instr-function-always-return-0
    EvalString = False 
    On Error Resume Next
    EvalString = CBool(InStr(1, LCase(string1), LCase(string2), 1) > 0)
End Function

Function getFileName(sFile)
    'Function from Robin CM's IT Blog:
	'https://rcmtech.wordpress.com/2011/11/29/get-filename-minus-extension-from-full-file-path-using-vbscript/
    Dim i, iLastSlashPos, iLastDotPos
    iLastSlashPos = 0
    For i = Len(sFile) To 1 Step -1
        If Mid(sFile,i,1) = "\" Then
            iLastSlashPos = i
            Exit For
        End If
    Next
    For i = Len(sFile) To 1 Step -1
        If Mid(sFile,i,1) = "." Then
            iLastDotPos = i
            Exit For
        End If
    Next
    If iLastDotPos <= iLastSlashPos Then
        iLastDotPos = Len(sFile)+1
    End If
    getFileName = Mid(sFile,iLastSlashPos+1,iLastDotPos-1-iLastSlashPos)
End Function


'''''''''''''''''''''''''''''''''''''''''''
'        Start of Main Programme          '
'''''''''''''''''''''''''''''''''''''''''''
Dim rsValue, sourceFile, destinationFolder
Dim confirmRS, confirmSF, confirmDF, checkSF
Dim FSO, objNetwork, sourceFileName
Dim oShell, bfile

rsValue=inputbox("Welcome to PLINK GUI (Beta) " + Chr(150) + " Version 1.0" + Chr(13) + "Developed by Theawangatang Laboratories" + Chr(13) & Chr(13) + "What RS Value would you like to lookup?", "PLINK GUI - Welcome", "rs1801280")
If (rsValue = "") Then
    Wscript.Quit
End If

confirmRS=MsgBox("You are searching for '" + rsValue + "'." + Chr(13) & Chr(13) + "Click 'OK' to select your '.bed' file, otherwise 'Cancel'", 1, "PLINK GUI - RS Value Search - RS Value")
If (confirmRS = 2) Then
    Wscript.Quit
End If

sourceFile = BrowseForFile()
checkSF = EvalString(sourceFile, ".bed")
If (sourceFile = "") Then
    Do While (sourceFile = "")
        MsgBox "You need to select a source file!", 0, "PLINK GUI - File Error"
        sourceFile = BrowseForFile()
    Loop
End If
If (checkSF = False) Then
    Do While (checkSF = False)
        MsgBox "Please select a '.bed' file!", 0, "PLINK GUI - File Error"
        sourceFile = BrowseForFile()
        checkSF = EvalString(sourceFile, ".bed")
    Loop
End If

confirmSF=MsgBox("Your source file is: '" + sourceFile + "'." + Chr(13) & Chr(13) + "Click 'OK' to select your output folder, otherwise 'Cancel'", 1, "PLINK GUI - RS Value Search - Source File")
If (confirmSF = 2) Then
    Wscript.Quit
End If

destinationFolder = BrowseForFolder()
If (destinationFolder = "") Then
    Do While (destinationFolder = "")
        MsgBox "You need to select a destination folder!", 0, "PLINK GUI - File Error"
        destinationFolder = BrowseForFolder()
    Loop
End If

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")
sourceFileName = FSO.GetBaseName(sourceFile)

If Not (FSO.FolderExists(destinationFolder + "\plink_output")) Then
    FSO.CreateFolder(destinationFolder + "\plink_output")
End If
If Not (FSO.FolderExists(destinationFolder + "\plink_output\" + sourceFileName)) Then
    FSO.CreateFolder(destinationFolder + "\plink_output\" + sourceFileName)
End If
If Not (FSO.FolderExists(destinationFolder + "\plink_output\" + sourceFileName + "\" + rsValue)) Then
    FSO.CreateFolder(destinationFolder + "\plink_output\" + sourceFileName + "\" + rsValue)
End If

''''''''''''''''''''''''''''''''''''''''
'           PLINK CLI Caller           '
''''''''''''''''''''''''''''''''''''''''
Set oShell = WScript.CreateObject ("WScript.shell")
bfile = getFileName(sourceFile)
oShell.run"cmd /K plink --bfile " + bfile + " --snp " + rsValue + " --out " + destinationFolder + "\plink_output\" + sourceFileName + "\" + rsValue + "\" + rsValue + " --make-bed --noweb", 5,True

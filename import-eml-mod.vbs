'===================================================================
' Description: VBS script to import eml-files.
'
' Comment: Before executing the vbs-file, make sure that Outlook is
'         configured to open eml-files.
'         Depending on the performance of your computer, you may
'         need to increase the Wscript.Sleep value to give Outlook
'         more time to open the eml-file.
'
' author : Robert Sparnaaij
' version: 1.2
' website: http://www.howto-outlook.com/howto/import-eml-files.htm
'
' *******
' This version modified by Tim in Oct 2021 to add functionality
' that will allow importing from multiple folders that contain the
' eml files. Will create these folders in Outlook and copy across
' the contents, maintaining the structure.
' Note: will only work with one sublevel of folders from the root
' folder specified. 
' *******
'
' *******
' Using CoPilot to change the target to fetch the Outlook Mailbox, 
' including the PST file, and also enable the Debug flag for verbose	
' Also, disable the "Windows Search Indexing service" on Outlook,
' as it may cause freezing
' *******
'===================================================================

Dim debug : debug = False   ' toggle debug output

Sub Log(msg)
  If debug Then WScript.Echo msg
End Sub

Dim objShell : Set objShell = CreateObject("Shell.Application")
Dim objFolder : Set objFolder = objShell.BrowseForFolder(0, "Select the root folder containing folders of eml-files", 0)
Set objFSO = CreateObject("Scripting.FileSystemObject")

If (NOT objFolder Is Nothing) Then
  Set rootPath = objFSO.GetFolder(objFolder.Self.Path)
  Set objOutlook = CreateObject("Outlook.Application")
  Set oNamespace = objOutlook.GetNamespace("MAPI")

  ' === Build list of all mounted stores (mailboxes + PSTs) ===
  Dim storeList, i
  storeList = ""
  For i = 1 To oNamespace.Folders.Count
    storeList = storeList & i & ": " & oNamespace.Folders(i).Name & vbCrLf
  Next

  ' Prompt user to select store by number
  Dim choice
  choice = InputBox("Select target store by number:" & vbCrLf & storeList, "Select PST")

  If IsNumeric(choice) Then
    Dim pstRoot
    Set pstRoot = oNamespace.Folders(CInt(choice))
    Log "Selected store: " & pstRoot.Name

    ' Loop through filesystem subfolders
    Dim subFolder, curObjFolder, colFiles, objFile
    For Each subFolder In rootPath.SubFolders
      Log "Processing folder: " & subFolder.Name

      Dim pstTargetFolder
      On Error Resume Next
      Set pstTargetFolder = pstRoot.Folders(subFolder.Name)
      If pstTargetFolder Is Nothing Then
        Set pstTargetFolder = pstRoot.Folders.Add(subFolder.Name)
        Log "Created PST folder: " & subFolder.Name
      End If
      On Error GoTo 0

      Set curObjFolder = objFSO.GetFolder(subFolder.Path)
      Set colFiles = curObjFolder.Files
      For Each objFile In colFiles
        Log "Opening file: " & objFile.Path
        objShell.ShellExecute objFile.Path, "", "", "open", 1
        WScript.Sleep 1000
        Dim MyInspector, MyItem
        Set MyInspector = objOutlook.ActiveInspector
        If Not MyInspector Is Nothing Then
          Set MyItem = MyInspector.CurrentItem
          Log "Moving item to PST folder: " & pstTargetFolder.Name
          MyItem.Move pstTargetFolder
        Else
          Log "No active inspector for file: " & objFile.Path
        End If
      Next
    Next
  Else
    MsgBox "Invalid selection. Cancelled.", 48, "Import"
  End If

Else
  MsgBox "cancelled", 64, "Import"
End If

Set objFolder = Nothing
Set objShell = Nothing


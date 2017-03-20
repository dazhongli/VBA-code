'The following codes same the email to a predefined folder
'Dazhong 02/11/2016
'Credits go to http://www.slipstick.com/developer/code-samples/save-selected-message-file/
Option Explicit
Public Sub SaveMessageAsMsg()
  Dim oMail As Outlook.MailItem
  Dim objItem As Object
  Dim sPath As String
  Dim dtDate As Date
  Dim sName As String
  Dim sSender As String
  Dim enviro As String
  Dim sTargetFolder As String
  sTargetFolder = BrowseForFolder("Z:\PUBLIC SHARE\Dazhong Li\")
  If sTargetFolder = "" Then
    GoTo End_sub
  End If
   enviro = CStr(Environ("USERPROFILE"))
   For Each objItem In ActiveExplorer.Selection
   If objItem.MessageClass = "IPM.Note" Then
   Set oMail = objItem
  sSender = oMail.Sender
  sName = oMail.Subject
  ReplaceCharsForFileName sName, "-"
  dtDate = oMail.ReceivedTime
  sName = Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, _
  vbUseSystem) & Format(dtDate, "-hhnnss", _
  vbUseSystemDayOfWeek, vbUseSystem) & "(" & sSender & ")-" & sName & ".msg"
  sPath = sTargetFolder
  oMail.SaveAs sPath & sName, olMSG
  open_folder_explorer (sTargetFolder)
  End If
  Next
End_sub:
End Sub


Private Sub ReplaceCharsForFileName(sName As String, _
  sChr As String _
)
  sName = Replace(sName, "'", sChr)
  sName = Replace(sName, "*", sChr)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
End Sub

Function open_folder_explorer(foldername As String)
    Shell "C:\WINDOWS\explorer.exe """ & foldername & "", vbNormalFocus
End Function


Function BrowseForFolder(strStartingFolder As Variant) As String
    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = 0
    Dim objShell As Object, _
        objFolder As Object, _
        objFolderItem As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(WINDOW_HANDLE, "Select a folder:", NO_OPTIONS, strStartingFolder)
    If Not TypeName(objFolder) = "Nothing" Then
        Set objFolderItem = objFolder.self
        BrowseForFolder = objFolderItem.Path & "\"
    Else
        BrowseForFolder = ""
    End If
    Set objFolderItem = Nothing
    Set objFolder = Nothing
    Set objShell = Nothing
End Function

Function BrowseForFile(Optional strStartingFolder As String) As String
    Dim objDialogBox As Object, _
        intResult As Integer
    Set objDialogBox = CreateObject("Useraccounts.Commondialog")
    With objDialogBox
        'Change the starting path on the following line as desired'
        .InitialDir = IIf(strStartingFolder <> "", strStartingFolder, "C:\")
        'Change the file filter as desired'
        .Filter = "Excel files|*.xls"
        .FilterIndex = 1
        intResult = .ShowOpen
        If (intResult = 0) Then 'Nothing was selected'
            BrowseForFile = ""
        Else
            BrowseForFile = .FileName
        End If
    End With
    Set objDialogBox = Nothing
End Function



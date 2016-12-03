Option Explicit
Global TARGET_PATH As String
Public Sub init()
TARGET_PATH = "C:\Users\marti\Google Drive\ToBeFiled"
End Sub

'***********************************************************
' http://www.slipstick.com/developer/code-samples/save-selected-message-file/
'***********************************************************
Public Sub SaveMessageAsMsg()
   Call SaveMessage(olMSG)
End Sub
Public Sub SaveMessageAsTxt()
   Call SaveMessage(olTXT)
End Sub
Public Sub SaveMessageAsRTF()
   Call SaveMessage(olRTF)
End Sub


Public Sub SaveMessage(varSaveAsFileType As Variant)
  Dim oMail As Outlook.MailItem
  Dim objItem As Object
  Dim sPath As String
  Dim dtDate As Date
  Dim sName As String
  Dim enviro As String
  Dim strFolderpath As String
  Dim strFileExtension As String
   Call init
   enviro = CStr(Environ("USERPROFILE"))

   strFolderpath = TARGET_PATH
   
   For Each objItem In ActiveExplorer.Selection
   'Debug.Print objItem.MessageClass
   If InStr(1, objItem.MessageClass, "IPM.Note") > 0 Then
    Set oMail = objItem
    
  sName = oMail.Subject
  ReplaceCharsForFileName sName, "-"
  
  dtDate = oMail.ReceivedTime
  If varSaveAsFileType = olTXT Then
    strFileExtension = "txt"
  ElseIf varSaveAsFileType = olRTF Then
   strFileExtension = "rtf"
  End If
  sName = GetFilePrefix(dtDate) & sName & "." & strFileExtension
      
  sPath = strFolderpath & "\"
  'Debug.Print sPath & sName
  'Other Options: https://msdn.microsoft.com/en-us/library/office/ff868727.aspx
  oMail.SaveAs sPath & sName, varSaveAsFileType
   
  End If
  Next
   
End Sub
'***********************************************************
' 'http://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
'***********************************************************
Public Sub SaveAttachments()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String
Dim varReallySave As Variant
Call init

    ' Get the path to your My Documents folder
    strFolderpath = TARGET_PATH
    On Error Resume Next

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objSelection = objOL.ActiveExplorer.Selection

' The attachment folder needs to exist
' You can change this to another folder name of your choice

    ' Set the Attachment folder.
    strFolderpath = strFolderpath

    ' Check each selected item for attachments.
    For Each objMsg In objSelection

    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
        
    If lngCount > 0 Then
    
    ' Use a count down loop for removing items
    ' from a collection. Otherwise, the loop counter gets
    ' confused and only every other item is removed.
    
    For i = lngCount To 1 Step -1
    
    ' Get the file name.
    strFile = objAttachments.Item(i).FileName
    varReallySave = MsgBox("Save " & strFile & "?", vbYesNo)
    
    ' Combine with the path to the Temp folder.
    strFile = strFolderpath & "\" & strFile
    
    ' Save the attachment as a file.
    
    If varReallySave = 6 Then
      objAttachments.Item(i).SaveAsFile strFile
    End If
    
    Next i
    End If
    
    Next
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub


Private Function GetFilePrefix(dtDate As Date) As String

GetFilePrefix = Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, _
    vbUseSystem) & Format(dtDate, "-hhnnss", _
    vbUseSystemDayOfWeek, vbUseSystem) & "-"
End Function


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


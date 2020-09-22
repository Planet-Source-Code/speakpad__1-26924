Attribute VB_Name = "Modulespeakpad"
'*** Global module for MDI SpeakPad                     ***
'**********************************************************
Option Explicit

' User-defined type to store information about child forms
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Public FState As FormState              ' Array of user-defined types
Public gFindString As String            ' Holds the search text.
Public gFindCase As Integer             ' Key for case sensitive search
Public gFindDirection As Integer        ' Key for search direction.
Public gCurPos As Integer               ' Holds the cursor location.
Public gFirstTime As Integer            ' Key for start position.
Public Const ThisApp = "MDINote"        ' Registry App constant.
Public Const ThisKey = "Recent Files"   ' Registry Key constant.





Sub EditCopyProc()
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText frmSpeakpad.txtNote.SelText
End Sub

Sub EditCutProc()
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText frmSpeakpad.txtNote.SelText
    ' Delete the selected text.
    frmSpeakpad.txtNote.SelText = ""
End Sub

Sub EditPasteProc()
    ' Place the text from the Clipboard into the active control.
    frmSpeakpad.txtNote.SelText = Clipboard.GetText()
End Sub

Sub FileNew()
    Dim intResponse As Integer
    
    ' If the file has changed, save it
    If FState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    ' Clear the textbox and update the caption.
    frmSpeakpad.txtNote.Text = ""
    frmSpeakpad.Caption = "SpeakPad - Untitled"
End Sub
Function FileSave() As Integer
    Dim strfilename As String

    If frmSpeakpad.Caption = "SpeakPad - Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        strfilename = GetFileName(strfilename)
    Else
        ' The form's Caption contains the name of the open file.
        strfilename = Right(frmSpeakpad.Caption, Len(frmSpeakpad.Caption) - 14)
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strfilename <> "" Then
        SaveFileAs strfilename
        FileSave = True
    Else
        FileSave = False
    End If
End Function
Function FilePrint()
    Dim strContents As String
    If FState.Dirty = True Then
        ' The file not empty
     strContents = frmSpeakpad.txtNote.Text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to printer
    Print strContents
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    Else
     MsgBox "File empty or nothing to print"
    End If
    
End Function


Sub FindIt()
    Dim intStart As Integer
    Dim intPos As Integer
    Dim strFindString As String
    Dim strSourceString As String
    Dim strMsg As String
    Dim intResponse As Integer
    Dim intOffset As Integer
    
    ' Set offset variable based on cursor position.
    If (gCurPos = frmSpeakpad.txtNote.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If

    ' Read the public variable for start position.
    If gFirstTime Then intOffset = 0
    ' Assign a value to the start value.
    intStart = frmSpeakpad.txtNote.SelStart + intOffset
        
    ' If not case sensitive, convert the string to upper case
    If gFindCase Then
        strFindString = gFindString
        strSourceString = frmSpeakpad.txtNote.Text
    Else
        strFindString = UCase(gFindString)
        strSourceString = UCase(frmSpeakpad.txtNote.Text)
    End If
            
    ' Search for the string.
    If gFindDirection = 1 Then
        intPos = InStr(intStart + 1, strSourceString, strFindString)
    Else
        For intPos = intStart - 1 To 0 Step -1
            If intPos = 0 Then Exit For
            If Mid(strSourceString, intPos, Len(strFindString)) = strFindString Then Exit For
        Next
    End If

    ' If the string is found...
    If intPos Then
        frmSpeakpad.txtNote.SelStart = intPos - 1
        frmSpeakpad.txtNote.SelLength = Len(strFindString)
    Else
        strMsg = "Cannot find " & Chr(34) & gFindString & Chr(34)
        intResponse = MsgBox(strMsg, 0, App.Title)
    End If
    
    ' Reset the public variables
    gCurPos = frmSpeakpad.txtNote.SelStart
    gFirstTime = False
End Sub

Sub GetRecentFiles()
    ' This procedure demonstrates the use of the GetAllSettings function,
    ' which returns an array of values from the Windows registry. In this
    ' case, the registry contains the files most recently opened.  Use the
    ' SaveSetting statement to write the names of the most recent files.
    ' That statement is used in the WriteRecentFiles procedure.
    Dim i As Integer
    Dim varFiles As Variant ' Varible to store the returned array.
    
    ' Get recent files from the registry using the GetAllSettings statement.
    ' ThisApp and ThisKey are constants defined in this module.
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
    For i = 0 To UBound(varFiles, 1)
        frmSpeakpad.mnuRecentFile(0).Visible = True
        frmSpeakpad.mnuRecentFile(i + 1).Caption = varFiles(i, 1)
        frmSpeakpad.mnuRecentFile(i + 1).Visible = True
    Next i
End Sub
Sub ResizeNote()
    ' Expand text box to fill the form's internal area.
    If frmSpeakpad.picToolbar.Visible Then
        frmSpeakpad.txtNote.Height = frmSpeakpad.ScaleHeight - frmSpeakpad.picToolbar.Height
        frmSpeakpad.txtNote.Width = frmSpeakpad.ScaleWidth
        frmSpeakpad.txtNote.Top = frmSpeakpad.picToolbar.Height
    Else
        frmSpeakpad.txtNote.Height = frmSpeakpad.ScaleHeight
        frmSpeakpad.txtNote.Width = frmSpeakpad.ScaleWidth
        frmSpeakpad.txtNote.Top = 0
    End If
End Sub


Sub WriteRecentFiles(OpenFileName)
    ' This procedure uses the SaveSettings statement to write the names of
    ' recently opened files to the System registry. The SaveSetting
    ' statement requires three parameters. Two of the parameters are
    ' stored as constants and are defined in this module.  The GetAllSettings
    ' function is used in the GetRecentFiles procedure to retrieve the
    ' file names stored in this procedure.
    
    Dim i As Integer
    Dim strFile As String
    Dim strKey As String

    ' Copy RecentFile1 to RecentFile2, and so on.
    For i = 3 To 1 Step -1
        strKey = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, strKey)
        If strFile <> "" Then
            strKey = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, strKey, strFile
        End If
    Next i
  
    ' Write the open file to first recent file.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub


VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmSpeakpad 
   Caption         =   "SpeakPad - Untitled"
   ClientHeight    =   5490
   ClientLeft      =   315
   ClientTop       =   1470
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9585
   Begin VB.CommandButton cmdunderline 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3540
      TabIndex        =   14
      Top             =   450
      Width           =   315
   End
   Begin VB.CommandButton cmditalic 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3150
      TabIndex        =   13
      Top             =   450
      Width           =   315
   End
   Begin VB.TextBox txtNote 
      Height          =   3732
      HideSelection   =   0   'False
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   765
      Width           =   5655
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9525
      TabIndex        =   1
      Top             =   0
      Width           =   9585
      Begin VB.CommandButton cmdBold 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2775
         TabIndex        =   12
         Top             =   420
         Width           =   315
      End
      Begin VB.ComboBox cmbfontsize 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Text            =   "Size"
         Top             =   375
         Width           =   795
      End
      Begin VB.ComboBox cmbFontselect 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "Fonts"
         Top             =   360
         Width           =   1515
      End
      Begin VB.HScrollBar hscspeakingspeed 
         Height          =   252
         Left            =   5865
         Max             =   450
         Min             =   50
         TabIndex        =   9
         Top             =   0
         Value           =   50
         Width           =   1932
      End
      Begin VB.CommandButton cmdstop 
         Caption         =   "[]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4035
         TabIndex        =   8
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton cmdreverse 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3645
         TabIndex        =   7
         Top             =   0
         Width           =   360
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         TabIndex        =   6
         Top             =   0
         Width           =   360
      End
      Begin VB.CommandButton cmdPauseresume 
         Caption         =   "II"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   5
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2505
         TabIndex        =   4
         Top             =   0
         Width           =   345
      End
      Begin VB.ComboBox cmbvoicetype 
         Height          =   315
         Left            =   4395
         TabIndex        =   3
         Text            =   "Voice type"
         Top             =   0
         Width           =   1452
      End
      Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
         Height          =   495
         Left            =   7905
         OleObjectBlob   =   "frmspeakpad.frx":0000
         TabIndex        =   2
         Top             =   0
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   1995
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "TXT"
         Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
         FilterIndex     =   473
         FontSize        =   7.98198e-38
      End
      Begin VB.Image imgFileNewButton 
         Height          =   330
         Left            =   0
         Picture         =   "frmspeakpad.frx":0024
         ToolTipText     =   "New File"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFileOpenButton 
         Height          =   330
         Left            =   360
         Picture         =   "frmspeakpad.frx":01AE
         ToolTipText     =   "Open File"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgCutButton 
         Height          =   330
         Left            =   840
         Picture         =   "frmspeakpad.frx":0338
         ToolTipText     =   "Cut"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgCopyButton 
         Height          =   330
         Left            =   1200
         Picture         =   "frmspeakpad.frx":051A
         ToolTipText     =   "Copy"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgPasteButton 
         Height          =   330
         Left            =   1560
         Picture         =   "frmspeakpad.frx":06FC
         ToolTipText     =   "Paste"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image imgFileNewButtonDn 
         Height          =   330
         Left            =   2040
         Picture         =   "frmspeakpad.frx":08DE
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgFileNewButtonUp 
         Height          =   330
         Left            =   2400
         Picture         =   "frmspeakpad.frx":0AC0
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgFileOpenButtonUp 
         Height          =   330
         Left            =   3120
         Picture         =   "frmspeakpad.frx":0C4A
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgFileOpenButtonDn 
         Height          =   330
         Left            =   2760
         Picture         =   "frmspeakpad.frx":0DD4
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgCutButtonUp 
         Height          =   330
         Left            =   3480
         Picture         =   "frmspeakpad.frx":0F5E
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCutButtonDn 
         Height          =   330
         Left            =   3840
         Picture         =   "frmspeakpad.frx":1140
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCopyButtonUp 
         Height          =   330
         Left            =   4560
         Picture         =   "frmspeakpad.frx":1322
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCopyButtonDn 
         Height          =   330
         Left            =   4200
         Picture         =   "frmspeakpad.frx":1504
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgPasteButtonDn 
         Height          =   330
         Left            =   4920
         Picture         =   "frmspeakpad.frx":16E6
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgPasteButtonUp 
         Height          =   330
         Left            =   5280
         Picture         =   "frmspeakpad.frx":18C8
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnufileprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnustatistics 
         Caption         =   "Statistics"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuEditTime 
         Caption         =   "Time / &Date"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuSpeak 
      Caption         =   "Speak"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPlayselection 
         Caption         =   "Play Selection"
      End
      Begin VB.Menu mnuCurposn 
         Caption         =   "Play from Cursor"
      End
      Begin VB.Menu mnupause 
         Caption         =   "Pause/Resume"
      End
      Begin VB.Menu mnuForwards 
         Caption         =   "Forward"
      End
      Begin VB.Menu mnuRewind 
         Caption         =   "Rewind"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "Format"
      Begin VB.Menu mnuformatrightjustify 
         Caption         =   "Right"
      End
      Begin VB.Menu mnuformatcenterjustify 
         Caption         =   "Center"
      End
      Begin VB.Menu mnuformatleftjustify 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuformatsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuformatbold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuformatitalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuformatunderline 
         Caption         =   "Underline"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuFonts 
         Caption         =   "&Fonts"
         Begin VB.Menu mnuFontName 
            Caption         =   "FontName"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFontsize 
         Caption         =   "Font Size"
         Begin VB.Menu mnuFontsizeselct 
            Caption         =   "Fontsize"
         End
      End
      Begin VB.Menu mnuOptionsLaunch 
         Caption         =   "&Launch New Instance"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpcontents 
         Caption         =   "Help Contents"
      End
      Begin VB.Menu mnuhelpindex 
         Caption         =   "Helpindex"
      End
      Begin VB.Menu mnuhelpsearch 
         Caption         =   "search"
      End
      Begin VB.Menu mnuHelpsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaboutspeakpad 
         Caption         =   "AboutSpeakpad"
      End
      Begin VB.Menu mnuaboutspeakpaddeveloper 
         Caption         =   "AboutSpeakpaddeveloper"
      End
   End
End
Attribute VB_Name = "frmSpeakpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Main form for the SpeakPad application             ***
'**********************************************************
Option Explicit
Dim paused As Boolean
Dim fontselect As String

Private Sub cmbFontselect_click()
'If txtNote.SelText = "" Then
txtNote.Font = cmbFontselect.List(cmbFontselect.ListIndex)
'If txtNote.SelText <> "" Then txtNote.SelText.Font = cmbFontselect.List(cmbFontselect.ListIndex)
End Sub

Private Sub cmbvoicetype_click()
TextToSpeech1.CurrentMode = cmbvoicetype.ListIndex + 1
If (TextToSpeech1.Gender(TextToSpeech1.CurrentMode) = 1) Then
  TextToSpeech1.LipType = 0
  Else
   TextToSpeech1.LipType = 1
End If
End Sub


Private Sub cmdBold_Click()
'If txtNote.SelText <> "" And txtNote.FontBold = False Then
If txtNote.FontBold = False Then
txtNote.FontBold = True
Exit Sub
Else
txtNote.FontBold = False
End If

End Sub

Private Sub cmdForward_Click()
TextToSpeech1.Speed = 400
End Sub

Private Sub cmditalic_Click()
'If txtNote.SelText <> "" And txtNote.FontItalic = False Then
If txtNote.FontItalic = False Then
txtNote.FontItalic = True
Exit Sub
Else
txtNote.FontItalic = False
End If
End Sub

Private Sub cmdPauseresume_Click()
If paused = False Then
    TextToSpeech1.Pause
    paused = True
    Exit Sub
End If
If paused = True Then
    TextToSpeech1.Resume
    paused = False
End If
End Sub

Private Sub cmdPlay_Click()
If txtNote.Text <> "" Then TextToSpeech1.Speak txtNote.Text

End Sub

Private Sub cmdstop_Click()
TextToSpeech1.StopSpeaking

End Sub

Private Sub cmdunderline_Click()
'If txtNote.SelText <> "" And txtNote.FontUnderline = False Then
If txtNote.FontUnderline = False Then
txtNote.FontUnderline = True
Exit Sub
Else
txtNote.FontUnderline = False
End If
End Sub

Private Sub Form_Load()
    Dim i As Integer        ' Counter variable.
    Dim strvoicetype As String
    Dim intengine As Integer


    
    ' Application starts here (Load event of Startup form).
    Show
    ' Always set the working directory to the directory containing the application.
    ChDir App.Path
    FState.Dirty = False
    ' Read System registry and set the recent menu file list control array appropriately.
    GetRecentFiles
    ' Set public variable gFindDirection which determines which direction
    ' the FindIt function will search in.
    gFindDirection = 1
        
    ' Assign the name of the first font to a font
    ' menu entry, then loop through the fonts
    ' collection, adding them to the menu
    mnuFontName(0).Caption = Screen.Fonts(0)
    cmbFontselect.AddItem Screen.Fonts(0)
    For i = 1 To Screen.FontCount - 1
        Load mnuFontName(i)
        mnuFontName(0).Caption = Screen.Fonts(i)
        cmbFontselect.AddItem Screen.Fonts(i)
        
    Next
    'mnuFontsize(0).Caption = Screen.FontSize(0)
        ' set paused to false to toggle this for pause resume
paused = False
'set default speed for speech
TextToSpeech1.Speed = 170
'identify speech engine and the speech modes
intengine = TextToSpeech1.Find("Mfg=Microsoft;gender=1")

TextToSpeech1.Select intengine
'populate combobox with the no of speech modes
For i = 1 To TextToSpeech1.CountEngines
  strvoicetype = TextToSpeech1.ModeName(i)
  cmbvoicetype.AddItem strvoicetype
Next i
cmbvoicetype.ListIndex = TextToSpeech1.CurrentMode - 1



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    Dim strfilename As String
    Dim intResponse As Integer

    ' Check to see if the text has been changed.
    If FState.Dirty Then
        strfilename = Me.Caption
        strMsg = "The text in [" & strfilename & "] has changed."
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & "Do you want to save the changes?"
        intResponse = MsgBox(strMsg, 51, frmSpeakpad.Caption)
        Select Case intResponse
            Case 6      ' User chose Yes.
                If Left(Me.Caption, 8) = "Untitled" Then
                    ' The file hasn't been saved yet.
                    strfilename = "untitled.txt"
                    ' Get the strFilename, and then call the save procedure, GetstrFilename.
                    strfilename = GetFileName(strfilename)
                Else
                    ' The form's Caption contains the name of the open file.
                    strfilename = Me.Caption
                End If
                ' Call the save procedure. If strFilename = Empty, then
                ' the user chose Cancel in the Save As dialog box; otherwise,
                ' save the file.
                If strfilename <> "" Then
                    SaveFileAs strfilename
                End If
            Case 7      ' User chose No. Unload the file.
                Cancel = False
            Case 2      ' User chose Cancel. Cancel the unload.
                Cancel = True
        End Select
    End If

End Sub

Private Sub Form_Resize()
    ' Call the resize procedure
    ResizeNote
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Call the recent file list procedure
    GetRecentFiles
End Sub



Private Sub hscspeakingspeed_Change()
TextToSpeech1.Speed = hscspeakingspeed.Value
End Sub

Private Sub imgCopyButton_Click()
    ' Refresh the image.
    imgCopyButton.Refresh
    ' Call the copy procedure
    EditCopyProc
End Sub

Private Sub imgCopyButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the down state.
    imgCopyButton.Picture = imgCopyButtonDn.Picture
End Sub

Private Sub imgCopyButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the button is pressed, display the up bitmap when the
    ' mouse is dragged outside the button's area; otherwise
    ' display the down bitmap.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgCopyButton.Width Or y < 0 Or y > imgCopyButton.Height Then
            imgCopyButton.Picture = imgCopyButtonUp.Picture
        Else
            imgCopyButton.Picture = imgCopyButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgCopyButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the up state.
    imgCopyButton.Picture = imgCopyButtonUp.Picture
End Sub

Private Sub imgCutButton_Click()
    ' Refresh the image.
    imgCutButton.Refresh
    ' Call the cut procedure
    EditCutProc
End Sub

Private Sub imgCutButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the down state.
    imgCutButton.Picture = imgCutButtonDn.Picture
End Sub

Private Sub imgCutButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the button is pressed, display the up bitmap when the
    ' mouse is dragged outside the button's area; otherwise,
    ' display the down bitmap.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgCutButton.Width Or y < 0 Or y > imgCutButton.Height Then
            imgCutButton.Picture = imgCutButtonUp.Picture
        Else
            imgCutButton.Picture = imgCutButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgCutButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the up state.
    imgCutButton.Picture = imgCutButtonUp.Picture
End Sub

Private Sub imgFileNewButton_Click()
    ' Refresh the image.
    imgFileNewButton.Refresh
    ' Call the new file procedure
    FileNew
End Sub

Private Sub imgFileNewButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the down state.
    imgFileNewButton.Picture = imgFileNewButtonDn.Picture
End Sub

Private Sub imgFileNewButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the button is pressed, display the up bitmap when the
    ' mouse is dragged outside the button's area; otherwise,
    ' display the down bitmap.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgFileNewButton.Width Or y < 0 Or y > imgFileNewButton.Height Then
            imgFileNewButton.Picture = imgFileNewButtonUp.Picture
        Else
            imgFileNewButton.Picture = imgFileNewButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgFileNewButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the up state.
    imgFileNewButton.Picture = imgFileNewButtonUp.Picture
End Sub

Private Sub imgFileOpenButton_Click()
    ' Refresh the image.
    imgFileOpenButton.Refresh
    ' Call the file open procedure
    FileOpenProc
End Sub

Private Sub imgFileOpenButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the down state.
    imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
End Sub

Private Sub imgFileOpenButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the button is pressed, display the up bitmap when the
    ' mouse is dragged outside the button's area; otherwise,
    ' display the down bitmap.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgFileOpenButton.Width Or y < 0 Or y > imgFileOpenButton.Height Then
            imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture
        Else
            imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgFileOpenButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the up state.
    imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture

End Sub

Private Sub imgPasteButton_Click()
    ' Refresh the image.
    imgPasteButton.Refresh
    ' Call the paste procedure
    EditPasteProc
End Sub

Private Sub imgPasteButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the down state.
    imgPasteButton.Picture = imgPasteButtonDn.Picture
End Sub

Private Sub imgPasteButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the button is pressed, display the up bitmap when the
    ' mouse is dragged outside the button's area; otherwise,
    ' display the down bitmap.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgPasteButton.Width Or y < 0 Or y > imgPasteButton.Height Then
            imgPasteButton.Picture = imgPasteButtonUp.Picture
        Else
            imgPasteButton.Picture = imgPasteButtonDn.Picture
        End If
    End Select
End Sub

Private Sub imgPasteButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the picture for the up state.
    imgPasteButton.Picture = imgPasteButtonUp.Picture
End Sub


Private Sub mnuCurposn_Click()
Dim strselectfromcurposn As String
gCurPos = frmSpeakpad.txtNote.SelStart
txtNote.SelLength = Len(txtNote.Text) - gCurPos
strselectfromcurposn = txtNote.SelText
TextToSpeech1.Speak strselectfromcurposn
End Sub

Private Sub mnuEditCopy_Click()
    ' Call the copy procedure
    EditCopyProc
End Sub

Private Sub mnuEditCut_Click()
    ' Call the cut procedure
    EditCutProc
End Sub

Private Sub mnuEditDelete_Click()
' If the mouse pointer is not at the end of the speakpad...
    If txtNote.SelStart <> Len(Screen.ActiveControl.Text) Then
        ' If nothing is selected, extend the selection by one.
        If txtNote.SelLength = 0 Then
            txtNote.SelLength = 1
            ' If the mouse pointer is on a blank line, extend the selection by two.
            If Asc(txtNote.SelText) = 13 Then
                txtNote.SelLength = 2
            End If
        End If
        ' Delete the selected text.
        txtNote.SelText = ""
    End If
End Sub

Private Sub mnuEditPaste_Click()
    ' Call the paste procedure.
    EditPasteProc
End Sub

Private Sub mnuEditSelectAll_Click()
    ' Use SelStart & SelLength to select the text.
    txtNote.SelStart = 0
    txtNote.SelLength = Len(txtNote.Text)
End Sub

Private Sub mnuEditTime_Click()
    ' Insert the current time and date.
    txtNote.SelText = Now
End Sub

Private Sub mnuFileExit_Click()
    ' End the application.
    Unload Me
End Sub

Public Sub mnuFileNew_Click()
    ' Call the new form procedure
    FileNew
End Sub

Private Sub mnuFileOpen_Click()
    ' Call the file open procedure.
    FileOpenProc
End Sub

Private Sub mnufileprint_Click()
FilePrint
End Sub

Private Sub mnuFileSave_Click()
    'Call the file save procedure
    FileSave
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim strSaveFileName As String
    Dim strDefaultName As String
    
    ' Assign the form caption to the variable.
    strDefaultName = Right$(Me.Caption, Len(Me.Caption) - 14)
    If Me.Caption = "SpeakPad   - Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, strSaveFileName.
        
        strSaveFileName = GetFileName("Untitled.txt")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    Else
        ' The form's Caption contains the name of the open file.
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    End If
End Sub

Private Sub mnuFontName_Click(Index As Integer)
    ' Assign the selected font to the textbox fontname property.
    txtNote.FontName = mnuFontName(Index).Caption
    'txtNote.FontName = cmbFont.ListIndex
        
End Sub

Private Sub mnuForwards_Click()
TextToSpeech1.Speed = 400
End Sub

Private Sub mnuOptions_Click()
    ' Toggle the Checked property to match the .Visible property.
    mnuOptionsToolbar.Checked = picToolbar.Visible
End Sub

Private Sub mnuOptionsLaunch_Click()
    Dim strApp As String
    
    ' Shell a new instance of the speakpad.
    strApp = App.Path & "\" & App.EXEName
    Shell strApp, 1
End Sub

Private Sub mnuOptionsToolbar_Click()
    ' Toggle the visible property of the toolbar
    picToolbar.Visible = Not picToolbar.Visible
    ' Change the check to match the current state
    mnuOptionsToolbar.Checked = picToolbar.Visible
    ' Call the resize procedure
    ResizeNote
End Sub

Private Sub mnupause_Click()
If paused = False Then
    TextToSpeech1.Pause
    paused = True
    Exit Sub
End If
If paused = True Then
    TextToSpeech1.Resume
    paused = False
End If
End Sub

Private Sub mnuPlay_Click()
If txtNote.Text <> "" Then TextToSpeech1.Speak txtNote.Text
End Sub

Private Sub mnuPlayselection_Click()
Dim strselect As String
If txtNote.SelText <> "" Then strselect = txtNote.SelText
TextToSpeech1.Speak strselect

End Sub

Private Sub mnuProperties_Click()
Dim strfilename, fileprop  As String
If frmSpeakpad.Caption <> "SpeakPad - Untitled" Then
strfilename = Right(frmSpeakpad.Caption, Len(frmSpeakpad.Caption) - 13)
fileprop = "Size  : " + Str(FileLen(strfilename)) + " bytes "
fileprop = fileprop + Chr(13) + "Date :  " + Str(FileDateTime(strfilename)) + " last Saved/ modififed"
MsgBox fileprop, vbOKOnly, "File properties " + strfilename



'txtNote.Text = FileSystem.FileAttr(strfilename)
End If
End Sub

Private Sub mnuRecentFile_Click(Index As Integer)
    ' Call the file open procedure, passing a
    ' reference to the selected file name
    OpenFile (mnuRecentFile(Index).Caption)
    ' Update the list of recently opened files in the File menu control array.
    GetRecentFiles
End Sub

Private Sub mnuSearchFind_Click()
    ' If there is text in the textbox, assign it to
    ' the textbox on the Find form, otherwise assign
    ' the last findtext value.
    If txtNote.SelText <> "" Then
        frmFind.txtFind.Text = txtNote.SelText
    Else
        frmFind.txtFind.Text = gFindString
    End If
    ' Set the public variable to start at the beginning.
    gFirstTime = True
    ' Set the case checkbox to match the public variable
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    ' Display the Find form.
    frmFind.Show vbModal
End Sub

Private Sub mnuSearchFindNext_Click()
    ' If the public variable isn't empty, call the
    ' find procedure, otherwise call the find menu
    If Len(gFindString) > 0 Then
        FindIt
    Else
        mnuSearchFind_Click
    End If
End Sub



Private Sub mnustatistics_Click()
' no of pages, paragraphs, lines, words, characters with spaces and without spaces
' no of pages = totla no of lines /72?
' no of paragraph = no of carriage return
' word



End Sub

Private Sub mnuStop_Click()
TextToSpeech1.StopSpeaking
End Sub

Private Sub txtNote_Change()
    ' Set the public variable to show that text has changed.
    FState.Dirty = True
End Sub

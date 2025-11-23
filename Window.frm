VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Window 
   BackColor       =   &H80000005&
   Caption         =   "NoteFolioEdit"
   ClientHeight    =   9120
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4935
   Icon            =   "Window.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "Debug"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Note_Name 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "UNTITLED"
      Top             =   360
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Unten ausrichten
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   8625
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Row: 1, Col 1"
            TextSave        =   "Row: 1, Col 1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "0 Chars"
            TextSave        =   "0 Chars"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "TI Size: 40 B"
            TextSave        =   "TI Size: 40 B"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Slider Zoom 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8370
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   2
      Min             =   -5
      Max             =   5
   End
   Begin VB.TextBox Note 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label SavedLabel 
      BackColor       =   &H80000005&
      Caption         =   "*Unsaved"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E6E6E6&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Name (8 chars max)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E6E6E6&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu Menu_File 
      Caption         =   "File"
      NegotiatePosition=   1  'Links
      Begin VB.Menu Menu_File_Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Menu_File_Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu_File_SaveAs 
         Caption         =   "Save as"
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "Edit"
      Begin VB.Menu Menu_Edit_Archived 
         Caption         =   "Archived"
         Checked         =   -1  'True
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Edit_DelFile 
         Caption         =   "Delete File"
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu Menu_View 
      Caption         =   "View"
      Begin VB.Menu Menu_View_CapWidth 
         Caption         =   "Cap to TI-84 Screen Width"
      End
   End
End
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function AddFontMemResourceEx Lib "gdi32.dll" (ByVal pFileView As Long, ByVal cjSize As Long, ByVal pvReserved As Long, pNumFonts As Long) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4

Dim document As DOC

Dim edit As String 'contains the path of the file currently editing
Dim editType As Long '2->.txt 3->.8xv

Dim Archived As Boolean
Dim Saved As Boolean 'will be deleted, replaced by editType = 0

Private Const TEXT_SAVED As String = "Saved"
Private Const TEXT_UNSAVED As String = "*Unsaved"

Private Sub Command1_Click()
    
    Dim i As Long
    
    With App
        Print .title
        Print "v" & .Major & "." & .Minor & "." & .Revision
    End With
    
    Note.text = "ASCII table from 0x21 to 0xFF" & vbCrLf
    
    For i = &H21 To &HFF
        Note.text = Note.text & i & " (0x" & Hex$(i) & ") " & ChrW$(i) & vbCrLf
    Next
    
End Sub

Private Sub Form_Load()

    Dim font() As Byte
    Dim nFont As Long
    
    document.title = Note_Name.text
    Archived = True
    
    font = LoadResData(1, "FONT")
    AddFontMemResourceEx VarPtr(font(0)), UBound(font) + 1, 0, nFont
    Note.font = "TI-84 Plus Calculator font"
    
    Saved = True
    
    'hMenu = GetMenu(hWnd)
    'InsertMenu GetSubMenu(hMenu, 0), 1, MF_BYPOSITION Or MF_SEPARATOR, 0, vbNullString
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not Saved Then
        Dim result As Long
    
        Beep
        result = MsgBox("Do you want to save changes to [" & Note_Name.text & "]?", vbYesNoCancel Or vbQuestion)
        If result = vbCancel Then Cancel = 1
        If result = vbYes Then
            Menu_File_Save_Click
        End If
    End If
        
End Sub

Private Sub Form_Resize()
    
    Zoom.Top = Window.Height - 1635
    
    Note.Width = Window.Width - 720
    If Window.Height > 5070 Then Note.Height = Window.Height - 5070
    
End Sub

Private Sub Menu_File_Click()
    'make Note_Name loose focus
    Note.SetFocus
    'DoEvents
End Sub

Private Sub Menu_File_Open_Click()
    
    Dim ofn As OPENFILENAME
    Dim path As String
    
    Dim result As Long
    
    If Not Saved Then
        Beep
        result = MsgBox("Do you want to save changes to [" & Note_Name.text & "]?", vbYesNoCancel Or vbQuestion)
        If result = vbCancel Then Exit Sub
        If result = vbYes Then
            Menu_File_Save_Click
        End If
    End If
    
    Window.MousePointer = vbHourglass
    
    With ofn
        .lStructSize = LenB(ofn)
        .hwndOwner = hWnd
        .lpstrFilter = "All Files (*.*)" & ChrW$(0) & "*.*" & ChrW$(0) & "Text Files (*.txt)" & ChrW$(0) & "*.txt" & ChrW$(0) & "TI 83/84 Variable Files (*.8xv)" & ChrW$(0) & "*.8xv" & ChrW$(0)
        .nFilterIndex = 3
        .lpstrFile = String$(260, 0)
        .nMaxFile = 260
        .Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    End With
    
    If GetOpenFileName(ofn) <> 0 Then
        path = Left$(ofn.lpstrFile, lstrlenW(StrPtr(ofn.lpstrFile)))
        
        editType = ofn.nFilterIndex - 1
        edit = path
        
        Select Case ofn.nFilterIndex
        
            Case 2 'txt
                With document
                    .title = "UNTITLED"
                    .text = ReadFileTxt(path)
                    
                    Note_Name.text = .title
                    Note.text = .text
                End With
                
            Case 3 '8xv
                document = DOC_Read(path)
                With document
                    Note_Name.text = .title
                    Note.text = .text
                End With
        
        End Select
        
        Caption = App.title & " - " & Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, lstrlenW(StrPtr(ofn.lpstrFile)))
        Saved = True
        SavedLabel.Caption = TEXT_SAVED
        Menu_Edit_DelFile.Enabled = True
    End If
    
    Window.MousePointer = vbDefault
    
End Sub

Private Sub Menu_File_Save_Click()
    
    Window.MousePointer = vbHourglass
        
    Select Case editType
        
        Case 0 'not saved
            Menu_File_SaveAs_Click
            
        Case 1 'txt
            WriteFileTxt edit, Note.text
            
            Saved = True
            SavedLabel.Caption = TEXT_SAVED
            
        Case 2 '8xv
            With document
                .title = Note_Name.text
                .text = Note.text
            End With
            DOC_Write edit, document, Archived
            
            Saved = True
            SavedLabel.Caption = TEXT_SAVED
            Menu_Edit_DelFile.Enabled = True
            
    End Select
    
    Window.MousePointer = vbDefault
    
End Sub

Private Sub Menu_File_SaveAs_Click()
    
    Dim ofn As OPENFILENAME
    Dim path As String
    
    Window.MousePointer = vbHourglass
    
    With ofn
        .lStructSize = LenB(ofn)
        .hwndOwner = hWnd
        .lpstrFilter = "Text Files (*.txt)" & ChrW$(0) & "*.txt" & ChrW$(0) & "TI 83/84 Variable Files (*.8xv)" & ChrW$(0) & "*.8xv" & ChrW$(0)
        .nFilterIndex = 2
        .lpstrFile = document.title & String$(260 - Len(document.title), 0)
        .nMaxFile = 260
        .Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
    End With
    
    If GetSaveFileName(ofn) <> 0 Then
        path = Left$(ofn.lpstrFile, lstrlenW(StrPtr(ofn.lpstrFile)))
        
        editType = ofn.nFilterIndex
        edit = path
        
        Select Case ofn.nFilterIndex
        
            Case 1 'txt
                WriteFileTxt path & ".txt", Note.text
                Caption = App.title & " - " & Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, lstrlenW(StrPtr(ofn.lpstrFile) + (ofn.nFileOffset + 1) * 2) + 1) & ".txt"
                edit = path & ".txt"
                
            Case 2 '8xv
                With document
                    .title = Note_Name.text
                    .text = Note.text
                End With
                DOC_Write path & ".8xv", document, Archived
                Caption = App.title & " - " & Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, lstrlenW(StrPtr(ofn.lpstrFile) + (ofn.nFileOffset + 1) * 2) + 1) & ".8xv"
                edit = path & ".8xv"
        
        End Select
        
        Saved = True
        SavedLabel.Caption = TEXT_SAVED
    End If
    
    Window.MousePointer = vbDefault
    
End Sub

Private Sub Menu_Edit_Archived_Click()
    
    Menu_Edit_Archived.Checked = Not Menu_Edit_Archived.Checked
    Archived = Menu_Edit_Archived.Checked
    If Saved And (editType <> 1) Then
        SavedLabel.Caption = TEXT_UNSAVED
        Saved = False
    End If
    
End Sub

Private Sub Menu_Edit_DelFile_Click()
    
    On Error GoTo ErrorHandler
    
    Dim result As Long
    
    result = MsgBox("Are you sure you want to permanently delete [" & edit & "]?", vbYesNo Or vbExclamation)
    
    If result = vbYes Then
        Kill edit
        editType = 0
        Saved = False
        SavedLabel.Caption = TEXT_UNSAVED
        Window.Caption = App.title
        Menu_Edit_DelFile.Enabled = False
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox Err.DESCRIPTION, vbCritical
    Err.Clear
    
End Sub

Private Sub Menu_View_CapWidth_Click()
    
    Window.WindowState = vbNormal
    Window.Width = 5175
    
End Sub

Private Sub Note_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then UpdatePos
    
End Sub

Private Sub Note_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Note.SelLength = 0 Then StatusBar.Panels(2).text = "Chars: " & Len(Note.text)
    
End Sub

Private Sub Note_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        UpdatePos
        StatusBar.Panels(2).text = "Chars: " & Len(Note.SelText) & "/" & Len(Note.text)
    End If
    
End Sub

Private Sub Note_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode >= 37 And KeyCode <= 40 Then
        UpdatePos
        If Shift Then StatusBar.Panels(2).text = "Chars: " & Len(Note.SelText) & "/" & Len(Note.text)
    End If
    
End Sub

Private Sub Note_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyA And (Shift And vbCtrlMask) Then
        Note.SelStart = 0
        Note.SelLength = Len(Note.text)
        
        StatusBar.Panels(2).text = "Chars: " & Len(Note.SelText) & "/" & Len(Note.text)
    End If
    
End Sub

Private Sub Note_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
            
        Case 1
            KeyAscii = 0
    
    End Select
    
End Sub

Private Sub Note_Change()
    
    If Saved Then
        SavedLabel.Caption = TEXT_UNSAVED
        Saved = False
    End If
    
    StatusBar.Panels(2).text = "Chars: " & Len(Note.text)
    StatusBar.Panels(3).text = "TI Size: " & NOTE_EMPTY_SIZE + TiLen(Note.text) & " B"
    UpdatePos
    
End Sub

Private Sub Note_Name_Change()
    
    If Len(Note_Name.text) > 8 Then
        Note_Name.ForeColor = vbRed
    Else
        If Saved And (editType <> 1) Then
            SavedLabel.Caption = TEXT_UNSAVED
            Saved = False
        End If
    
        Note_Name.ForeColor = vbWindowText
    End If
    
End Sub

'ctrl a
Private Sub Note_Name_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyA And (Shift And vbCtrlMask) Then
        Note_Name.SelStart = 0
        Note_Name.SelLength = Len(Note_Name.text)
    End If
    
End Sub

Private Sub Note_Name_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        
        Case vbKeyReturn
            Note.SetFocus
            KeyAscii = 0
            
        Case 1
            KeyAscii = 0
    
    End Select
    
End Sub

Private Sub Note_Name_LostFocus()

    If Len(Note_Name.text) > 8 Then
        MsgBox "Name can only be 8 characters long!", vbExclamation
        Note_Name.text = document.title
        Note_Name.SetFocus
    ElseIf Len(Note_Name.text) = 0 Then
        MsgBox "This field cannot be empty!", vbExclamation
        Note_Name.text = document.title
        Note_Name.SetFocus
    Else
        Note_Name.text = StrConv(Note_Name.text, vbUpperCase)
        document.title = Note_Name.text
    End If
    
End Sub

Private Sub Zoom_Scroll()
    Note.FontSize = 12 + Zoom.Value
End Sub

Private Sub UpdatePos()
    
    Dim row As Long
    Dim col As Long
    
    With Note
        Pos .text, .SelStart + .SelLength, row, col
        StatusBar.Panels(1).text = "Row: " & row & ", Col: " & col
    End With
    
End Sub
















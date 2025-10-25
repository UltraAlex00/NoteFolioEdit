Attribute VB_Name = "Lib"
Public Const NOTE_EMPTY_SIZE As Long = 40

Private Declare Function AddFontMemResourceEx Lib "gdi32" (ByVal pbFont As Long, ByVal cbFont As Long, ByVal pdv As Long, ByRef pcFonts As Long) As Long

Public Function ReadFileTxt(ByVal FilePath As String) As String
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open FilePath For Input As #fileNum
    ReadFileTxt = Input$(LOF(fileNum), #fileNum)
    Close #fileNum
    
End Function

Public Sub WriteFileTxt(ByVal FilePath As String, ByVal Content As String)
    Dim fileNum As Integer

    fileNum = FreeFile
    Open FilePath For Output As #fileNum
    Print #fileNum, Content
    Close #fileNum
End Sub

Public Sub Pos(ByVal text As String, ByVal sel As Long, ByRef line As Long, ByRef column As Long)
    Dim i As Long
    Dim crlfPos As Long
    Dim lastCrLf As Long

    line = 1
    lastCrLf = 0
    crlfPos = InStr(1, text, vbCrLf)

    ' Loop through all vbCrLf occurrences before the sel position
    Do While crlfPos > 0 And crlfPos < sel
        line = line + 1
        lastCrLf = crlfPos + 1 ' position after the CRLF
        crlfPos = InStr(crlfPos + 2, text, vbCrLf) ' Skip past current vbCrLf
    Loop

    ' Column (row) is sel position minus last line start
    column = sel - lastCrLf + 1
End Sub

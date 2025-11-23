Attribute VB_Name = "Parser"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long

'file 103 bytes + text      (bytes)
'lenght_data 46 + text      (bytes)
'lenght_content 29 + text   (bytes)

'File header
'data
'   entry header
'   Value
'       NOTEFOLIO_HEADER
'       text <-
'       checksum offset <-
'checksum <-


'size 55
Private Type TI_8XV_FILE_HEADER
    signature(7) As Byte
    signatureEx(2) As Byte
    comment(41) As Byte
    lenght_data(1) As Byte 'Integer bc of dogshit vb6 padding
    'data
    'file checksum
End Type

'size 17
Private Type TI_8XV_ENTRY_HEADER
    signature As Integer
    lenght_content As Integer
    type As Byte
    Label(7) As Byte
    version As Byte
    Flags As Byte '0x80 is archived
    lenght_content2(1) As Byte 'copy of lenght_content, Integer
End Type

'size 26
Private Type NOTEFOLIO_HEADER
    lenght As Integer 'belongs to the Y_VAR/STR header
    signature(3) As Byte 'struct starts here
    reserved(3) As Byte
    Label(7) As Byte
    size As Integer 'text lenght + 24
    signature2 As Integer
    size2 As Integer 'copy of lenght text
    signature3 As Integer
    'text
    'magic (checksum offset, can be 0)
End Type

Public Type DOC
    title As String 'normally I would do wchar_t title[8] but vb6 initialises for some reason " " 0x20 insted of 0x00 so it fucks up everything
    text As String
End Type

Private Const FILE_SIZE_EMPTY = 103
Private Const DATA_SIZE_EMPTY = 46
Private Const CONTENT_SIZE_EMPTY = 29

Private Const TITLE_START_OFFSET = &H52
Private Const TEXT_START_OFFSET = &H62

Private Const DATA_START_OFFSET = &H37
Private Const NOTE_START_OFFSET = &H48

'Private Const DATA_LEN_PTR = &H35
'Private Const VALUE_LEN_PTR = &H46

Const DESCRIPTION As String = "Created by NoteFolioEdit v1.0"

Private Function ReadFile(ByVal Name As String) As Byte()

    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open Name For Binary As fileNum
    ReDim ReadFile(LOF(fileNum) - 1)
    Get fileNum, , ReadFile
    Close fileNum

End Function

Private Sub WriteFile(ByVal FilePath As String, data() As Byte)
    Dim fileNum As Integer
    fileNum = FreeFile

    Open FilePath For Binary Access Write As #fileNum
    Put #fileNum, , data
    Close #fileNum
End Sub

Public Function DOC_Read(ByVal path As String) As DOC
    
    Dim bytes() As Byte
    
    bytes = ReadFile(path)
    
    With DOC_Read
        .title = TiToStr(bytes, TITLE_START_OFFSET, 8)
        .text = TiToStr(bytes, TEXT_START_OFFSET, lstrlenA(bytes(TEXT_START_OFFSET)))
    End With
    
End Function

Public Sub DOC_Write(ByVal path As String, document As DOC, Optional ByVal Archived As Boolean)
    
    Dim bytes() As Byte
    Dim header As TI_8XV_FILE_HEADER
    Dim entry As TI_8XV_ENTRY_HEADER
    Dim Note As NOTEFOLIO_HEADER
    Dim TEXT_END_OFFSET As Long
    Dim textlen As Long
    
    'textlen = Len(Replace(document.text, vbCrLf, vbCr))
    textlen = TiLen(document.text)
    
    With header
        StrToTi .signature, "**TI83F*", True
        'buffer = StrConv("**TI83F*", vbFromUnicode)
        'CopyMemory .signature(0), buffer(0), 8
        .signatureEx(0) = &H1A
        .signatureEx(1) = &HA
        .signatureEx(2) = &HA
        StrToTi .comment, DESCRIPTION, True
        CopyMemory .lenght_data(0), CInt(DATA_SIZE_EMPTY + textlen), 2
    End With
    
    With entry
        .signature = &HD
        .lenght_content = CONTENT_SIZE_EMPTY + textlen
        .type = &H15
        StrToTi .Label, document.title, True
        '.version = &H0
        If Archived Then .Flags = &H80
        '.flags = &H80
        CopyMemory .lenght_content2(0), .lenght_content, 2
    End With
    
    With Note
        .lenght = entry.lenght_content - 2
        .signature(0) = &HF3
        .signature(1) = &H47
        .signature(2) = &HBF
        .signature(3) = &HAF
        CopyMemory .Label(0), entry.Label(0), 8
        .size = textlen + 24
        .signature2 = &H18
        .size2 = .size
        .signature3 = &H1
    End With
    
    
    With document
        ReDim bytes(FILE_SIZE_EMPTY + textlen - 1)
        TEXT_END_OFFSET = TEXT_START_OFFSET + textlen + 1
        
        CopyMemory bytes(0), header, 55
        CopyMemory bytes(DATA_START_OFFSET), entry, 17
        CopyMemory bytes(NOTE_START_OFFSET), Note, 26
        StrToTi bytes, .text, True, True, TEXT_START_OFFSET
        'write checksum on last 2 bytes, TEXT_START_OFFSET + textlen + 3 -> Ubound(bytes) - 2
        CopyMemory bytes(TEXT_START_OFFSET + textlen + 3), ChecksumI16(bytes, DATA_START_OFFSET, DATA_SIZE_EMPTY + textlen), 2
    End With
    
    WriteFile path, bytes
    
End Sub

Public Function ChecksumI16(bytes() As Byte, Optional ByVal offset As Long = 0, Optional ByVal lenght As Long = 0) As Integer
    Dim i As Long
    Dim checksum As Long
    
    If lenght = 0 Then lenght = UBound(bytes) + 1
    
    For i = offset To offset + lenght - 1
        checksum = checksum + bytes(i)
        If checksum > &HFFFF& Then
            checksum = checksum - &H10000 ' Subtract 65536 (Integer limit)
        End If
    Next i
    
    CopyMemory ChecksumI16, checksum, 2
    
End Function



Private Function TiToStr(chars() As Byte, Optional ByVal offset As Long, Optional ByVal lenght As Long) As String
    
    Dim i As Long
    
    If lenght = 0 Then lenght = UBound(chars) + 1
    
    For i = offset To offset + lenght - 1
        
        Select Case chars(i)
            
            Case &HD6: TiToStr = TiToStr & vbCrLf
            Case &HF1: TiToStr = TiToStr & " "
            Case Else: TiToStr = TiToStr & ChrW$(chars(i))
            
        End Select
    Next i

End Function
 
Sub StrToTi(chars() As Byte, ByVal str As String, Optional ByVal existingArray As Boolean, Optional ByVal translate As Boolean, Optional ByVal offset As Long, Optional ByVal lenght As Long)
    
    Dim i As Long
    Dim buffer() As Byte
    
    If translate Then str = Replace(str, vbCrLf, ChrW$(&HD6))
    
    ReDim buffer(LenB(str) - 1)
    CopyMemory buffer(0), ByVal StrPtr(str), LenB(str)
    
    If lenght = 0 Then lenght = Len(str)
    
    If Not existingArray Then ReDim chars(lenght - 1)
    
    If translate Then
        Dim j As Long
        Dim lastSpace As Long
        lastSpace = -1
        
        For i = 0 To lenght - 1
            If buffer(i * 2) = &HD6 Then
                lastSpace = -1
                j = 0
            End If
            If buffer(i * 2) = &H20 Then lastSpace = i
            If j Mod 16 = 0 Then
                If lastSpace >= 0 Then
                    chars(lastSpace + offset) = &HF1
                    j = 0
                End If
            End If
            
            chars(i + offset) = buffer(i * 2)
            
            j = j + 1
        Next i
    Else
        For i = 0 To lenght - 1
            chars(i + offset) = buffer(i * 2)
        Next i
    End If
    
    
    
End Sub

Public Function TiLen(ByVal text As String) As Long

    Dim lastCrLf As Long
    Dim n As Long
    
    lastCrLf = InStr(lastCrLf + 1, text, vbCrLf)
    
    Do Until lastCrLf = 0
        lastCrLf = InStr(lastCrLf + 1, text, vbCrLf)
        n = n + 1
    Loop
    
    TiLen = Len(text) - n
    
    
End Function




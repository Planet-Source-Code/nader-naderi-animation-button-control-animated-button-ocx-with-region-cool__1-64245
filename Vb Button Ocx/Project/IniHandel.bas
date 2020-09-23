Attribute VB_Name = "IniHandel"
Public Function FindSection(IniFileName As String, Section As String) As Boolean
Dim FileSize As Long
Dim NumberSeek As Long
Dim GetChar As String
Dim TempCh As String

Open IniFileName For Input As #1
NextSection:
'Revalue For GOTO State
GetChar = ""
TempCh = ""
'NumberSeek = 0

While TempCh <> "[" 'Search for first section cahracter
        
            If EOF(1) Then
               FindSection = False
               Close #1
               Exit Function
               End If
        NumberSeek = NumberSeek + 1
        Seek #1, NumberSeek
        TempCh = Input(1, #1)
        GetChar = TempCh
Wend

GetChar = "" 'Empty Agine
TempCh = ""

While TempCh <> "]"
            If EOF(1) Then
               FindSection = False
               Close #1
               Exit Function
               End If
NumberSeek = NumberSeek + 1
Seek #1, NumberSeek
GetChar = GetChar + TempCh
TempCh = Input(1, #1)
Wend

If GetChar = Section Then 'Looking for Section
  FindSection = True
  Close #1
  Exit Function
Else
    GoTo NextSection 'Go for searching another section match
End If

Close #1
End Function

Public Function GetIniData(IniFileName As String, Section As String, Key As String) As String
Dim FileSize As Long
Dim NumberSeek As Long
Dim GetChar As String
Dim TempCh As String

Open IniFileName For Input As #1
NextSection:
'Revalue For GOTO State
GetChar = ""
TempCh = ""
'NumberSeek = 0

While TempCh <> "[" 'Search for first section cahracter
        
            If EOF(1) Then
               Close #1
               Exit Function
               End If
        NumberSeek = NumberSeek + 1
        Seek #1, NumberSeek
        TempCh = Input(1, #1)
        GetChar = TempCh
Wend

GetChar = "" 'Empty Agine
TempCh = ""

While TempCh <> "]"
            If EOF(1) Then
               Close #1
               Exit Function
               End If
NumberSeek = NumberSeek + 1
Seek #1, NumberSeek
GetChar = GetChar + TempCh
TempCh = Input(1, #1)
Wend

'Print GetChar

If GetChar = Section Then 'Looking for Section
ReFindKey:
    GetChar = "" 'Empty Agine
    TempCh = ""
    While TempCh <> "="
            If EOF(1) Then
               Close #1
               Exit Function
               End If
            
        NumberSeek = NumberSeek + 1
        Seek #1, NumberSeek
        If TempCh = "[" Then GoTo NextSection
        If TempCh <> vbCr And TempCh <> vbLf Then
        GetChar = GetChar + TempCh
        TempCh = Input(1, #1)
        Else
        TempCh = Input(1, #1)
        TempCh = "="
        End If
    Wend
  If GetChar = Key And GetChar <> "" Then
    GetChar = "" 'Empty Agine
    TempCh = ""
    Do While TempCh <> vbCr
            If EOF(1) Then
               GetChar = GetChar + TempCh
               Exit Do
               End If
              NumberSeek = NumberSeek + 1
              Seek #1, NumberSeek
              GetChar = GetChar + TempCh
              TempCh = Input(1, #1)
              
    Loop
            GetIniData = GetChar
            Close #1
            Exit Function
  Else
            GoTo ReFindKey
  End If
Else
    GoTo NextSection 'Go for searching another section match

End If

Close #1
End Function

Public Function GetNumberSections(IniFileName As String) As Long
Dim FileSize As Long
Dim NumberSeek As Long
Dim TempCh As String

Open IniFileName For Input As #1
NextSection:
While Not EOF(1)
Do While TempCh <> "[" 'Search for first section cahracter
        If EOF(1) Then
               Close #1
              Exit Function
        End If
        NumberSeek = NumberSeek + 1
        Seek #1, NumberSeek
        TempCh = Input(1, #1)
Loop

GetChar = "" 'Empty Agine
TempCh = ""

Do While TempCh <> "]"
            If EOF(1) Then
               Close #1
               Exit Function
               End If
NumberSeek = NumberSeek + 1
Seek #1, NumberSeek
TempCh = Input(1, #1)
Loop

GetNumberSections = GetNumberSections + 1
Wend
Close #1
End Function

'Private Sub Command1_Click()
'MsgBox GetIniData("c:\2.ini", Text1.Text, Text2.Text)
'MsgBox GetNumberSections("c:\windows\win.ini")
'End Sub



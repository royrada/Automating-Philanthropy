Attribute VB_Name = "Strip"
'Locates the IRS990_ActivityOrMissionDesc column on the Parsed990Data sheet.
'Inserts a new adjacent column titled DescFiltered.
'Loads punctuation and stop words from their respective .txt files in the workbook's folder.
'Replaces each punctuation character with a space.
'Removes stop words when surrounded by word boundaries, ensuring "i" is only removed when isolated.

Sub Master()
Call CleanAndStripDescriptions
Call SortAndDeduplicateDescriptions
Call CleanseWebsiteAddresses
MsgBox "Descriptions stripped and cleaned successfully!", vbInformation
End Sub

Sub CleanAndStripDescriptions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Parsed990Data")
    ' Locate the column
    Dim lastCol As Long, headerCell As Range
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim descCol As Long, insertCol As Long
    descCol = 0
    For Each headerCell In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        If Trim(headerCell.value) = "IRS990_ActivityOrMissionDesc" Then
            descCol = headerCell.Column
            Exit For
        End If
    Next headerCell
    If descCol = 0 Then
        MsgBox "IRS990_ActivityOrMissionDesc column not found.", vbExclamation
        Exit Sub
    End If
    insertCol = descCol + 1
    ws.Columns(insertCol).Insert Shift:=xlToRight
    ws.Cells(1, insertCol).value = "DescFiltered"
    ' Load punctuation.txt
    Dim punctFile As String, stopFile As String
    punctFile = ThisWorkbook.Path & "\punctuation.txt"
    stopFile = ThisWorkbook.Path & "\stopwords.txt"
    Dim punctChars As String, stopWords() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    ' Read punctuation into string
    If fso.FileExists(punctFile) Then
        Set ts = fso.OpenTextFile(punctFile, 1)
        punctChars = ts.ReadAll
        ts.Close
    Else
        MsgBox "Punctuation.txt file not found.", vbExclamation
        Exit Sub
    End If
    ' Read stop words into array
    If fso.FileExists(stopFile) Then
        Set ts = fso.OpenTextFile(stopFile, 1)
        Dim stopContent As String
        stopContent = ts.ReadAll
        ts.Close
        stopWords = Split(stopContent, vbCrLf)
    Else
        MsgBox "Stopwords.txt file not found.", vbExclamation
        Exit Sub
    End If
    ' Process each row
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, descCol).End(xlUp).Row
    Dim i As Long, original As String, cleaned As String, sw As Variant
    For i = 2 To lastRow
        original = ws.Cells(i, descCol).value
        If Not IsError(original) And Trim(original) <> "" Then
            cleaned = LCase(original)
Dim charArray() As String
Dim j As Long
ReDim charArray(Len(punctChars) - 1)
For j = 1 To Len(punctChars)
    charArray(j - 1) = mid(punctChars, j, 1)
Next j
' Iterate using Variant
Dim ch As Variant
For Each ch In charArray
    cleaned = Replace(cleaned, ch, " ")
Next ch
            ' Pad with spaces for boundary-based stopword removal
            cleaned = " " & cleaned & " "
            ' Remove stop words surrounded by spaces
            For Each sw In stopWords
                If Trim(sw) <> "" Then
                    cleaned = Replace(cleaned, " " & LCase(sw) & " ", " ")
                End If
            Next sw
            ' Final trimming and spacing
            cleaned = Application.WorksheetFunction.Trim(cleaned)
            ws.Cells(i, insertCol).value = cleaned
        Else
            ws.Cells(i, insertCol).value = ""
        End If
    Next i
End Sub

Sub CleanseWebsiteAddresses()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colIdx As Long
    Dim i As Long
    Dim rawText As String
    Dim cleanedText As String
    Dim rgx As Object
    Set ws = ThisWorkbook.Sheets("Parsed990Data")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    colIdx = Application.Match("IRS990_WebsiteAddressTxt", ws.Rows(1), 0)
    ' Create RegExp object for stripping protocol patterns
    Set rgx = CreateObject("VBScript.RegExp")
    With rgx
        .Global = True
        .IgnoreCase = True
        ' Matches http or https with optional colons, spaces, and slashes
        .Pattern = "^\s*(https?)\s*[:]*\s*/{0,2}\s*"
    End With
    For i = 2 To lastRow
        rawText = Trim(ws.Cells(i, colIdx).value)
        Select Case UCase(rawText)
            Case "N/A", "NONE"
                ws.Cells(i, colIdx).value = ""
            Case Else
                cleanedText = rgx.Replace(rawText, "")
                ws.Cells(i, colIdx).value = cleanedText
        End Select
    Next i
End Sub

Sub SortAndDeduplicateDescriptions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Parsed990Data")

    Dim colIndex As Long
    colIndex = GetColumnIndex(ws, "DescFiltered")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        Dim cellValue As String
        cellValue = Trim(ws.Cells(i, colIndex).value)
        If Len(cellValue) > 0 Then
            Dim wordList() As String
            wordList = Split(cellValue, " ")
            Dim cleanedList() As String
            cleanedList = DeduplicateAndSort(wordList)
            ws.Cells(i, colIndex).value = Join(cleanedList, " ")
        End If
    Next i
End Sub

Function DeduplicateAndSort(words() As String) As String()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(words) To UBound(words)
        Dim w As String
        w = Trim(words(i))
        If Len(w) > 0 Then
            If Not dict.exists(w) Then dict.Add w, True
        End If
    Next i
    Dim v As Variant
    v = dict.Keys
    Dim uniqueWords() As String
    ReDim uniqueWords(LBound(v) To UBound(v))
    For i = LBound(v) To UBound(v)
        uniqueWords(i) = CStr(v(i))
    Next i
    Call QuickSortString(uniqueWords, LBound(uniqueWords), UBound(uniqueWords))
    DeduplicateAndSort = uniqueWords
End Function

Sub QuickSortString(arr() As String, low As Long, high As Long)
    Dim pivot As String, i As Long, j As Long, temp As String
    If low < high Then
    ' \ means integer division where result is always an integer
   '     pivot = arr((low + high) \ 2)
        pivot = arr(Int((low + high) / 2))
        i = low
        j = high
        Do While i <= j
            Do While arr(i) < pivot: i = i + 1: Loop
            Do While arr(j) > pivot: j = j - 1: Loop
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        Call QuickSortString(arr, low, j)
        Call QuickSortString(arr, i, high)
    End If
End Sub

Function GetColumnIndex(ws As Worksheet, nodename As String) As Integer
    Dim colNum As Integer
    Dim cleanNode As String
    cleanNode = Trim(Replace(nodename, Chr(160), ""))

    For colNum = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

        If Trim(Replace(ws.Cells(1, colNum).value, Chr(160), "")) = cleanNode Then
            GetColumnIndex = colNum
            Exit Function
        End If
    Next colNum

    GetColumnIndex = 0 ' Return 0 if not found
End Function




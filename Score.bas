Attribute VB_Name = "Score"
Option Explicit
'by declaring as Dim before any Sub or code the Dim applies to entire module but not any other module
' if did as Public then for all modules in all project
Option Base 1
'Option Compare Text 'vbTextCompare 1 Performs a textual comparison.
Dim wsParsed990 As Worksheet
Dim wsScored990 As Worksheet
Dim RulesTwoDimension() As Variant
Dim lastRowParsed990 As Long
Dim UboundRules As Integer
Dim Scores() As Integer 'to hold scores before passed to worksheet

Sub Score()
'The  Sub Score is the master of the scoring program.
Dim Counter1 As Long, Counter2 As Long, Counter3 As Long  ' these are counters for loops
Call Initialize
ReDim Scores(1 To lastRowParsed990 - 1)
'Next the Sub Score begins a For loop that iterates on Counter1 from 1 to NumRules,
'takes Rules(Counter1), extracts the first parameter from Rules(Counter1),
'puts that in a string RuleType and uses that as a parameter in a Select Case logic.
Dim ruleType As String
For Counter1 = 1 To UboundRules
    ruleType = RulesTwoDimension(Counter1, 1)
Select Case ruleType
    Case "Substring"
        Call Substr(Counter1)
    Case "Trend"
        Call Trend(Counter1)
    Case "Percentile"
        Call Percentile(Counter1)
    Case "Eval"
        Call Eval(Counter1)
    Case Else
        Debug.Print "In Sub Score Select failed; ruletype = " & ruleType
    End Select
Call CopyScorestoScored990(Scores(), Counter1)
ReDim Scores(1 To lastRowParsed990 - 1)
Next Counter1
Call FinalizeScoredData990
MsgBox "Score Finished", vbInformation
End Sub

Sub Initialize()
Dim Counter1 As Long
Set wsParsed990 = ThisWorkbook.Sheets("Parsed990Data")
Set wsScored990 = ThisWorkbook.Sheets("Scored990Data")
wsScored990.Cells.Clear
lastRowParsed990 = wsParsed990.Cells(wsParsed990.Rows.count, 1).End(xlUp).Row
'populate entity ids in Scored990
    wsScored990.Cells(1, 1).value = wsParsed990.Cells(1, 1).value
    For Counter1 = 2 To lastRowParsed990
        wsScored990.Cells(Counter1, 1).NumberFormat = "@"
        wsScored990.Cells(Counter1, 1).value = "'" & wsParsed990.Cells(Counter1, 1).text
    Next Counter1
    
Call BuildRuleFile
'now can put rule headers across row 1
For Counter1 = 1 To UboundRules
    wsScored990.Cells(1, Counter1 + 1).value = RulesTwoDimension(Counter1, 2)
Next Counter1
End Sub


Sub BuildRuleFile()
Dim ruleFile As String
Dim AllRuleContent As String
Dim AllRulesList() As String
Dim ruleType As String
Dim OneRuleLine() As String
Dim Counter1 As Long, Counter2 As Long, Counter3 As Long  ' these are counters for loops
    ' Read rules file
    ruleFile = ThisWorkbook.Path & "\rule.txt"
'    ruleFile = ThisWorkbook.Path & "\ruleWebEmpty.txt"
Dim DummyString As String
AllRuleContent = ReadFile(ruleFile)
AllRulesList = MakeSplitBase1(AllRuleContent, vbCrLf)
UboundRules = UBound(AllRulesList)
ReDim RulesTwoDimension(UboundRules, 6)
' number of rules and I know max number of parameters is currently 6.
    ' Parse rule file content
For Counter1 = 1 To UboundRules
    DummyString = AllRulesList(Counter1)
    OneRuleLine = MakeSplitBase1(DummyString, ";")
'    OneRuleLine = Split(dummyString, ";", -1, 1)
    ruleType = Trim(OneRuleLine(1))
    RulesTwoDimension(Counter1, 1) = ruleType
    Select Case ruleType
        Case "Substring", "Eval"
            For Counter2 = 2 To 5
                RulesTwoDimension(Counter1, Counter2) = OneRuleLine(Counter2)
            Next Counter2
        Case "Trend"
            For Counter2 = 2 To 3
                RulesTwoDimension(Counter1, Counter2) = OneRuleLine(Counter2)
            Next Counter2
        Case "Percentile"
            For Counter2 = 2 To 4
                RulesTwoDimension(Counter1, Counter2) = OneRuleLine(Counter2)
            Next Counter2
        Case Else
            Debug.Print "In Sub BuildRuleFil Select failed; ruletype = " & ruleType
End Select
Next Counter1

End Sub

Sub Substr(index)
    Dim ParsedColumn() As Variant
    Dim RuleNodeName As String, RulePresent As String, DummyString1 As String
    Dim RuleTokens() As String, DummyString2() As String, DummyString3 As String
    Dim Counter1 As Long, Counter2 As Long, Counter3 As Long
    Dim NodenameColumnIndex As Integer, returnIndex As Integer
    Dim NumberOfTokens As Integer

    RuleNodeName = RulesTwoDimension(index, 3)
    RulePresent = RulesTwoDimension(index, 4)
    DummyString1 = Trim(RulesTwoDimension(index, 5))

    If DummyString1 = "" Then Exit Sub ' No tokens to test

    RuleTokens = MakeSplitBase1(DummyString1, ",")
    NumberOfTokens = UBound(RuleTokens)

    ' Clean and lowercase tokens
    For Counter1 = 1 To NumberOfTokens
        RuleTokens(Counter1) = Trim(LCase(RuleTokens(Counter1)))
    Next Counter1

    NodenameColumnIndex = GetColumnIndex(wsParsed990, RuleNodeName)
    ReDim ParsedColumn(1 To lastRowParsed990 - 1)

    ' Populate ParsedColumn
    For Counter1 = 2 To lastRowParsed990
        ParsedColumn(Counter1 - 1) = wsParsed990.Cells(Counter1, NodenameColumnIndex).value
    Next Counter1

    ' Iterate rows of data
    For Counter1 = 1 To lastRowParsed990 - 1
        DummyString1 = ParsedColumn(Counter1)
        returnIndex = 0 ' Reset before each record

        If Not IsError(DummyString1) Then
            DummyString1 = Trim(LCase(DummyString1))
            If DummyString1 <> "" Then
                DummyString2 = MakeSplitBase1(DummyString1, " ")
                For Counter2 = 1 To NumberOfTokens
                    For Counter3 = LBound(DummyString2) To UBound(DummyString2)
                        DummyString3 = Trim(LCase(DummyString2(Counter3)))
                        If DummyString3 <> "" Then
                            If InStr(DummyString3, RuleTokens(Counter2)) > 0 Then
                                returnIndex = 1
                                Exit For
                            End If
                        End If
                    Next Counter3
                    If returnIndex = 1 Then Exit For
                Next Counter2
            End If
        End If

        ' Score logic
        If RulePresent = "T" Then
            Scores(Counter1) = IIf(returnIndex = 1, 1, 0)
        ElseIf RulePresent = "F" Then
            Scores(Counter1) = IIf(returnIndex = 1, 0, 1)
        End If
    Next Counter1
End Sub

Sub Trend(index)
    Dim ThisRuleNodes() As String
    Dim DummyString As String
    DummyString = RulesTwoDimension(index, 3)
    ThisRuleNodes = MakeSplitBase1(DummyString, ",")
    Dim NumNodenames As Integer
    NumNodenames = UBound(ThisRuleNodes)
    Dim NumbersToTrend() As Double
    ReDim NumbersToTrend(1 To lastRowParsed990 - 1, 1 To NumNodenames)
    Dim CounterRow As Long
    Dim CounterNode As Long
    Dim HoldColIndex As Integer
    ' Load data into NumbersToTrend
    For CounterNode = 1 To NumNodenames
        HoldColIndex = GetColumnIndex(wsParsed990, ThisRuleNodes(CounterNode))
        For CounterRow = 2 To lastRowParsed990
            NumbersToTrend(CounterRow - 1, CounterNode) = val(wsParsed990.Cells(CounterRow, HoldColIndex).value)
        Next CounterRow
    Next CounterNode
    Dim Up As Integer
    Dim Down As Integer
    For CounterRow = 1 To lastRowParsed990 - 1
        Up = 0: Down = 0
        For CounterNode = 1 To NumNodenames - 1
            If IsNumeric(NumbersToTrend(CounterRow, CounterNode)) And IsNumeric(NumbersToTrend(CounterRow, CounterNode + 1)) Then
                If NumbersToTrend(CounterRow, CounterNode + 1) > NumbersToTrend(CounterRow, CounterNode) Then
                    Up = Up + 1
                Else
                    Down = Down + 1
                End If
            End If
        Next CounterNode

        If Up > Down Then
            Scores(CounterRow) = 1
        Else
            Scores(CounterRow) = 0
        End If
    Next CounterRow
End Sub


Sub Percentile(index)
'Scores is for whole Module of type integer with dimension of NumRowsParsed990 -1
Dim ParsedColumn() As Variant
Dim numericVals() As Double
Dim RuleNodeName As String
Dim CutoffPercent As Single
ReDim ParsedColumn(1 To lastRowParsed990 - 1)
Dim Counter1 As Long
RuleNodeName = RulesTwoDimension(index, 3)
CutoffPercent = RulesTwoDimension(index, 4)
Dim NodenameColumnIndex As Integer
NodenameColumnIndex = GetColumnIndex(wsParsed990, RuleNodeName)
For Counter1 = 2 To lastRowParsed990
    ParsedColumn(Counter1 - 1) = wsParsed990.Cells(Counter1, NodenameColumnIndex)
Next Counter1
Dim CountNumberofNumbers As Long
CountNumberofNumbers = 0
    For Counter1 = 1 To lastRowParsed990 - 1
        If IsNumeric(ParsedColumn(Counter1)) And ParsedColumn(Counter1) <> 0 Then
            CountNumberofNumbers = CountNumberofNumbers + 1
            ReDim Preserve numericVals(CountNumberofNumbers)
            numericVals(CountNumberofNumbers) = CDbl(ParsedColumn(Counter1))
       End If
    Next Counter1
ReDim Preserve numericVals(1 To CountNumberofNumbers)
Dim CutoffActualNumber As Double
    ' Quicksort your array
    Call QuickSortDouble(numericVals, 1, CountNumberofNumbers)
   CutoffActualNumber = CustomPercentile(numericVals(), CutoffPercent)
'now that we have the cutoff actual number, go through Parsed Column and populate scores
For Counter1 = 1 To lastRowParsed990 - 1
    If IsNumeric(ParsedColumn(Counter1)) And _
        ParsedColumn(Counter1) > CutoffActualNumber Then
            Scores(Counter1) = 1
            Else
                Scores(Counter1) = 0
    End If
Next Counter1
End Sub

Function CustomPercentile(NumericValues() As Double, CutoffPercent As Single) As Single
Dim Position As Integer
    ' Compute rank position
    Position = Int(CutoffPercent * UBound(NumericValues()))
    CustomPercentile = NumericValues(Position)
End Function

Sub Eval(index)
    Dim ParsedColumn() As Variant
    ReDim ParsedColumn(1 To lastRowParsed990 - 1)
    Dim TxtOrNum As String
    Dim RuleNodeName As String
    Dim Expression As String
    Dim ExpressionToBeEvaluated As String
    Dim NodenameColumnIndex As Integer
    Dim Counter1 As Long
    RuleNodeName = RulesTwoDimension(index, 3)
    TxtOrNum = RulesTwoDimension(index, 4)
    Expression = RulesTwoDimension(index, 5)
    NodenameColumnIndex = GetColumnIndex(wsParsed990, RuleNodeName)
    
    ' Load column data from parsed worksheet
    For Counter1 = 2 To lastRowParsed990
        ParsedColumn(Counter1 - 1) = wsParsed990.Cells(Counter1, NodenameColumnIndex).value
    Next Counter1
    
    ' Evaluate each row against the rule
    For Counter1 = 1 To lastRowParsed990 - 1
        On Error GoTo EvaluationError
        Select Case TxtOrNum
            Case "Txt"
                ExpressionToBeEvaluated = Replace(Expression, RuleNodeName, """" & ParsedColumn(Counter1) & """")
                If Evaluate(ExpressionToBeEvaluated) Then
                    Scores(Counter1) = 1
                Else
                    Scores(Counter1) = 0
                End If
            Case "Num"
                If IsEmpty(ParsedColumn(Counter1)) Then
                    Scores(Counter1) = 0
                Else
                    ExpressionToBeEvaluated = Replace(Expression, RuleNodeName, ParsedColumn(Counter1))
                    If Evaluate(ExpressionToBeEvaluated) Then
                        Scores(Counter1) = 1
                    Else
                        Scores(Counter1) = 0
                    End If
                End If
            Case Else
                Scores(Counter1) = 0
        End Select
        On Error GoTo 0
        GoTo NextIteration
        
EvaluationError:
        Scores(Counter1) = 0
        Resume NextIteration
        
NextIteration:
    Next Counter1
End Sub



Sub CopyScorestoScored990(Scores() As Integer, ind)
'after Scores() has been filled, then need to write to worksheet
'find column
Dim ColumnIndex As Integer
Dim RuleNodeName As String
RuleNodeName = RulesTwoDimension(ind, 2)
ColumnIndex = GetColumnIndex(wsScored990, RuleNodeName)
Dim Counter1 As Long
For Counter1 = 1 To lastRowParsed990 - 1
       wsScored990.Cells(Counter1 + 1, ColumnIndex) = Scores(Counter1)
Next Counter1
'or to use Range operators as faster in Excel
'wsScored990.Range.Cells(Cells(2, ColumnIndex), _
'Cells(lastRowParsed990,ColumnIndex ).value  = Application.Transpose(Scores)
End Sub

' Function to read a file into a string
Function ReadFile(filePath As String) As String
    Dim fileNum As Integer, fileContent As String
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    ReadFile = fileContent
End Function


Function GetColumnIndex(ws As Worksheet, nodename As String) As Integer
    Dim colNum As Integer, cleanNode As String
    cleanNode = Trim(Replace(nodename, Chr(160), ""))
    For colNum = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If Trim(Replace(ws.Cells(1, colNum).value, Chr(160), "")) = cleanNode Then
            GetColumnIndex = colNum
            Exit Function
        End If
    Next colNum
    GetColumnIndex = 0 ' Return 0 if nodeName not found
End Function

Sub QuickSortDouble(arr() As Double, low As Long, high As Long)
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
        Call QuickSortDouble(arr, low, j)
        Call QuickSortDouble(arr, i, high)
    End If
End Sub


' Helper QuickSort
Sub QuickSort(arr() As Variant, low As Long, high As Long)
    Dim pivot As Double, i As Long, j As Long, temp As Double
    If low < high Then
        pivot = arr((low + high) \ 2)
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
        QuickSort arr, low, j
        QuickSort arr, i, high
    End If
End Sub

Function MakeSplitBase1(InputString As String, delim As String) As String()
    Dim r() As String, newR() As String
    Dim i As Long
    If Trim(InputString) = "" Then
        ReDim newR(1 To 1)
        newR(1) = ""
        MakeSplitBase1 = newR
        Exit Function
    End If
    r = Split(InputString, delim)
    ReDim newR(1 To UBound(r) + 1)
    For i = LBound(r) To UBound(r)
        newR(i + 1) = r(i)
    Next i
    MakeSplitBase1 = newR
End Function

Function BinarySearch(arr, value) As Long
    Dim low As Long, high As Long, mid As Long
    Dim temparr As String, tempvalue As String
    Dim tempInstr As Variant
    low = LBound(arr)
    high = UBound(arr)
    Do While low <= high
        mid = (low + high) \ 2
        temparr = Trim(arr(mid))
        tempvalue = Trim(value)
        If IsNull(temparr) Or IsNull(tempvalue) Then
            BinarySearch = 0
            Exit Function
        End If
        tempInstr = InStr(1, temparr, tempvalue, 1)
        If tempInstr > 0 Then
            BinarySearch = mid
            
Debug.Print "Found string in binary search at location " & mid
Debug.Print "Found string was " & temparr
            Exit Function
        End If
        If temparr < tempvalue Then low = mid + 1
        If temparr > tempvalue Then high = mid - 1
    Loop
    BinarySearch = -1 ' Value not found
End Function

Sub FinalizeScoredData990()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim total As Double

    Set ws = ThisWorkbook.Worksheets("Scored990Data")

    ' Identify last row and column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' Insert new column for TotalperForm
    ws.Cells(1, lastCol + 1).value = "TotalperForm"
    
    For i = 2 To lastRow
        total = 0
        For j = 2 To lastCol - 1 ' exclude unique ID and final score
            total = total + ws.Cells(i, j).value
        Next j
        ws.Cells(i, lastCol + 1).value = total
    Next i

    ' Sort descending by TotalperForm
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(lastCol + 1), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol + 1))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Add TotalPerRule row at bottom
    ws.Cells(lastRow + 1, 1).value = "TotalPerRule"
    For j = 2 To lastCol
        total = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, j), ws.Cells(lastRow, j)))
        ws.Cells(lastRow + 1, j).value = total
    Next j
End Sub

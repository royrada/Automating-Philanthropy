Attribute VB_Name = "Parse"
Sub ParseXML990Files()
Dim folderPath As String, filename As String
Dim xmlDoc As Object, ws As Worksheet
Dim dataRow As Long, configPath As String
Dim nodeDefs As Collection, nodeMap As Collection
Dim nodeInfo As Variant, nodePackage As Variant
Dim colMap As Object, colHeaders As Collection  ' <-- Right here
Dim colIndex As Long, uniqueName As String, nodeValue As String
Dim colName As String
    
    Set ws = ThisWorkbook.Sheets("Parsed990Data")
    ws.Cells.Clear
    ws.Cells(1, 1).value = "FormFile"

    folderPath = "C:\ALL\trusts\Form990\testforms\990\"
    configPath = ThisWorkbook.Path & "\nodenames.txt"
    Set nodeDefs = LoadNodeDefinitions(configPath)

    Set colMap = CreateObject("Scripting.Dictionary")
    
    Set colHeaders = New Collection
    
    Set nodeMap = New Collection
    colIndex = 2
    dataRow = 2

    ' Build headers and map nodeDefs to colNames
    For Each nodeInfo In nodeDefs
 '       uniqueName = GetUniqueColumnName(CStr(nodeInfo(2)), colMap)
 uniqueName = GetUniqueColumnName(CStr(nodeInfo(2)), colHeaders)

 
        colMap.Add uniqueName, colIndex
        ws.Cells(1, colIndex).value = uniqueName
        colIndex = colIndex + 1

        nodePackage = Array(nodeInfo, uniqueName)
        nodeMap.Add nodePackage
    Next nodeInfo

    ' Parse files
    filename = Dir(folderPath & "*.xml")
    Do While filename <> ""
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
        xmlDoc.async = False
        xmlDoc.Load folderPath & filename

        If xmlDoc.parseError.ErrorCode = 0 Then
'            ws.Cells(dataRow, 1).NumberFormat = "@"
'            ws.Cells(dataRow, 1).value = Replace(filename, "_public.xml", "")
ws.Cells(dataRow, 1).Errors(xlNumberAsText).Ignore = True
ws.Cells(dataRow, 1).value = "'" & Trim(CStr(Replace(filename, "_public.xml", "")))
'should suppress green arrow and make more manipulable in worksheet
            For Each nodePackage In nodeMap
                nodeInfo = nodePackage(0)
                colName = nodePackage(1)
                nodeValue = GetFormattedNodeValue(xmlDoc, nodeInfo)

                If colMap.exists(colName) Then
                    ws.Cells(dataRow, colMap(colName)).value = nodeValue
                Else
                    Debug.Print "Unexpected colName: " & colName
                End If
            Next nodePackage
            dataRow = dataRow + 1
        End If
        filename = Dir
    Loop

    MsgBox "Parsing complete!", vbInformation
End Sub

Function LoadNodeDefinitions(filePath As String) As Collection
    Dim fso As Object, ts As Object
    Dim lines As Collection, line As String
    Dim parts() As String, nodeInfo As Variant

    Set lines = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If line <> "" Then
            parts = Split(line, ";")
            If UBound(parts) = 2 Then
                nodeInfo = Array(Trim(parts(0)), CLng(parts(1)), Trim(parts(2)))
                lines.Add nodeInfo
            End If
        End If
    Loop
    ts.Close
    Set LoadNodeDefinitions = lines
End Function

Function GetFormattedNodeValue(xmlDoc As Object, nodeInfo As Variant) As String
    Dim formatType As String, maxLen As Long, xpath As String
    Dim rawVal As String, cleanVal As String
    Dim xmlNode As Object

    formatType = nodeInfo(0)
    maxLen = nodeInfo(1)
    xpath = nodeInfo(2)

    Set xmlNode = xmlDoc.SelectSingleNode(xpath)
    
'    If xmlNode Is Nothing Then
'    Debug.Print "Missing node: " & xpath
'End If

    
    If xmlNode Is Nothing Then
        GetFormattedNodeValue = ""
        Exit Function
    End If

    rawVal = Trim(xmlNode.text)
    Select Case UCase(formatType)
        Case "STRING"
            cleanVal = FilterToLegitChars(rawVal)
            If Len(cleanVal) > maxLen Then cleanVal = Left(cleanVal, maxLen)
            GetFormattedNodeValue = cleanVal
        Case "DATE"
            If IsDate(rawVal) Then
                cleanVal = Format(CDate(rawVal), "yyyy/mm/dd")
                GetFormattedNodeValue = cleanVal
            Else
                GetFormattedNodeValue = ""
            End If
        Case "INTEGER"
            If IsNumeric(rawVal) Then
                GetFormattedNodeValue = CDbl(rawVal)
            Else
                GetFormattedNodeValue = ""
            End If
        Case "ABSINT"
            If IsNumeric(rawVal) Then
                GetFormattedNodeValue = Abs(CDbl(rawVal))
            Else
                GetFormattedNodeValue = ""
            End If
        Case Else
            GetFormattedNodeValue = ""
    End Select
End Function


Function GetUniqueColumnName(xpath As String, colHeaders As Collection) As String
    Dim parts() As String, colName As String
    Dim item As Variant, exists As Boolean

    parts = Split(xpath, "/")
    colName = parts(UBound(parts) - 1) & "_" & parts(UBound(parts)) ' e.g., CYMinus1YrEndwmtFundGrp_GrantsOrScholarshipsAmt

    exists = False
    For Each item In colHeaders
        If item = colName Then
            exists = True
            Exit For
        End If
    Next

    If Not exists Then colHeaders.Add colName
    GetUniqueColumnName = colName
End Function


Private Function FilterToLegitChars(val As String) As String
    Dim i As Long, ch As String, result As String
    For i = 1 To Len(val)
        ch = mid(val, i, 1)
        If ch Like "[A-Za-z0-9 /\\'"".,\-]" Then
            result = result & ch
        Else
            result = result & " "
        End If
    Next i
    FilterToLegitChars = result
End Function




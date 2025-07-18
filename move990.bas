Attribute VB_Name = "move990"
Sub Move990Files()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim xmlDoc As Object
    Dim returnTypeNode As Object
    Dim returnType As String
    Dim sourcePath As String
    Dim destPath As String
    Dim errantPath As String
    Dim errantCount As Long
    
    sourcePath = "C:\ALL\trusts\Form990\testforms\"
    destPath = sourcePath & "990\"
    errantPath = sourcePath & "errant\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(sourcePath)
    If Not fso.FolderExists(errantPath) Then fso.CreateFolder errantPath
    
    For Each file In folder.Files
        If Right(file.Name, 3) = "xml" Then
            Set xmlDoc = CreateObject("MSXML2.DOMDocument")
            xmlDoc.async = False
            xmlDoc.Load file.Path
            If xmlDoc.parseError.ErrorCode = 0 Then
                Set returnTypeNode = xmlDoc.SelectSingleNode("//ReturnTypeCd")
                If Not returnTypeNode Is Nothing Then
                    returnType = returnTypeNode.text
                    If returnType = "990" Then
                        fso.MoveFile file.Path, destPath & file.Name
                    End If
                Else
                    fso.MoveFile file.Path, errantPath & file.Name
                    errantCount = errantCount + 1
                End If
            Else
                fso.MoveFile file.Path, errantPath & file.Name
                errantCount = errantCount + 1
            End If
        End If
    Next
    
    MsgBox "Files moved successfully! " & vbCrLf & _
           "Errant files moved: " & errantCount, vbInformation
End Sub



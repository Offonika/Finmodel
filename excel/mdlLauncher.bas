Attribute VB_Name = "mdlLauncher"
Option Explicit

Private Function ReadInterpreter(confPath As String) As String
    Dim fso As Object, ts As Object, line As String
    Dim projectPath As String, value As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(confPath) Then Exit Function
    Set ts = fso.OpenTextFile(confPath, 1)
    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If UCase$(Left$(line, 12)) = "PROJECT_PATH" Then
            value = Split(line, "=")(1)
            projectPath = Trim(value)
        ElseIf UCase$(Left$(line, 11)) = "INTERPRETER" Then
            value = Split(line, "=")(1)
            value = Trim(value)
            If projectPath <> "" Then
                value = Replace(value, "%(PROJECT_PATH)s", projectPath)
            End If
            ReadInterpreter = value
            Exit Do
        End If
    Loop
    ts.Close
End Function

Public Sub RunTask()
    Dim confPath As String
    confPath = ThisWorkbook.Path & "\.xlwings.conf"
    Dim interp As String
    interp = ReadInterpreter(confPath)
    If interp <> "" Then
        If Dir(interp) = "" Then
            MsgBox "Python interpreter not found. Run setup again or update .xlwings.conf.", vbExclamation
            Exit Sub
        End If
    End If

    ' Call Python task
    RunPython "import xlwings_macro; xlwings_macro.run_aggregation()"
End Sub

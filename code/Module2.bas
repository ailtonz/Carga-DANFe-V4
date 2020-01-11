Attribute VB_Name = "Module2"
'Function GetExecutablePath(strFileType As String) As String
'' returns the full path to the executable associated with the given file type
'Dim strFileName As String, f As Integer, strExecutable As String, r As Long
'    If Len(strFileType) = 0 Then Exit Function ' no file type
'    strFileName = String$(255, " ")
'    strExecutable = String$(255, " ")
'    GetTempFileName CurDir, "", 0&, strFileName ' get a temporary file name
'    strFileName = Application.Trim(strFileName)
'    strFileName = Left$(strFileName, Len(strFileName) - 3) & strFileType ' add the given file type
'    f = FreeFile
'    Open strFileName For Output As #f ' create the temporary file
'    Close #f
'    r = FindExecutable(strFileName, vbNullString, strExecutable) ' look for an associated executable
'    Kill strFileName ' remove the temporary file
'    If r > 32 Then ' associated executable found
'        strExecutable = Left$(strExecutable, InStr(strExecutable, Chr(0)) - 1)
'    Else ' no associated executable found
'        strExecutable = vbNullString
'    End If
'    GetExecutablePath = strExecutable
'End Function
'
'
'Sub OpenPDFDocument()
'Dim strDocument As String, strExecutable As String
'    strDocument = Application.GetOpenFilename("PDF Files,*.pdf,All Files,*.*", 1, "Open File", , False) ' get pdf document name
'    If Len(strDocument) < 6 Then Exit Sub
'    strExecutable = GetExecutablePath("pdf") ' get the path to Acrobat Reader
'    If Len(strExecutable) > 0 Then
'        Shell strExecutable & " " & strDocument, vbMaximizedFocus ' open pdf document
'    End If
'End Sub

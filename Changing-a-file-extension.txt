Attribute VB_Name = "Module4"
'This cool tip changes the extension of a file name passed to it, and returns it quickly and easily. To call the function use:

MsgBox ChgFileExt("url.doc", "txt")
'This would return url.txt

'Copy this function into a module:

Function ChgFileExt(sFile As String, sNewExt As String) As String
    Dim lRet As Long
        If sFile = "" Then Exit Function
        If lRet = 0 Then lRet = Len(sFile) + 1
    ChgFileExt = Left$(sFile, lRet - 1) & "." & sNewExt
End Function
 


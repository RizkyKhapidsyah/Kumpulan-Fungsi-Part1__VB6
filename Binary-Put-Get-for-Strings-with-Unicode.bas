Attribute VB_Name = "Module1"
Function GetString(ByVal Filenumber As Integer, _
   ByVal Lng As Boolean) As String
    Dim StrLengthLng As Long
    Dim StrLengthInt As Integer
    Dim StrLength As Long

    If Lng Then
        Get #Filenumber, , StrLengthLng
        StrLength = StrLengthLng
    Else
        Get #Filenumber, , StrLengthInt
        StrLength = StrLengthInt
    End If
    GetString = String$(StrLength, " ")
    Get #Filenumber, , GetString
End Function

Sub PutString(ByVal Filenumber As Integer, Strng As String, _
   ByVal Lng As Boolean)
    If Lng Then
        Put #Filenumber, , CLng(Len(Strng))
    Else
        Put #Filenumber, , CInt(Len(Strng))
    End If
    Put #Filenumber, , Strng
End Sub

Function GetStringU(ByVal Filenumber As Integer, _
      ByVal Lng As Boolean) As String
    Dim StrLengthLng As Long
    Dim StrLengthInt As Integer
    Dim StrLength As Long

    If Lng Then
        Get #Filenumber, , StrLengthLng
        StrLength = StrLengthLng
    Else
        Get #Filenumber, , StrLengthInt
        StrLength = StrLengthInt
    End If
    
    If StrLength = 0 Then
        GetStringU = ""
    Else
        ReDim rwert(StrLength * 2 - 1) As Byte
        Get #Filenumber, , rwert
        GetStringU = rwert
    End If
End Function

Sub PutStringU(ByVal Filenumber As Integer, _
   Strng As String, ByVal Lng As Boolean)
    If Lng Then
        Put #Filenumber, , CLng(Len(Strng))
    Else
        Put #Filenumber, , CInt(Len(Strng))
    End If
    Dim b() As Byte
    b = Strng
    Put #Filenumber, , b
End Sub

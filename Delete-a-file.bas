Attribute VB_Name = "Module7"
Private Sub cmdDelete_Click()
Dim strPath As String
        strPath = InputBox$("Enter file path:")
        Kill strPath
End Sub

 


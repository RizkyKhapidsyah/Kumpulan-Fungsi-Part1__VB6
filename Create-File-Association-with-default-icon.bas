Attribute VB_Name = "Module6"
'// You may use this code all you want on the condition you keep this simple comment
'// Anyone who improves the code please let me know.
'// Date     : 21/1/2000
'// Author   : Damien McGivern
'// E-Mail   : Damien@Dingo-Delights.co.uk
'// Web Site : www.dingo-delights.co.uk
'// Purpose  : To create file associations with default icons

'// Improved 23/1/200 - New parameters 'Switch', 'PromptOnError', better error handling

'// Parameters
'// Required    Extension       (Str) ie ".exe"
'// Required    FileType        (Str) ie "VB.Form"
'// Required    FileTYpeName    (Str) ie. "Visual Basic Form"
'// Required    Action          (Str) ie. "Open" or "Edit"
'// Required    AppPath         (Str) ie. "C:\Myapp"
'// Optional    Switch          (Str) ie. "/u"                  Default = ""
'// Optional    SetIcon         (Bol)                           Default = False
'// Optional    DefaultIcon     (Str) ie. "C:\Myapp,0"
'// Optional    PromptOnError   (Bol)                           Default = False

'// HOW IT WORKS
'// Extension(Str)   Default = FileType(Str)

'// FileType(Str)    Default = FileTypeName(Str)
'// "DefaultIcon"     Default = DefaultIcon(Str)
'// "shell"
'// Action(Str)
'// "command"   Default = AppPath(Str) & switch(Str) & " %1"

Option Explicit

Private Const REG_SZ As Long = 1

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const ERROR_SUCCESS = 0
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private PromptOnErr As Boolean

Private Declare Function RegCloseKey Lib "advapi32.dll" _
                                    (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
                                    Alias "RegCreateKeyExA" _
                                    (ByVal hKey As Long, _
                                    ByVal lpSubKey As String, _
                                    ByVal Reserved As Long, _
                                    ByVal lpClass As String, _
                                    ByVal dwOptions As Long, _
                                    ByVal samDesired As Long, _
                                    ByVal lpSecurityAttributes As Long, _
                                    phkResult As Long, _
                                    lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
                                    Alias "RegOpenKeyExA" _
                                    (ByVal hKey As Long, _
                                    ByVal lpSubKey As String, _
                                    ByVal ulOptions As Long, _
                                    ByVal samDesired As Long, _
                                    phkResult As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" _
                                    Alias "RegSetValueExA" _
                                    (ByVal hKey As Long, _
                                    ByVal lpValueName As String, _
                                    ByVal Reserved As Long, _
                                    ByVal dwType As Long, _
                                    ByVal lpValue As String, _
                                    ByVal cbData As Long) As Long


Public Function CreateFileAss(Extension As String, _
                              FileType As String, _
                              FileTypeName As String, _
                              Action As String, _
                              AppPath As String, _
                              Optional Switch As String = "", _
                              Optional SetIcon As Boolean = False, _
                              Optional DefaultIcon As String, _
                              Optional PromptOnError As Boolean = False) _
                              As Boolean

    On Error GoTo ErrorHandler:

    PromptOnErr = PromptOnError

    '// Check that AppPath exists.
    If Dir(AppPath, vbNormal) = "" Then
        If PromptOnError Then MsgBox "The application path '" & _
            AppPath & "' cannot be found.", _
            vbCritical + vbOKOnly, "DLL/OCX Register"

        CreateFileAss = False
        Exit Function
    End If

    Dim ERROR_CHARS As String: ERROR_CHARS = "\/:*?<>|" & Chr(34)
    Dim i As Integer

    If Asc(Extension) <> 46 Then Extension = "." & Extension
    '// Check extension has "." at front

    '// Check for invalid chars within extension
    For i = 1 To Len(Extension)
        If InStr(1, ERROR_CHARS, Mid(Extension, i, 1), vbTextCompare) Then
            If PromptOnError Then MsgBox "The file extension '" _
                  & Extension & "' contains an illegal char (\/:*?<>|" _
                  & Chr(34) & ").", vbCritical + vbOKOnly, "DLL/OCX Register"

            CreateFileAss = False
            Exit Function
        End If
    Next

    If Switch <> "" Then Switch = " " & Trim(Switch
   Action = FileType & "\shell\" & Action & "\command"

    Call CreateSubKey(HKEY_CLASSES_ROOT, Extension)        '// Create .xxx key
    Call CreateSubKey(HKEY_CLASSES_ROOT, Action)           '// Create action key

    If SetIcon Then
        Call CreateSubKey(HKEY_CLASSES_ROOT, (FileType _
            & "\DefaultIcon"))    '// Create default icon key

        If DefaultIcon = "" Then
            '// This line of code sets the application's own icon as the _
             default file icon

            Call SetKeyDefault(HKEY_CLASSES_ROOT, _
               FileType & "\DefaultIcon", Trim(AppPath & ",0"))

        Else
            Call SetKeyDefault(HKEY_CLASSES_ROOT, FileType _
                 & "\DefaultIcon", Trim(DefaultIcon))

        End If
    End I
   Call SetKeyDefault(HKEY_CLASSES_ROOT, Extension, _
                FileType)                  '// Set .xxx key default

    Call SetKeyDefault(HKEY_CLASSES_ROOT, FileType, _
                FileTypeName)               '// Set file type default

    Call SetKeyDefault(HKEY_CLASSES_ROOT, Action, _
                AppPath & Switch & " %1")     '// Set Command line

    CreateFileAss = True
    Exit Function

ErrorHandler:

    If PromptOnError Then MsgBox "An error occured while _
          attempting to create the file extension '" _
          & Extension & "'.", vbCritical + vbOKOnly, "DLL/OCX Register"

    CreateFileAss = False

End Function

Private Function CreateSubKey(RootKey As Long, NewKey As String) As Boolean
    '// This function creates a new sub key
    Dim hKey As Long, regReply As Long
    regReply = RegCreateKeyEx(RootKey, NewKey, _
         0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, 0&)

    If regReply <> ERROR_SUCCESS Then
        If PromptOnErr Then MsgBox "An error occured _
         while attempting to to create a registery key.", vbCritical + _
         vbOKOnly, "DLL/OCX Register"

        CreateSubKey = False
    Else
        CreateSubKey = True
    End If

    Call RegCloseKey(hKey)
End Function


Private Function SetKeyDefault(RootKey As Long, Address As String, Value As String) As Boolean
    '// This function sets the default vaule of the key which is always a string
    Dim regReply As Long, hKey As Long
    regReply = RegOpenKeyEx(RootKey, Address, 0, KEY_ALL_ACCESS, hKey)

    If regReply <> ERROR_SUCCESS Then
        If PromptOnErr Then MsgBox "An error occured while attempting _
          to access the registery.", vbCritical + vbOKOnly, "DLL/OCX Register"
        SetKeyDefault = False
        Exit Function
    End If

    Value = Value & Chr(0)

    regReply = RegSetValueExString(hKey, "", 0&, REG_SZ, Value, Len(Value))

    If regReply <> ERROR_SUCCESS Then
        If PromptOnErr Then MsgBox "An error occured while _
           attempting to set key default value.", vbCritical + _
           vbOKOnly, "DLL/OCX Register"

        SetKeyDefault = False
    Else
        SetKeyDefault = True
    End If

    Call RegCloseKey(hKey)
End Function


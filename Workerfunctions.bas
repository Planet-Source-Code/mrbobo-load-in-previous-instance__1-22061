Attribute VB_Name = "Workerfunctions"
'These functions are just for this demo - they are
'not part of the Load In Previous Instance method

'Used to check if a file exists
Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
'Registry functions for association purposes
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Force icons to update after association
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const HKEY_CLASSES_ROOT = &H80000000
Const SHCNE_ASSOCCHANGED = &H8000000
Const SHCNF_IDLIST = &H0&
'Used to load the 'Save File' commondialog
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim OFName As OPENFILENAME

Public Sub Associate()
'Write the required data to registry to associate "bbt" files
'with this application
Dim RKey As Long
Dim result As Long
Dim strData As String
result = RegCreateKey(HKEY_CLASSES_ROOT, "BoboFile\DefaultIcon", RKey)
strData = App.Path + "\" + App.EXEName + ".exe,0"
result = RegSetValueEx(RKey, "", 0, 1, ByVal strData, Len(strData))
strData = App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"
result = RegCreateKey(HKEY_CLASSES_ROOT, "BoboFile\Shell\Open\Command", RKey)
strData = App.Path + "\" + App.EXEName + ".exe" + " %1"
result = RegSetValueEx(RKey, "", 0, 1, ByVal strData, Len(strData))
result = RegCloseKey(RKey)
result = RegCreateKey(HKEY_CLASSES_ROOT, "BoboFile", RKey)
strData = "Bobo Code Viewer File"
result = RegSetValueEx(RKey, "", 0, 1, ByVal strData, Len(strData))
result = RegCloseKey(RKey)
result = RegCreateKey(HKEY_CLASSES_ROOT, ".bbt", RKey)
strData = "BoboFile"
result = RegSetValueEx(RKey, "", 0, 1, ByVal strData, Len(strData))
result = RegCloseKey(RKey)
'A lot of 'Associate' functions fail to include this call
'It forces the icons of the associated filetype to change
'according to the association. It informs Windows the registry
'has changed and it should do an update.
LockWindowUpdate GetDesktopWindow
SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
LockWindowUpdate 0
End Sub
Public Sub DisAssociate()
'DisAssociation is relatively easy. Just delete
'the keys created in association
Dim lRegResult As Long
lRegResult = RegDeleteKey(HKEY_CLASSES_ROOT, "BoboFile")
lRegResult = RegDeleteKey(HKEY_CLASSES_ROOT, ".bbt")
End Sub
Public Function FileExists(Path As String) As Boolean
'Check for fileexistance
Dim RETval As Integer
RETval = SHFileExists(Path)
If RETval <> 0 Then
    FileExists = True
Else
    FileExists = False
End If
End Function
Public Sub FileSave(Text As String, sFilename As String)
'Save a plain text file
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open sFilename For Output As #f
        Print #f, Text
    Close #f
End Sub

Public Function OneGulp(sFilename As String) As String
'Open a plain text file
    Dim temp As String
    Dim f As Integer
    f = FreeFile
    Open sFilename For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get f, , temp
    Close #f
    OneGulp = temp
End Function
Public Function ShowSave() As String
'Show a commondialog to save a file
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Form1.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Bobo test filetype (*.bbt)" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = "Save a Bobo test file"
    OFName.flags = 5
    If GetSaveFileName(OFName) Then
        ShowSave = StripTerminator(Trim$(OFName.lpstrFile))
    Else
        ShowSave = ""
    End If
End Function

Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
'Used to add the correct extension to a filename
'when saving a file
Dim temp As String
If newext <> "" Then newext = "." + newext
temp = Mid$(filepath, 1, InStrRev(filepath, "."))
If temp <> "" Then
    temp = Left(temp, Len(temp) - 1)
    ChangeExt = temp + newext
Else
    ChangeExt = filepath + newext
End If
End Function
Public Function StripTerminator(ByVal strString As String) As String
'This function comes from MS itself !
'API often returns strings with a trailing nullstring
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function


Public Function txtHow() As String
'Couldn't be bothered with a RES file or the like
'so I just stuck the strings here
txtHow = "How does it work ?" + vbCrLf + _
"Requirements : On the main form of your application include a textbox" + vbCrLf + _
"called txtCommand. On the Form_Load event the window handle of" + vbCrLf + _
"txtCommand is saved to registry using the Savesetting method. If the" + vbCrLf + _
"application was shelled through Explorer by the user clicking on a" + vbCrLf + _
"file associated to your application, a check for previous instance" + vbCrLf + _
"is made. If true the name of the shelled file is written to the window" + vbCrLf + _
"handle recorded in registry(txtcommand.hwnd) using the Sendmessage" + vbCrLf + _
"API. The txtCommand_Change event triggers the loading of the file." + vbCrLf + vbCrLf + _
"Click on the View Code button to see the code required."
End Function

Public Function txtCode() As String
txtCode = "Private Declare Function SendMessage Lib " + Chr(34) + "user32" + Chr(34) + " Alias " + Chr(34) + "SendMessageA" + Chr(34) + " (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long" + vbCrLf + _
"Private Const WM_SETTEXT = &HC" + vbCrLf + _
"Private Sub Form_Load()" + vbCrLf + _
"Dim mycommand As String" + vbCrLf + _
"mycommand = Command()" + vbCrLf + _
"If mycommand <> " + Chr(34) + Chr(34) + " Then" + vbCrLf + _
"    If App.PrevInstance Then" + vbCrLf + _
"        LoadinPrevInst mycommand" + vbCrLf + _
"        End" + vbCrLf + _
"    Else" + vbCrLf + _
"        Text1.Text = OneGulp(mycommand)" + vbCrLf + _
"    End If" + vbCrLf + _
"End If" + vbCrLf + _
"SaveSetting App.Title, " + Chr(34) + "ActiveWindow" + Chr(34) + ", " + Chr(34) + "Handle" + Chr(34) + ", Str(txtCommand.hwnd)" + vbCrLf + _
"End Sub" + vbCrLf + _
"Private Sub LoadinPrevInst(mfile As String)" + vbCrLf + _
"Dim temp As String, mhw As Long" + vbCrLf + _
"    temp = GetSetting(App.Title, " + Chr(34) + "ActiveWindow" + Chr(34) + ", " + Chr(34) + "Handle" + Chr(34) + ")" + vbCrLf + _
"    mhw = CLng(Val(temp))" + vbCrLf + _
"    SendMessage mhw, WM_SETTEXT, 0, ByVal CStr(mfile)" + vbCrLf + _
"End Sub" + vbCrLf + _
"Private Sub txtCommand_Change()" + vbCrLf + _
"If FileExists(txtCommand.Text) Then Text1.Text = OneGulp(txtCommand.Text)" + vbCrLf + _
"End Sub"

End Function

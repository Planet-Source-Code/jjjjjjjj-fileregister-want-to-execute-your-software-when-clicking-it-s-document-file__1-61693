Attribute VB_Name = "mdlFileRegister"

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Module    : mdlFileRegister
' Auther    : Jim Jose
' Credits   : James Crowley for the basics
' Purpose   : Associate a document file with an Application sothat
'           : the explorer can execute the ApplicationPath when
'           : the user clicks on the document file on the explorer(Folder)
' License   : You are free to use this code in any of ur softwares..
'           : But please don't change the Auther/Credits
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Special Note : If you are registering a NEW FILE-TYPE.. then the document file
' icon will be loaded instently!!!. If you are trying to RE-ASSIGN the icon for
' existing FILE-TYPE(use mOverWrite=True for this), then you must restart
' your system to take any effect on document icon!!!
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

'--Registry windows api calls
Private Declare Function RegCreateKey& Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, lphKey As Long)
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal fdwType As Long, ByVal lpszValue As String, ByVal dwLength As Long)
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long

'--Required constants
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 256&
Private Const REG_SZ = 1

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : RegisterFile
' Auther    : Jim Jose
' Syntax    : Call RegisterFile(".tdb", "Text DataBase File", App.EXEName, App.Path & "\" & App.EXEName)
' Return    : True if done!
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function RegisterFile(ByVal mFileExt As String, _
                            ByVal mFileDes As String, _
                            ByVal mExeName As String, _
                            ByVal mExeFilePath As String, _
                            Optional ByVal mIconFile As String = "App", _
                            Optional ByVal mOverWrite As Boolean = False, _
                            Optional ByVal mCmdSwitch As String = " %1") As Boolean
                            
' /mFileExt     : Extension for the document file
' /mFileDes     : The file description
' /mExeName     : Name of application to be executed
' /mExeFilePath : Application file path
' /mIconFile    : Icon source for the document file(default value is the exe itself)
' /mOverWrite   : If the filetype is already registered... OverWrite???
' /mCmdSwitch   : Optional parameter for command line switches

 Dim sKeyName  As String                                '--Holds Key Name in registry.
 Dim sKeyValue As String                                '--Holds Key Value in registry.
 Dim lphKey    As Long                                  '--Holds created key handle from RegCreateKey.
 Dim lpIconCmd As String
 Const lpszSubKey  As String = "shell\open\command"     '--Holds The subkey path
 Const lpszIconKey As String = "DefaultIcon"            '--Holds The subkey for default icon

    '--Setting the optonal icon file
    If mIconFile = "App" Then mIconFile = mExeFilePath
        
    '--Filtering some common errors
    mExeFilePath = Replace(mExeFilePath, "\\", "\")
    If mExeFilePath = vbNullString Or mFileExt = vbNullString Then Exit Function
    If Not Left$(mFileExt, 1) = "." Then mFileExt = "." & mFileExt
    If FileExists(mExeFilePath) = False Then Err.Raise 404, , "Executable file not found!!!": Exit Function
    If FileExists(mIconFile) = False Then Err.Raise 404, , "Icon file not found!!!": Exit Function
    If Not StrComp(Right$(mExeFilePath, 3), "exe", vbTextCompare) = 0 Then Err.Raise 404, , "The given executable not a '*.exe' file!!!": Exit Function
    If StrComp(Right$(mIconFile, 3), "exe", vbTextCompare) = 0 Then lpIconCmd = ",0"
    
    ' Check Overwrite!!!
    If mOverWrite = False Then
        Call RegOpenKey(HKEY_CLASSES_ROOT, mExeName, lphKey)
        If Not lphKey = 0 Then Err.Raise 1, , "File Type is already registered!!!": Exit Function
    End If
    
    On Error GoTo Handle
    '--This creates a Root entry for the Application
    sKeyName = mExeName
    sKeyValue = mFileDes
    Call RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    Call RegSetValue&(lphKey&, Empty, REG_SZ, sKeyValue, 0&)

    '--This creates a Root entry for
    '--Extesion associated with the Application
    sKeyName = mFileExt
    sKeyValue = mExeName
    Call RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    Call RegSetValue&(lphKey, Empty, REG_SZ, sKeyValue, 0&)

    '--This sets the command line for the Application
    sKeyName = mExeName
    sKeyValue = mExeFilePath & mCmdSwitch
    Call RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    Call RegSetValue&(lphKey, lpszSubKey, REG_SZ, sKeyValue, MAX_PATH)

    '--This sets the default icon for the file
    sKeyName = mExeName
    sKeyValue = mIconFile & lpIconCmd
    Call RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    Call RegSetValue&(lphKey, lpszIconKey, REG_SZ, sKeyValue, MAX_PATH)

    '--Return Success
    RegisterFile = True

Exit Function
Handle:
    '--Return Failure
    RegisterFile = False

End Function

' Checks the existance of a file
Private Function FileExists(sFile As String) As Boolean
On Error GoTo Check
    If Trim(sFile) = "" Then
                FileExists = False
                Exit Function
    End If
    If Dir$(sFile, vbNormal) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
Exit Function
Check:
    FileExists = False
End Function

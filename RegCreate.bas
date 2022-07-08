Attribute VB_Name = "RegCreate"
'/////////////////////////RegCreate.bas\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////Created by Ziegs\\\\\\\\\\\\\\\\\\\\\\\\\\\
'This is Copyright 2000-2001 Ziegs2001 Inc.
'http://www.envy.nu/ziegs
'http://www.envy.nu/syndesigns
'MattzStar@aol.com

'// Registry api calls
Private Declare Function RegCreateKey& Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, lphKey As Long)
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal fdwType As Long, ByVal lpszValue As String, ByVal dwLength As Long)
'// Required constants
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 256&
Private Const REG_SZ = 1

'// procedure you call to associate the extension with your program.
Sub MakeDefault(strExt As String, strApp As String, strFileDisc As String)
    Dim sKeyName  As String  '// Holds Key Name in registry.
    Dim sKeyValue As String  '// Holds Key Value in registry.
    Dim ret       As Long    '// Holds error status if any from API calls.
    Dim lphKey    As Long    '// Holds created key handle from RegCreateKey.
   
    '// This creates a Root entry called whatever you need,
    '// for example a program called "TextEdit2000"
    sKeyName = strApp '// Application Name, for Example "TextEdit2000"
    sKeyValue = strFileDisc '// File Type Description, for example "TextEdit Document"
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey&, Empty, REG_SZ, sKeyValue, 0&)

    '// This creates a Root entry for your extension, associated with your program name.
    sKeyName = strExt '// File Extension, for Example "txe"
    sKeyValue = strApp '// Application Name, for Example "TextEdit2000"
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)

    ret = RegSetValue&(lphKey, Empty, REG_SZ, sKeyValue, 0&)

    '//This sets the command line for the Program.
    sKeyName = strApp '// Application Name, for Example "TextEdit2000"
    If App.Path Like "*\" Then
        sKeyValue = App.Path & App.EXEName & ".exe %1" '// Application Path, do not append!
    Else
        sKeyValue = App.Path & "\" & App.EXEName & ".exe %1" '// Application Path, do not append!
    End If
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
End Sub

Function rtfCommand(rtfBox As RichTextBox)
If Command <> "" Then
rtfBox.LoadFile Command
End If
End Function

Function txtCommand(txtBox As TextBox)
If Command <> "" Then
txtBox.Text = Command
End If
End Function

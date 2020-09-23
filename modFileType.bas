Attribute VB_Name = "modFileType"
'FileType Manager (modFileType.bas)
'======================================
'Copyright (c) Jan Botha 1998 - 2000
'eMail ~> c03jabot@prg.wcape.school.za
'Credits ~> Randy Mcdowell
'
'With these functions you can do all the basic
' management of a file type. It is not really
' commented or anything as I thought you would
' be able to figure out what it does.
'
'Note that you need to add the module named
' RegistryAccess (Winreg32.bas) to your project,
' as my functions use it to access the registry.
' Also to note, is that I did not write the
' RegistryAccess module. It was written by
' Randy Mcdowell.
'
'I programmed all the functions to return a
' Boolean value when executed. If a function
' returns False it means that an error has
' occurred.
'
'Note that when you specify an icon path, you
' must also append ",0" or ",3" to the path
' (without the quotes). This is the index number
' of the icon to use inside that file.
'
'Thank you for using my code. If you modify
' this code and/or use it in your application
' please send me a copy of your application or
' the modified module. Also feel free to drop
' me a comment via email.
'==================================

'Sample Call:
'   MakeFileType "txt", "Text Document", "C:\windows\notepad.exe,0", "open", "C:\windows\notepad.exe %1"

Public Function MakeFileType(ByVal Extension As String, ByVal NameOfType As String, ByVal DefaultIcon As String, ByVal NameOfAction As String, ByVal ActionPath As String) As Boolean
    On Error GoTo Oops
    Dim dotExtension As String, Extensionfile As String
    Dim correctNameOfAction As String
    dotExtension = "." & Extension
    Extensionfile = Extension & "file"
    correctNameOfAction = ReplaceChars(NameOfAction, " ", "_")
    
    CreateKey HKEY_CLASSES_ROOT, dotExtension
    CreateKey HKEY_CLASSES_ROOT, Extensionfile
    CreateKey HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon"
    CreateKey HKEY_CLASSES_ROOT, Extensionfile, "Shell"
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell", correctNameOfAction
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command"
        
    SaveString HKEY_CLASSES_ROOT, dotExtension, "", "", Extensionfile
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "", "", NameOfType
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon", "", DefaultIcon
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "Shell", "", correctNameOfAction
    SaveString HKEY_CLASSES_ROOT, Extensionfile & "\Shell", correctNameOfAction, "", "&" & NameOfAction
    SaveString HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command", "", ActionPath
    
    MakeFileType = True
    Exit Function
Oops:
    MakeFileType = False
    Exit Function
    Resume Next
End Function

'Sample Call:
'    RemoveFileType "txt"
    
Public Function RemoveFileType(ByVal Extension As String)
    On Error GoTo Oopsie
    
    DeleteKey HKEY_CLASSES_ROOT, "." & Extension
    DeleteKey HKEY_CLASSES_ROOT, Extension & "file"
    RemoveFileType = True
    Exit Function

Oopsie:
    RemoveFileType = False
    Exit Function
    Resume Next
End Function

'Sample call:
'    EditActionPath "txt", "open", "c:\program files\accessories\wordpad.exe"
    
Public Function EditActionPath(ByVal Extension As String, ByVal NameOfAction As String, ByVal NewPath As String) As Boolean
    On Error GoTo AnotherOops
    Dim Extensionfile As String, correctNameOfAction As String
    Extensionfile = Extension & "file"
    correctNameOfAction = ReplaceChars(NameOfAction, " ", "_")
    
    SaveString HKEY_CLASSES_ROOT, Extensionfile & "\shell\" & correctNameOfAction, "command", "", NewPath
    EditAction = True
    Exit Function
AnotherOops:
    EditActionPath = False
    Exit Function
    Resume Next
End Function

'Sample call:
'    AddAction "txt", "open with Wordpad", "C:\program files\accessories\wordpad.exe", True
Public Function AddAction(ByVal Extension As String, ByVal NameOfAction As String, ByVal ActionPath As String, Optional ByVal SetAsDefault As Boolean) As Boolean
    On Error GoTo OopsAgain
    Dim dotExtension As String, Extensionfile As String
    Dim Replaced As String
    
    Replaced = ReplaceChars(NameOfAction, " ", "_")
    dotExtension = "." & Extension
    Extensionfile = Extension & "file"
    
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell", Replaced
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & Replaced, "command"
    SaveString HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & Replaced, "command", "", ActionPath
    SaveString HKEY_CLASSES_ROOT, Extensionfile & "\Shell", Replaced, "", "&" & NameOfAction
    If Not IsMissing(SetAsDefault) Then
        If SetAsDefault = True Then
            SaveString HKEY_CLASSES_ROOT, Extensionfile, "Shell", "", Replaced
        End If
    End If
    AddAction = True
    Exit Function
OopsAgain:
    AddAction = False
    Beep
    Exit Function
    Resume Next
End Function

'Sample call:
'    EditDefaultIcon "txt", "C:\windows\calc.exe,0"
    
Public Function EditDefaultIcon(ByVal Extension As String, ByVal NewIconPath As String) As Boolean
    On Error GoTo IconOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon", "", NewIconPath
    EditDefaultIcon = True
    Exit Function

IconOops:
    EditDefaultIcon = False
    Exit Function
    Resume Next
End Function

'Sample call:
'    RemoveAction "txt", "open with Wordpad"

Public Function RemoveAction(ByVal Extension As String, ByVal NameOfAction As String) As Boolean
    On Error GoTo RemoveOops
    Dim correctNameOfAction As String
    Dim Extensionfile As String
    
    correctNameOfAction = ReplaceChars(NameOfAction, " ", "_")
    Extensionfile = Extension & "file"
    
    DeleteKey HKEY_CLASSES_ROOT, Extensionfile & "\shell\" & correctNameOfAction
    RemoveAction = True
    Exit Function
    
RemoveOops:
    RemoveAction = False
    Exit Function
    Resume Next
    
End Function

'Sample call:
'    EnableQuickView "txt", True

Public Function EnableQuickView(ByVal Extension As String, ByVal QuickView As Boolean) As Boolean
    On Error GoTo QuickViewOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    If QuickView = True Then
        'enable QuickView
        CreateKey HKEY_CLASSES_ROOT, Extensionfile, "QuickView"
        SaveString HKEY_CLASSES_ROOT, Extensionfile, "QuickView", "", "*"
      Else
        'disable QuickView
        DeleteKey HKEY_CLASSES_ROOT, Extensionfile & "\QuickView"
    End If
    
    EnableQuickView = True
    Exit Function
    
QuickViewOops:
    EnableQuickView = False
    Exit Function
    Resume Next
    
End Function

'Sample call:
'    AlwaysShowExt "txt", False

Public Function AlwaysShowExt(ByVal Extension As String, ByVal ShowExt As Boolean) As Boolean
    On Error GoTo ExtOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    If ShowExt = True Then
        'always show extension
        SaveString HKEY_CLASSES_ROOT, Extensionfile, "", "AlwaysShowExt", ""
      Else
        'don't show extension
        DeleteValue HKEY_CLASSES_ROOT, Extensionfile, "", "AlwaysShowExt"
    End If
    AlwaysShowExt = True
    Exit Function
    
QuickViewOops:
    AlwaysShowExt = False
    Exit Function
    Resume Next
End Function

'Sample call:
'    SetAsDefaultAction "txt", "open with Wordpad"
' The default action is used when a file is double-clicked

Public Function SetAsDefaultAction(ByVal Extension As String, ByVal NameOfAction As String) As Boolean
    On Error GoTo DefOops
    Dim Extensionfile As String, correctNameOfAction As String
    Extensionfile = Extension & "file"
    correctNameOfAction = ReplaceChars(NameOfAction, " ", "_")
    
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "Shell", "", correctNameOfAction
    
    SetAsDefaultAction = True
    Exit Function
    
DefOops:
    SetAsDefaultAction = False
    Exit Function
    Resume Next
End Function

'Sample call:
'    Existance = ExistType("txt")
' used to see if a file type exists

Public Function ExistType(ByVal Extension As String) As Boolean
    On Error GoTo OopsExist
    Dim Extensionfile As String, dotExtension As String
    
    Extensionfile = GetString(HKEY_CLASSES_ROOT, Extension & "file", "", "")
    If Extensionfile <> "" Then
        ExistType = True
      Else
        ExistType = False
    End If
    Exit Function
    
OopsExist:
    ExistType = False
    Exit Function
    Resume Next
End Function

'Sample call:
'    Replaced = ReplaceChars("Hello there. Happy New Year", " ", "_")
'   Returns "Hello_there. Happy_New_Year."

Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function

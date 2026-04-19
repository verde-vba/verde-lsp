Attribute VB_Name = "MiscStatements"

' Beep
Sub BeepDemo()
    Beep
End Sub

' AppActivate
Sub ActivateApp()
    AppActivate "Microsoft Word"
End Sub

' SendKeys
Sub SendKeysDemo()
    SendKeys "^c", True   ' Ctrl+C
End Sub

' SetAttr
Sub SetAttributes()
    SetAttr "TESTFILE", vbHidden
    SetAttr "TESTFILE", vbHidden + vbReadOnly
End Sub

' Kill, MkDir, RmDir, FileCopy, ChDir, ChDrive
Sub FileSystemOps()
    MkDir "C:\TempDir"
    ChDir "C:\TempDir"
    ChDrive "D"
    FileCopy "source.txt", "dest.txt"
    Kill "dest.txt"
    RmDir "C:\TempDir"
End Sub

' Randomize and Rnd
Sub RandomDemo()
    Dim MyValue As Long
    Randomize
    MyValue = Int((6 * Rnd) + 1)
End Sub

' Stop statement
Sub StopDemo()
    Debug.Print "Before stop"
    Stop   ' Breaks into debugger
    Debug.Print "After stop"
End Sub

' Load/Unload
Sub LoadUnloadDemo()
    Load UserForm1
    Unload UserForm1
End Sub

' SavePicture
Sub SavePictureDemo()
    SavePicture Image1.Picture, "C:\picture.bmp"
End Sub

' SaveSetting/DeleteSetting
Sub RegistryOps()
    SaveSetting "MyApp", "Settings", "Theme", "Dark"
    DeleteSetting "MyApp", "Settings", "Theme"
End Sub

' Width statement
Sub WidthDemo()
    Dim I As Long
    Open "TESTFILE" For Output As #1
    VBA.Width 1, 5
    For I = 0 To 9
       Print #1, Chr(48 + I);
    Next I
    Close #1
End Sub

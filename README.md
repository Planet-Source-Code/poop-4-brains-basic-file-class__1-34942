<div align="center">

## Basic File Class


</div>

### Description

pretty basic file class.... you can get the filename, path, drive, and extension of a file (string)... and plus you can copy files, delete files, create blanks for files (useful for prop files), and wipe files... havent tested the wipe but from what i see it should work... please vote and PLEASE COMMENT!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[poop\_4\_brains](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/poop-4-brains.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/poop-4-brains-basic-file-class__1-34942/archive/master.zip)





### Source Code

```

Function GetFilename(path As String)
GetFilename = Mid(path, InStrRev(path, "/") + 1)
End Function
Function GetFilepath(path As String)
GetFilepath = Mid(path, 1, InStrRev(path, "/"))
End Function
Function GetFileExt(path As String)
GetFileExt = Mid(path, InStrRev(path, "."))
End Function
Function GetFileDrive(path As String)
GetFileDrive = Mid(path, 1, InStr(1, path, "/"))
End Function
Function CreateBlank(path As String)
Dim Free As Integer
Free = FreeFile
Open path For Binary As Free
Close Free
End Function
Function FileExist(path As String)
On Error GoTo Oop
Dim Free As Integer
Free = FreeFile
FileExist = True
Open path For Input As Free
Close Free
Exit Function
Oop:
FileExist = False
End Function
Function CopyFile(oldpath As String, newpath As String)
FileSystem.FileCopy oldpath, newpath
End Function
Function DeleteFile(path As String)
Kill path
End Function
Private Function Rand(L, U) As Long
Dim I As Long
For I = L To U
If Int(Rnd * (U - L)) = Int(Rnd * (U - L + 1)) Then Rand = I: Exit Function
Next I
End Function
Private Function ScrambleArray(ar())
Dim I As Long
For I = LBound(ar()) To UBound(ar())
ar(I) = ar(Rand(LBound(ar()), UBound(ar())))
Next I
End Function
Function WipeFile(path As String)
Dim I As Long, B() As Byte, Free As Integer
Free = FreeFile
Open path For Binary As FreeFile
For I = 0 To 500
Get #1, , B()
ScrambleArray B()
Put #1, , B()
Next I
Close FreeFile
End Function
```


<div align="center">

## Open a  password protected MDB using DAO or Data control or OLE automation


</div>

### Description

Open a password protected MDB using DAO or Data control or OLE automation
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\~DanDa311\~](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/danda311.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/danda311-open-a-password-protected-mdb-using-dao-or-data-control-or-ole-automation__1-25115/archive/master.zip)

### API Declarations

```
'sorry the first looked bad
'I found this code somewhere and the real author was Dalin Nie he deserves all the credit. I just thought that since it was so hard for me to find it i might as well make it easy to find for the rest of society
```


### Source Code

```
'Use DAO:
Set MyDB = DBEngine.OpenDatabase(TheMDBNameWithFullPath,False,False,";Pwd=" & pwd)
Use Data Control
With Data1
.DatabaseName=App.Path & "\my.mdb"
.RecordSource="mytable"
.Connect=";Pwd=" & pwd
.Refresh
End With
Use OLE Automation
     Dim objAccess as Object
     '----------------------------------------------------------------------
     'This procedure sets a module-level variable, objAccess, to refer to
     'an instance of Microsoft Access. The code first tries to use GetObject
     'to refer to an instance that might already be open. If an instance is
     'not already open, the Shell() function opens a new instance and
     'specifies the user and password, based on the arguments passed to the
     'procedure.
     '
     'Calling example: OpenSecured varUser:="Admin", varPw:=""
     '----------------------------------------------------------------------
     Sub OpenSecured(Optional varUser As Variant, Optional varPw As Variant)
       Dim cmd As String
       On Error Resume Next
       Set objAccess = GetObject(, "Access.Application")
       If Err <> 0 Then 'no instance of Access is open
        If IsMissing(varUser) Then varUser = "Admin"
        cmd = "C:\Program Files\Microsoft Office\Office\MSAccess.exe"
        cmd = cmd & " /nostartup /user " & varUser
        If Not IsMissing(varPw) Then cmd = cmd & " /pwd " & varPw
        Shell pathname:=cmd, windowstyle:=6
        Do 'Wait for shelled process to finish.
         Err = 0
         Set objAccess = GetObject(, "Access.Application")
        Loop While Err <> 0
       End If
```


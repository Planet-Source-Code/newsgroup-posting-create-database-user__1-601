<div align="center">

## Create database user


</div>

### Description



The following function creates a user. You can execute it under any user you like.

dror-a@euronet.co.il (Dror Dotan A')
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |4.2 (164 globes from 39 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-create-database-user__1-601/archive/master.zip)

### API Declarations

```
const ADMIN_USERNAME = "Admin"
const ADMIN_PASSWORD = "adminpass (or whatever)"
const SHOWICON_STOP = 16
```


### Source Code

```
Function CreateNewUser% (ByVal username$, ByVal password$, ByVal PID$)
  '- create a new user.
  '- username$ - name
  '- password$ - user password
  '- PID$ - PID of user
  '-----------------------------------
  Dim NewUser As User
  Dim admin_ws As WorkSpace
  '=====================================
  '- check PID
  If (Len(PID$) < 4 Or Len(PID$) > 20) Then
    MsgBox "Invalid PID", SHOWICON_STOP
    CreateNewUser% = True
    Exit Function
  End If
  '- verify that user does not yet exist
  If (UserExist%(username$)) Then
    CreateNewUser% = True
    Exit Function
  End If
  '- open new workspace and database as admin
  dbEngine.Workspaces.Refresh
  Set admin_ws = dbEngine.CreateWorkspace("TempWorkSpace",
                     ADMIN_USER, ADMIN_PASSWORD)
  If (Err) Then
    '- failed opening workspace
    MsgBox "invalid administrator password", SHOWICON_STOP
    MsgBox "Error: " & Error$, SHOWICON_STOP, SystemName
    CreateNewUser% = True
    Exit Function
  End If
  On Error Resume Next
  '- create the new user
  Set NewUser = admin_ws.CreateUser(username$, PID$, password$)
  If (Err) Then
    MsgBox "Can't create new user.", SHOWICON_STOP
    MsgBox Error$, SHOWICON_STOP
    GoTo CreateNewUser_end
  End If
  '- add user to user list
  admin_ws.Users.Append NewUser
  '- add user to "Users" group
  Set NewUser = admin_ws.CreateUser(username$)
  admin_ws.Groups("Users").Users.Append NewUser
  admin_ws.Users(username$).Groups.Refresh
  admin_ws.Close
  CreateNewUser% = False
CreateNewUser_end:
  On Error GoTo 0
End Function
```


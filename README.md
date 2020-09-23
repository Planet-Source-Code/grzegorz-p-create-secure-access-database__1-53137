<div align="center">

## Create secure Access database


</div>

### Description

This sample shows how to secure Access database using ADOX and SIDs. In apps where database is created on-the-fly from code its only way to prevent your database from being modified.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Grzegorz P\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/grzegorz-p.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/grzegorz-p-create-secure-access-database__1-53137/archive/master.zip)





### Source Code

```
Public Function SecureAccessDB( _
                ByVal DatabaseFileName As String, _
                ByVal SystemMDWFileName As String, _
                ByVal UID As String, _
                ByVal PWD As String, _
                ByVal GroupName As String, _
                ByVal GroupPID As String, _
                ByVal UserID As String, _
                ByVal UserPWD As String, _
                ByVal UserPID As String) _
                As Boolean
 '********************************************************************
 'Inputs:
 ' DatabaseFileName - DB to secure file name
 ' SystemMDWFileName - System.mdw file name
 ' UID - current owner name (usually 'admin')
 ' PWD - current owner password (usually empty string)
 ' GroupName - new group to create
 ' GroupPID - new group PID
 ' UserID - new user account to create
 ' UserPWD - new user password
 ' UserPID - new user PID
 '
 'Returns:
 'True - if no errors occurs
 'False - otherwise
 '
 'References:
 'ADO and ADOX libraries ver. 2.8
 '*********************************************************************
 Dim catDatabase As ADOX.Catalog
 Dim tblTemp As ADOX.Table
 Dim cmdCreator As ADODB.Command
 Dim strTableName As String
 Dim strConnString As String
 'if DB was formerly secured by us, then we can login to DB using our user account and is nothing to do
  strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & DatabaseFileName & ";" & _
          "User ID=" & UserID & ";" & _
          "Password=" & UserPWD & ";" & _
          "jet oledb:system database=" & SystemMDWFileName
  Set catDatabase = New ADOX.Catalog
  Err.Clear
  On Error Resume Next
    catDatabase.ActiveConnection = strConnString
    If Err.Number = 0 Then 'can login
      Set catDatabase = Nothing
      SecureAccessDB = True
      Exit Function
     Else 'can't login
      Set catDatabase = Nothing
      SecureAccessDB = False
    End If
    'secure DB
  On Error GoTo EH_SecureAccessDB
  'login to DB using current owner account
  strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & DatabaseFileName & ";" & _
          "User ID=" & UID & ";" & _
          "Password=" & PWD & ";" & _
          "jet oledb:system database=" & SystemMDWFileName
  'create new group and user account
  Set cmdCreator = New ADODB.Command
  cmdCreator.ActiveConnection = strConnString
  cmdCreator.CommandText = "CREATE GROUP " & GroupName & " " & GroupPID & ";"
  cmdCreator.Execute
  cmdCreator.CommandText = "CREATE USER " & UserID & " " & UserPWD & " " & UserPID & ";"
  cmdCreator.Execute
  Set cmdCreator = Nothing
  Set catDatabase = New ADOX.Catalog
  catDatabase.ActiveConnection = strConnString
  With catDatabase
    .Users(UserID).Groups.Append "Admins"
    .Users(UserID).Groups.Append GroupName
    .Users(UserID).SetPermissions "", adPermObjDatabase, adAccessGrant, adRightMaximumAllowed
    .Users(UserID).SetPermissions Null, adPermObjTable, adAccessGrant, adRightMaximumAllowed
    .Groups(GroupName).SetPermissions "", adPermObjDatabase, adAccessGrant, adRightMaximumAllowed
    .Groups(GroupName).SetPermissions Null, adPermObjTable, adAccessGrant, adRightMaximumAllowed
    'Grant/Revoke rights to user tables
    For Each tblTemp In .Tables
      If tblTemp.Type = "TABLE" Then 'keep Access and System tables unchanged
        strTableName = tblTemp.Name
        'Change owner of each table
        .SetObjectOwner strTableName, adPermObjTable, UserID
        'Grant all rights to MyUser and MyGroup
        .Users(UserID).SetPermissions strTableName, adPermObjTable, adAccessGrant, adRightMaximumAllowed
        .Groups(GroupName).SetPermissions strTableName, adPermObjTable, adAccessGrant, adRightMaximumAllowed
        'Revoke all rights to admin account and Admins/Users groups
        .Users("admin").SetPermissions strTableName, adPermObjTable, adAccessRevoke, adRightMaximumAllowed
        .Groups("Admins").SetPermissions strTableName, adPermObjTable, adAccessRevoke, adRightMaximumAllowed
        .Groups("Users").SetPermissions strTableName, adPermObjTable, adAccessRevoke, adRightMaximumAllowed
      End If
    Next 'tblTemp
    'Revoke all database rights to admin account and Admins/Users groups
    .Groups("Users").SetPermissions Null, adPermObjTable, adAccessRevoke, adRightMaximumAllowed
    .Groups("Users").SetPermissions "", adPermObjDatabase, adAccessRevoke, adRightMaximumAllowed
    .Users("admin").SetPermissions Null, adPermObjTable, adAccessRevoke, adRightMaximumAllowed
    .Users("admin").SetPermissions "", adPermObjDatabase, adAccessRevoke, adRightMaximumAllowed
    .Groups("Admins").SetPermissions Null, adPermObjTable, adAccessRevoke, adRightMaximumAllowed
    .Groups("Admins").SetPermissions "", adPermObjDatabase, adAccessRevoke, adRightMaximumAllowed
  End With 'catDatabase
  Set catDatabase = Nothing
  Set tblTemp = Nothing
  SecureAccessDB = True
Exit Function
EH_SecureAccessDB:
  Set cmdCreator = Nothing
  Set catDatabase = Nothing
  Set tblTemp = Nothing
  SecureAccessDB = False
End Function
```


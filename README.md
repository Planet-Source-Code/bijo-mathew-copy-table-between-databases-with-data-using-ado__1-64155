<div align="center">

## Copy Table Between Databases With Data using ADO


</div>

### Description

This code copies a table from 1 ms access database to another using a Select Query....Easy to use and aptly commented...Check it out!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bijo Mathew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bijo-mathew.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bijo-mathew-copy-table-between-databases-with-data-using-ado__1-64155/archive/master.zip)

### API Declarations

You need a refrence to ADO 2.4/5/6/7/8 Library


### Source Code

```
Option Explicit
Dim cnOld As New ADODB.Connection
Dim cnNew As New ADODB.Connection
Private Sub Command1_Click()
'set your select statement here
Dim rsOld As New ADODB.Recordset
Set rsOld = Nothing
rsOld.Open "select * from 1T", cnOld
Call createTable(rsOld, cnNew)
End Sub
Private Sub Form_Load()
'set 2 databases here
cnOld.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & App.Path & "\1.mdb" & "' ;Jet OLEDB:Database Password=''")
cnNew.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & App.Path & "\2.mdb" & "' ;Jet OLEDB:Database Password=''")
End Sub
Function createTable(rsOld As ADODB.Recordset, cnNew As ADODB.Connection)
On Error GoTo Err
Dim intX As Integer
Dim strTable As String
Dim rsNew As New ADODB.Recordset
'set table name
strTable = rsOld.Fields.Item(0).Properties.Item("BASETABLENAME").Value
intX = 0
'deletes if table exists...comment this line if you -
'dont want the existing table to be deleted
On Error GoTo err1
cnNew.Execute "Drop table [" & strTable & "]"
'create table
cnNew.Execute "Create table [" & strTable & "]"
While intX < rsOld.Fields.Count
  With rsOld.Fields.Item(intX)
    cnNew.Execute "Alter table " & strTable & " Add Column [" & .Name & "] " & dataType(.Type)
    intX = intX + 1
  End With
Wend
'transfer data
rsNew.Open "Select * from " & strTable, cnNew, adOpenDynamic, adLockOptimistic
If rsOld.EOF = False Then
  rsOld.MoveFirst
  While rsOld.EOF = False
    intX = 0
    rsNew.AddNew
    While intX < rsOld.Fields.Count
      rsNew(intX) = rsOld(intX)
      intX = intX + 1
    Wend
    rsNew.Update
    rsOld.MoveNext
  Wend
End If
MsgBox "Table and data transferred", vbInformation
Exit Function
Err:
MsgBox Err.Description, vbExclamation
Exit Function
err1:
Resume Next
End Function
Function dataType(intType As Long) As String
If CInt(intType) = 3 Then
  dataType = "Long"
ElseIf CInt(intType) = 6 Then
  dataType = "Currency"
ElseIf CInt(intType) = 7 Then
  dataType = "Date"
ElseIf CInt(intType) = 11 Then
  dataType = "YesNo"
ElseIf CInt(intType) = 203 Then
  dataType = "Memo"
Else
  dataType = "VarChar"
End If
End Function
```


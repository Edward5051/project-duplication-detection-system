VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Sub DisplayUsers(lstUsers As ListView)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
    sql = "Select * From Accounts Where Hidden <> True Order By Username"
   rs.Open sql, conn
   lstUsers.ListItems.Clear
   Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstUsers.ListItems.Add(, , a)
            lstItem.SubItems(1) = rs(0).Value
           ' lstItem.SubItems(2) = rs(2).Value
           ' lstItem.SubItems(5) = rs(1).Value
            rs.MoveNext
            Loop
End Sub


Sub AddUser(Username As String)

If rs.State = adStateOpen Then rs.Close
    sql = "Select * From Accounts Where username='" & Username & "'"
    rs.Open sql, conn
        If rs.RecordCount >= 1 Then
            MsgBox "Duplicate project found.Please enter another project.", vbCritical, ""
            Exit Sub
        End If
    With rs
        .AddNew
        !Username = Username
        
        .Update
        End With
        MsgBox "Record(s) Updated", vbInformation, ""
End Sub

Sub Initialize()
With mrs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
End With
End Sub



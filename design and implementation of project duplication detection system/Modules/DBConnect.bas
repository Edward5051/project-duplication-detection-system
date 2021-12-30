Attribute VB_Name = "DBConnect"
Public rs As New ADODB.Recordset
Public conn As New ADODB.Connection
Public sql As String
Public ConString As String
Public CurrentUser As String
Public LoginSuccess As Boolean
Public UserTitle As String
Public TempBorrowerID As String
Public TempBorrowerName As String
Public TempBookID As String
Public TempBookName As String
Public TempBookCopy As String
Public TempCopyBorrow As String
Public BorrowDate As String
Public DueDate As String
Public mTransID As String
Public userlog As Integer
Public TempContact As String
'Public rAdd As Boolean, rDelete As Boolean, rUpdate As Boolean, rPrint As Boolean

 Sub Main()
'ConString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=" & App.Path & "\Data.mdb;Jet OLEDB:Database Password="
ConString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=" & App.Path & "\Data.mdb;Jet OLEDB:Database Password="
conn.Open ConString
With rs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
End With
AccountsFrm.Show
'MainFrm.Show
End Sub


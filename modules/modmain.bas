Attribute VB_Name = "modmain"
Global conn As ADODB.Connection
Global rst As ADODB.Recordset
Global dbstring
Global useraccess, currentuser

Sub main()
dbstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\libsystem.mdb" & ";Persist Security Info=False"
'frmlogin.Show
mainform.Show
End Sub

Sub backupfile(ByVal f1 As String, ByVal f2 As String)
On Error Resume Next
Dim b() As Byte
  Open f1 For Binary Access Read As #1
  ReDim b(1 To LOF(1))
  Get #1, 1, b
  Close #1
  Open f2 For Binary Access Write As #2
  Put #2, 1, b
  Close #2
  Erase b
End Sub

Sub closedb()
On Error GoTo handler
rst.Close
conn.Close
handler:
If Err.Number <> 0 Then MsgBox Err.Description, vbInformation
End Sub

Sub connectdb(sqlcode)
On Error GoTo handler
Set conn = New ADODB.Connection
Set rst = New ADODB.Recordset
conn.Open dbstring
rst.Open sqlcode, conn, 1, 2
handler:
If Err.Number <> 0 Then MsgBox Err.Description, vbInformation
End Sub

Function toproper(st)
On Error GoTo errhand
If st <> "" Then
tmptxt = ""
s = ""
For no = 1 To Len(st)
s = LCase(Mid(st, no, 1))
If no > 1 Then
If Asc(Mid(st, no - 1, 1)) = 32 Then s = UCase(Mid(st, no, 1))
Else
s = UCase(Mid(st, no, 1))
End If
tmptxt = tmptxt & s
Next
toproper = tmptxt
Else
toproper = ""
End If
errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbInformation
End Function



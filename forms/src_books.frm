VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8AD9B863-ED9F-467A-960D-1CDA7D4C6F75}#2.0#0"; "madgear.ocx"
Object = "{F4BC31DC-6CF1-4FF5-A49B-E65FD905E80D}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form src_books 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6900
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10665
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4530
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "src_books.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      TabIndex        =   5
      Top             =   570
      Width           =   8475
   End
   Begin MSComctlLib.ListView list1 
      Height          =   5355
      Left            =   810
      TabIndex        =   3
      Top             =   960
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   9446
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Book Number"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Year Publish"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Category"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Location"
         Object.Width           =   1764
      EndProperty
   End
   Begin madgearXControls.mButton cmdnew 
      Height          =   615
      Left            =   30
      TabIndex        =   2
      Top             =   990
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      BTYPE           =   1
      TX              =   "&New"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "src_books.frx":039A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.gradient gradient1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   820
      GradColor1      =   4194304
      GradColor2      =   8388608
      Begin LaVolpeAlphaImg.AlphaImgCtl hicon 
         Height          =   345
         Left            =   60
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         Image           =   "src_books.frx":03B6
         Frame           =   4100
         Attr            =   513
         Effects         =   "src_books.frx":3793
         BkgImage        =   "src_books.frx":37AB
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Details"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   540
         TabIndex        =   1
         Top             =   90
         Width           =   1155
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl controlbtn 
         Height          =   330
         Index           =   0
         Left            =   10260
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         Image           =   "src_books.frx":55E7
         Settings        =   16777216
         Attr            =   513
         Object.Index           =   65552
         Effects         =   "src_books.frx":3C84C
      End
   End
   Begin madgearXControls.mButton cmdedit 
      Height          =   615
      Left            =   30
      TabIndex        =   6
      Top             =   1620
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      BTYPE           =   1
      TX              =   "&Edit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "src_books.frx":3C864
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.mButton cmddelete 
      Height          =   615
      Left            =   30
      TabIndex        =   7
      Top             =   2250
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      BTYPE           =   1
      TX              =   "&Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "src_books.frx":3C880
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.mButton cmdprint 
      Height          =   615
      Left            =   30
      TabIndex        =   8
      Top             =   2880
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      BTYPE           =   1
      TX              =   "&Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "src_books.frx":3C89C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.gradient gradient2 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   9
      Top             =   6435
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   820
      GradColor1      =   8388608
      GradColor2      =   0
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Count : "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7860
         TabIndex        =   13
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label reccount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9090
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
   End
   Begin madgearXControls.mButton cmdClose 
      Height          =   615
      Left            =   30
      TabIndex        =   10
      Top             =   3510
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
      BTYPE           =   1
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "src_books.frx":3C8B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.mButton cmdsearch 
      Height          =   405
      Left            =   9300
      TabIndex        =   11
      Top             =   540
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      BTYPE           =   1
      TX              =   "&Search"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "src_books.frx":3C8D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   645
   End
End
Attribute VB_Name = "src_books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function genbooknum()
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
strsql = "SELECT * From tbl_books"
db.Open dbstring
rs.Open strsql, db, 1, 1
xx = rs.RecordCount + 1
Select Case xx
Case 0 To 99999
tmpbknum = "BA" & Format(xx, "00000")
Case 100000 To 19999
xx = xx - 99999
tmpbknum = "BB" & Format(xx, "00000")
End Select
genbooknum = tmpbknum
rs.Close
db.Close
End Function

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
msg = MsgBox("Are you sure you want to delete this book?", vbQuestion + vbYesNo)
If msg <> vbYes Then Exit Sub



End Sub

Private Sub cmdedit_Click()
On Error GoTo errhand
If list1.ListItems.Count = 0 Then Exit Sub
frmbooks.fld(0).Text = list1.SelectedItem
frmbooks.loadbooks
frmbooks.cmdSave.Caption = "&Update"
frmbooks.Show 1, Me
errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdnew_Click()
frmbooks.fld(0) = genbooknum
frmbooks.Show 1, Me
End Sub

Private Sub cmdsearch_Click()
searchlst Text1
End Sub

Private Sub controlbtn_Click(Index As Integer)
Unload Me
End Sub
Private Sub controlbtn_MouseEnter(Index As Integer)
controlbtn(Index).GrayScale = lvicNoGrayScale
End Sub

Private Sub controlbtn_MouseExit(Index As Integer)
controlbtn(Index).GrayScale = lvicNTSCPAL
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub


Sub searchlist(st)
On Error GoTo errhand
strsql = "SELECT * from tbl_books"
psearch = st
If psearch <> "" Then
    dbwhere = ""
    psearch = Replace(psearch, "'", "''")
    psearch = Replace(psearch, "[", "[[]")
    dbwhere = dbwhere & "[booknum] like '%" & psearch & "%' or "
    dbwhere = dbwhere & "[booktitle] like '%" & psearch & "%' or "
    dbwhere = dbwhere & "[bookdesc] like '%" & psearch & "%' or "
    dbwhere = dbwhere & "[bookauthor] like '%" & psearch & "%' or "
    dbwhere = dbwhere & "[yearpublish] like '%" & psearch & "%' or "
    dbwhere = dbwhere & "[category] like '%" & psearch & "%' or "
    dbwhere = dbwhere & "[location] like '%" & psearch & "%' or "
    dbwhere = Mid(dbwhere, 1, Len(dbwhere) - 3)
strsql = strsql & " WHERE (" & dbwhere & " )"
End If

Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
db.Open dbstring
rs.Open strsql, db, adOpenKeyset, adLockReadOnly, _
      adCmdTableDirect

If rs.RecordCount = 0 Then
list1.ListItems.Clear
Exit Sub
End If

With rs
list1.ListItems.Clear
Do While Not .EOF
Set lvx = list1.ListItems.Add(, , .Fields(0))
lvx.SmallIcon = 1
lvx.SubItems(1) = .Fields(1)
lvx.SubItems(2) = .Fields(2)
lvx.SubItems(3) = .Fields(7)
.MoveNext
Loop
rs.Close
db.Close
End With
reccount = list1.ListItems.Count
errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
searchlist ""
End Sub



Private Sub list1_DblClick()
Call cmdedit_Click
End Sub


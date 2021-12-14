VERSION 5.00
Object = "{8AD9B863-ED9F-467A-960D-1CDA7D4C6F75}#2.0#0"; "madgear.ocx"
Object = "{F4BC31DC-6CF1-4FF5-A49B-E65FD905E80D}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmbooks 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox fld 
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
      Index           =   5
      Left            =   1500
      TabIndex        =   6
      Top             =   4140
      Width           =   1245
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1500
      TabIndex        =   5
      Top             =   3720
      Width           =   2715
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1500
      TabIndex        =   4
      Top             =   3300
      Width           =   2055
   End
   Begin VB.TextBox fld 
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
      Index           =   4
      Left            =   1500
      TabIndex        =   3
      Top             =   2850
      Width           =   1905
   End
   Begin VB.TextBox fld 
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
      Index           =   3
      Left            =   1500
      TabIndex        =   2
      Top             =   2400
      Width           =   3585
   End
   Begin VB.TextBox fld 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   2
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1620
      Width           =   3255
   End
   Begin VB.TextBox fld 
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
      Index           =   1
      Left            =   1500
      TabIndex        =   0
      Top             =   1170
      Width           =   4515
   End
   Begin VB.TextBox fld 
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
      Index           =   0
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   660
      Width           =   1875
   End
   Begin madgearXControls.mButton cmdSave 
      Height          =   495
      Left            =   3210
      TabIndex        =   7
      Top             =   4710
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "&Save"
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
      FOCUSR          =   0   'False
      BCOL            =   8421504
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmbooks.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.mButton cmdClose 
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   4710
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
      BTYPE           =   9
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
      FOCUSR          =   0   'False
      BCOL            =   8421504
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmbooks.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
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
      TabIndex        =   10
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   820
      GradColor1      =   0
      GradColor2      =   192
      Begin LaVolpeAlphaImg.AlphaImgCtl controlbtn 
         Height          =   330
         Index           =   0
         Left            =   5850
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         Image           =   "frmbooks.frx":0038
         Settings        =   16777216
         Attr            =   513
         Object.Index           =   65552
         Effects         =   "frmbooks.frx":3729D
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Entry"
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
         TabIndex        =   17
         Top             =   90
         Width           =   990
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
         Height          =   345
         Left            =   60
         Top             =   60
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         Image           =   "frmbooks.frx":372B5
         Frame           =   4100
         Attr            =   513
         Effects         =   "frmbooks.frx":390F1
         BkgImage        =   "frmbooks.frx":39109
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Initial QTY : "
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
      Index           =   7
      Left            =   180
      TabIndex        =   19
      Top             =   4170
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Location :"
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
      Index           =   6
      Left            =   150
      TabIndex        =   18
      Top             =   3750
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category :"
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
      Index           =   5
      Left            =   180
      TabIndex        =   16
      Top             =   3330
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Year Publish :"
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
      Index           =   4
      Left            =   180
      TabIndex        =   15
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author :"
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
      Index           =   3
      Left            =   180
      TabIndex        =   14
      Top             =   2430
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
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
      Index           =   2
      Left            =   180
      TabIndex        =   13
      Top             =   1590
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title :"
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
      Index           =   1
      Left            =   180
      TabIndex        =   12
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Book Number : "
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   690
      Width           =   1320
   End
End
Attribute VB_Name = "frmbooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
For no = 0 To 5
If fld(no) = "" Then
MsgBox "This is a required field!", vbInformation
fld(0).SetFocus
Exit Sub
End If
Next

If Combo1.Text = "" Then
MsgBox "This is a required field!", vbInformation
Combo1.SetFocus
Exit Sub
End If

If Combo2.Text = "" Then
MsgBox "This is a required field!", vbInformation
Combo2.SetFocus
Exit Sub
End If


If cmdSave.Caption = "&Save" Then
connectdb "select * from tbl_books"
With rst
.AddNew
.Fields("booknum") = fld(0).Text
.Fields("booktitle") = fld(1).Text
.Fields("bookdesc") = fld(2).Text
.Fields("bookauthor") = fld(3).Text
.Fields("yearpublish") = fld(4).Text
.Fields("category") = Combo1.Text
.Fields("location") = Combo2.Text
.Fields("initialqty") = fld(5).Text
.Fields("dateadd") = Now
.Fields("addby") = currentuser
.Update
End With
closedb
cmdSave.Caption = "&Update"
Else
connectdb "select * from tbl_books where booknum = '" & fld(0).Text & "'"
If Not rst.EOF Then
With rst

.Fields("booktitle") = fld(1).Text
.Fields("bookdesc") = fld(2).Text
.Fields("bookauthor") = fld(3).Text
.Fields("yearpublish") = fld(4).Text
.Fields("category") = Combo1.Text
.Fields("location") = Combo2.Text
.Fields("datemod") = Now
.Fields("modby") = currentuser
.Update
MsgBox "Record Updated!"
End With
End If
closedb
End If
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

Private Sub Form_Load()
connectdb "select * from tbl_loc"
Do While Not rst.EOF
Combo2.AddItem rst.Fields(0)
rst.MoveNext
Loop
End Sub

Sub loadbooks()
'On Error GoTo errhand
connectdb "select * from tbl_books where booknum = '" & fld(0).Text & "'"
If Not rst.EOF Then
With rst
fld(1).Text = .Fields("booktitle")
fld(2).Text = .Fields("bookdesc")
fld(3).Text = .Fields("bookauthor")
fld(4).Text = .Fields("yearpublish")
fld(5).Text = .Fields("initialqty")
Combo1.Text = .Fields("category")
Combo2.Text = .Fields("location")
fld(5).Enabled = False
End With
End If
closedb
errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbInformation
End Sub

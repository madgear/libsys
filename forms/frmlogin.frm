VERSION 5.00
Object = "{8AD9B863-ED9F-467A-960D-1CDA7D4C6F75}#2.0#0"; "madgear.ocx"
Object = "{F4BC31DC-6CF1-4FF5-A49B-E65FD905E80D}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmlogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5100
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox fld 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1410
      MaxLength       =   20
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   3510
      Width           =   2625
   End
   Begin VB.TextBox fld 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2820
      Width           =   2625
   End
   Begin madgearXControls.gradient gradient1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   820
      GradColor1      =   16384
      GradColor2      =   32768
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
         Height          =   375
         Left            =   60
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Image           =   "frmlogin.frx":0000
         Frame           =   4100
         Attr            =   513
         Effects         =   "frmlogin.frx":3C62
         BkgImage        =   "frmlogin.frx":3C7A
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log-In"
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
         Index           =   2
         Left            =   540
         TabIndex        =   6
         Top             =   90
         Width           =   570
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl controlbtn 
         Height          =   330
         Index           =   0
         Left            =   3870
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         Image           =   "frmlogin.frx":5AB6
         Settings        =   16777216
         Attr            =   513
         Object.Index           =   65552
         Effects         =   "frmlogin.frx":3CD1B
      End
   End
   Begin madgearXControls.mButton cmdLogin 
      Height          =   495
      Left            =   900
      TabIndex        =   2
      Top             =   4260
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "&Log-In"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   8421504
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmlogin.frx":3CD33
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
      Left            =   2310
      TabIndex        =   3
      Top             =   4260
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   8421504
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmlogin.frx":3CD4F
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   555
      Left            =   1290
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   990
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl2 
      Height          =   1770
      Left            =   1260
      Top             =   600
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3122
      Image           =   "frmlogin.frx":3CD6B
      Attr            =   513
      Effects         =   "frmlogin.frx":3DC9B
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   555
      Left            =   1290
      Shape           =   4  'Rounded Rectangle
      Top             =   2700
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      Width           =   1065
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
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
If KeyCode = 27 Then End
If KeyCode = 13 Then checkuser
End Sub

Sub checkuser()
On Error GoTo handler
sqlcode = "SELECT * From tbl_profile WHERE (((fld_username)='" & fld(0) & "') AND ((fld_password)='" & fld(1) & "'));"
connectdb sqlcode
If Not rst.EOF Then
useraccess = rst.Fields(7)
currentuser = rst.Fields(0)
Unload Me
mainform.Show
Else
MsgBox "Invalid username or Password!", vbInformation
fld(1) = ""
fld(0).SelStart = 0
fld(0).SelLength = Len(fld(0))
fld(0).SetFocus
End If
closedb
handler:
If Err.Number <> 0 Then MsgBox Err.Description, vbInformation
End Sub

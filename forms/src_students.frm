VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8AD9B863-ED9F-467A-960D-1CDA7D4C6F75}#2.0#0"; "madgear.ocx"
Object = "{F4BC31DC-6CF1-4FF5-A49B-E65FD905E80D}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form src_students 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6900
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10680
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
      TabIndex        =   0
      Top             =   570
      Width           =   8475
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5355
      Left            =   810
      TabIndex        =   1
      Top             =   960
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   9446
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
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
      NumItems        =   0
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
      MICON           =   "src_students.frx":0000
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
      TabIndex        =   3
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   820
      GradColor1      =   4210688
      GradColor2      =   8421376
      Begin LaVolpeAlphaImg.AlphaImgCtl controlbtn 
         Height          =   330
         Index           =   0
         Left            =   10260
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         Image           =   "src_students.frx":001C
         Settings        =   16777216
         Attr            =   513
         Object.Index           =   65552
         Effects         =   "src_students.frx":37281
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Details"
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
         TabIndex        =   4
         Top             =   90
         Width           =   1425
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
         Height          =   345
         Left            =   120
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Image           =   "src_students.frx":37299
         Frame           =   4100
         Attr            =   513
         Effects         =   "src_students.frx":39B6E
      End
   End
   Begin madgearXControls.mButton mButton1 
      Height          =   615
      Left            =   30
      TabIndex        =   5
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
      MICON           =   "src_students.frx":39B86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.mButton mButton2 
      Height          =   615
      Left            =   30
      TabIndex        =   6
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
      MICON           =   "src_students.frx":39BA2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.mButton mButton3 
      Height          =   615
      Left            =   30
      TabIndex        =   7
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
      MICON           =   "src_students.frx":39BBE
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
      TabIndex        =   8
      Top             =   6435
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   820
      GradColor1      =   4210688
      GradColor2      =   0
   End
   Begin madgearXControls.mButton cmdClose 
      Height          =   615
      Left            =   30
      TabIndex        =   9
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
      MICON           =   "src_students.frx":39BDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin madgearXControls.mButton mButton5 
      Height          =   405
      Left            =   9300
      TabIndex        =   10
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
      MICON           =   "src_students.frx":39BF6
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
      TabIndex        =   11
      Top             =   600
      Width           =   645
   End
End
Attribute VB_Name = "src_students"
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
If KeyCode = 27 Then Unload Me
End Sub


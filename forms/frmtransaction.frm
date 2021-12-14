VERSION 5.00
Object = "{8AD9B863-ED9F-467A-960D-1CDA7D4C6F75}#2.0#0"; "madgear.ocx"
Object = "{F4BC31DC-6CF1-4FF5-A49B-E65FD905E80D}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmtransaction 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8580
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8805
   Begin madgearXControls.gradient gradient1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   820
      GradColor1      =   4210816
      GradColor2      =   16512
      Begin LaVolpeAlphaImg.AlphaImgCtl controlbtn 
         Height          =   330
         Index           =   0
         Left            =   8370
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         Image           =   "frmtransaction.frx":0000
         Settings        =   16777216
         Attr            =   513
         Object.Index           =   65552
         Effects         =   "frmtransaction.frx":37265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Transaction"
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
         Left            =   510
         TabIndex        =   1
         Top             =   90
         Width           =   1575
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
         Height          =   345
         Left            =   90
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Image           =   "frmtransaction.frx":3727D
         Attr            =   513
         Effects         =   "frmtransaction.frx":39DC2
      End
   End
End
Attribute VB_Name = "frmtransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

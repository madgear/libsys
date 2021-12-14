VERSION 5.00
Object = "{8AD9B863-ED9F-467A-960D-1CDA7D4C6F75}#2.0#0"; "madgear.ocx"
Object = "{F4BC31DC-6CF1-4FF5-A49B-E65FD905E80D}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form mainform 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7245
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin madgearXControls.gradient mLeft 
      Align           =   3  'Align Left
      Height          =   5625
      Left            =   0
      TabIndex        =   2
      Top             =   825
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   9922
      GradColor1      =   4210688
      GradColor2      =   8421376
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   0
         Left            =   90
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   330
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "&HOME"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   1
         Left            =   90
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "&BOOKS"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   2
         Left            =   90
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1350
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "&STUDENTS"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   3
         Left            =   90
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1860
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "&TRANSACTION"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   4
         Left            =   90
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2370
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "BOOK &MONITORING"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   5
         Left            =   90
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "&REPORT"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   6
         Left            =   90
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3390
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "T&OOLS/SETTINGS"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   7
         Left            =   90
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3900
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "&LOG OUT"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":00C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin madgearXControls.mButton menubutton 
         Height          =   465
         Index           =   8
         Left            =   90
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4410
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   820
         BTYPE           =   9
         TX              =   "E&XIT"
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
         BCOL            =   64
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mainform.frx":00E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin madgearXControls.gradient upper 
      Align           =   1  'Align Top
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1455
      GradColor1      =   4210752
      GradColor2      =   0
      Begin LaVolpeAlphaImg.AlphaImgCtl controlbtn 
         Height          =   480
         Index           =   1
         Left            =   10530
         Top             =   150
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Image           =   "mainform.frx":00FC
         Settings        =   16777216
         Object.Index           =   65554
         Effects         =   "mainform.frx":40A3E
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl controlbtn 
         Height          =   480
         Index           =   0
         Left            =   11040
         Top             =   150
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Image           =   "mainform.frx":40A56
         Settings        =   16777216
         Object.Index           =   65552
         Effects         =   "mainform.frx":77CBB
      End
      Begin VB.Label headertitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Automated Library System "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   780
         TabIndex        =   3
         Top             =   90
         Width           =   3105
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
         Height          =   660
         Left            =   60
         Top             =   90
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1164
         Image           =   "mainform.frx":77CD3
         Frame           =   4101
         Attr            =   513
         Effects         =   "mainform.frx":7A104
      End
      Begin VB.Label headertitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teodoro M. Luansing College of Rosario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   540
         Index           =   1
         Left            =   780
         TabIndex        =   8
         Top             =   270
         Width           =   7065
      End
   End
   Begin madgearXControls.gradient footer 
      Align           =   2  'Align Bottom
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   6450
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1402
      GradColor1      =   4210752
      GradColor2      =   0
      Begin madgearXControls.gradient keyfrm 
         Height          =   795
         Left            =   3750
         TabIndex        =   11
         Top             =   0
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1402
         GradColor1      =   4210752
         GradColor2      =   0
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NUM :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   19
            Top             =   390
            Width           =   525
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CAPS :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   18
            Top             =   90
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "SCROLL : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   5
            Left            =   1485
            TabIndex        =   17
            Top             =   90
            Width           =   900
         End
         Begin VB.Label keys 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   0
            Left            =   690
            TabIndex        =   16
            Top             =   90
            Width           =   315
         End
         Begin VB.Label keys 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   690
            TabIndex        =   15
            Top             =   390
            Width           =   315
         End
         Begin VB.Label keys 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   2
            Left            =   2430
            TabIndex        =   14
            Top             =   90
            Width           =   315
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "INSERT : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   4
            Left            =   1485
            TabIndex        =   13
            Top             =   360
            Width           =   900
         End
         Begin VB.Label keys 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   3
            Left            =   2430
            TabIndex        =   12
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Timer keytimer 
         Interval        =   1
         Left            =   8010
         Top             =   180
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "12/12/2013"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   9810
         TabIndex        =   10
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lbltime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00 AM"
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
         Left            =   10755
         TabIndex        =   9
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2190
         TabIndex        =   7
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label cuser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "madgear"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2190
         TabIndex        =   6
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "User Access :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Current User :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   60
         Width           =   1245
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl2 
         Height          =   540
         Left            =   210
         Top             =   120
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Frame           =   4101
         Attr            =   513
         Effects         =   "mainform.frx":7A11C
         BkgImage        =   "mainform.frx":7A134
      End
   End
   Begin madgearXControls.gradient tb 
      Height          =   705
      Left            =   3420
      TabIndex        =   28
      Top             =   810
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   1244
      GradColor1      =   15779735
      GradColor2      =   12017457
      GradientStyle   =   0
      Begin LaVolpeAlphaImg.AlphaImgCtl tbicon 
         Height          =   570
         Left            =   150
         Top             =   60
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1005
         Image           =   "mainform.frx":7E200
         Attr            =   513
         Effects         =   "mainform.frx":808A8
      End
      Begin VB.Label tbtitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   900
         TabIndex        =   29
         Top             =   120
         Width           =   960
      End
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Sub showform(frm As Form)
On Error Resume Next

frm.Top = (tb.Top + tb.Height) + 200
frm.Left = ((Me.Width \ 2) + (mLeft.Width \ 2)) - (frm.Width \ 2)
frm.Show ' 1, Me
End Sub

Private Sub controlbtn_Click(Index As Integer)
Select Case Index
Case 0
msg = MsgBox("are you sure you want to quit the application?", vbQuestion + vbYesNo)
If msg <> vbYes Then Exit Sub
End
Case 1
Me.WindowState = vbMinimized
End Select
End Sub

Private Sub controlbtn_MouseEnter(Index As Integer)
controlbtn(Index).GrayScale = lvicNoGrayScale
End Sub

Private Sub controlbtn_MouseExit(Index As Integer)
controlbtn(Index).GrayScale = lvicNTSCPAL
End Sub

Private Sub Form_Resize()
On Error Resume Next
controlbtn(0).Left = (Me.Width - controlbtn(0).Width) - (controlbtn(0).Width \ 3)
controlbtn(1).Left = (Me.Width - controlbtn(0).Width) - (controlbtn(0).Width * 1.5)
lbltime.Left = (Me.Width - lbltime.Width) - 100
lbldate.Left = (Me.Width - lbldate.Width) - 100
keyfrm.Left = (Me.Width \ 2) - (keyfrm.Width \ 2)
tb.Left = mLeft.Width
tb.Width = Me.Width
End Sub

Private Sub keytimer_Timer()
    If GetKeyState(vbKeyCapital) = 0 Then
        keys(0) = "OFF"
        keys(0).ForeColor = &H808080
    Else
        keys(0) = "ON"
        keys(0).ForeColor = vbGreen
    End If
    If GetKeyState(vbKeyNumlock) = 0 Then
        keys(1) = "OFF"
        keys(1).ForeColor = &H808080
    Else
        keys(1) = "ON"
        keys(1).ForeColor = vbGreen
    End If
    If GetKeyState(vbKeyScrollLock) = 0 Then
        keys(2) = "OFF"
        keys(2).ForeColor = &H808080
    Else
        keys(2) = "ON"
        keys(2).ForeColor = vbGreen
    End If
    If GetKeyState(vbKeyInsert) = 0 Then
        keys(3) = "OFF"
        keys(3).ForeColor = &H808080
    Else
        keys(3) = "ON"
        keys(3).ForeColor = vbGreen
    End If
    
    lbltime = Format(Time, "hh:mm AM/PM")
    lbldate = Format(Date, "mm/dd/yyyy")
    
End Sub

Private Sub menubutton_Click(Index As Integer)
On Error Resume Next
For no = 0 To Forms.Count - 1
If Forms(no).Name <> "mainform" Then
 If Forms(no).Name <> frm.Name Then
    Unload Forms(no)
 End If
End If
Next

Select Case Index
Case 0
Case 1
Set tbicon.Picture = src_books.hicon.Picture
tbtitle = "BOOKS"
showform src_books
Case 2
showform src_students
Case 3
showform frmtransaction
Case 4

Case 5
Case 6
Case 7
msg = MsgBox("are you sure you want to logout?", vbQuestion + vbYesNo)
If msg <> vbYes Then Exit Sub
useraccess = ""
currentuser = ""
Unload Me
frmlogin.Show
Case 8
msg = MsgBox("are you sure you want to quit the application?", vbQuestion + vbYesNo)
If msg <> vbYes Then Exit Sub
End
End Select

End Sub

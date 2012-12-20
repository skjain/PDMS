VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frm_er1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ER-1"
   ClientHeight    =   9630
   ClientLeft      =   360
   ClientTop       =   675
   ClientWidth     =   13425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13425
   Begin VB.ComboBox cmbmnt 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   96
      Top             =   120
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Modvate Detail"
      TabPicture(0)   =   "frm_er1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmmodvat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Stock Detail"
      TabPicture(1)   =   "frm_er1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FramePlus1"
      Tab(1).ControlCount=   1
      Begin FramePlusCtl.FramePlus frmmodvat 
         Height          =   8595
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   15161
         BackgroundPictureAlignment=   5
         BorderStyle     =   10
         BackColorGradient=   8438015
         BackColor       =   16777215
         Style           =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " "
         Begin VB.TextBox placurpcess 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   4440
            Width           =   1575
         End
         Begin VB.TextBox rg23acurpcess 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox rg23acurcnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox rg23acuredcs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox rg23ccurcnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox rg23ccuredcs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   3000
            Width           =   1575
         End
         Begin VB.TextBox placurcnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   3960
            Width           =   1575
         End
         Begin VB.TextBox placuredcs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   4920
            Width           =   1575
         End
         Begin VB.TextBox placurhedcs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   5400
            Width           =   1575
         End
         Begin VB.TextBox stxcurcnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   24
            Text            =   "0.00"
            Top             =   5760
            Width           =   1575
         End
         Begin VB.TextBox stxcuredcs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   6240
            Width           =   1575
         End
         Begin VB.TextBox rg23acurhedcs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox rg23ccurhedcs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11520
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   3480
            Width           =   1575
         End
         Begin WelchButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   11040
            TabIndex        =   2
            Top             =   7680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Caption         =   "&Save"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_er1.frx":0038
            cBack           =   -2147483633
         End
         Begin WelchButton.lvButtons_H cmdExit 
            Height          =   495
            Left            =   12120
            TabIndex        =   3
            Top             =   7680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Caption         =   "E&xit"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_er1.frx":0DC2
            cBack           =   -2147483633
         End
         Begin VB.Label ADJBALPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   115
            Top             =   7920
            Width           =   1335
         End
         Begin VB.Label SALPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   114
            Top             =   7320
            Width           =   1335
         End
         Begin VB.Line Line32 
            X1              =   8280
            X2              =   8280
            Y1              =   6600
            Y2              =   8280
         End
         Begin VB.Line Line31 
            X1              =   6720
            X2              =   6720
            Y1              =   6600
            Y2              =   8280
         End
         Begin VB.Line Line30 
            X1              =   5040
            X2              =   5040
            Y1              =   6600
            Y2              =   8280
         End
         Begin VB.Line Line29 
            X1              =   3480
            X2              =   3480
            Y1              =   6600
            Y2              =   8280
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "P.Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   113
            Top             =   6720
            Width           =   1335
         End
         Begin VB.Label RG23ABALPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   112
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label RG23ADBPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   111
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label RG23ACRDPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   110
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label RG23AOPNPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   109
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "P.Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   108
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Line Line28 
            X1              =   1800
            X2              =   13200
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line27 
            X1              =   1800
            X2              =   13200
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label PLABALPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   107
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label PLADBPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   106
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label PLACRDPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   105
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label PLAOPNPCESS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   104
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "P.Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   103
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label PLACRDHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   102
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Line Line9 
            X1              =   120
            X2              =   10200
            Y1              =   7680
            Y2              =   7680
         End
         Begin VB.Line Line26 
            X1              =   13200
            X2              =   13200
            Y1              =   4920
            Y2              =   7200
         End
         Begin VB.Shape Shape5 
            Height          =   735
            Left            =   10920
            Shape           =   4  'Rounded Rectangle
            Top             =   7560
            Width           =   2295
         End
         Begin VB.Shape Shape3 
            Height          =   6495
            Left            =   120
            Top             =   120
            Width           =   13095
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   13200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line2 
            X1              =   3480
            X2              =   3480
            Y1              =   120
            Y2              =   6000
         End
         Begin VB.Line Line3 
            X1              =   5400
            X2              =   5400
            Y1              =   120
            Y2              =   6000
         End
         Begin VB.Line Line4 
            X1              =   7440
            X2              =   7440
            Y1              =   120
            Y2              =   6000
         End
         Begin VB.Line Line5 
            X1              =   9480
            X2              =   9480
            Y1              =   120
            Y2              =   6600
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Opening"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   95
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   94
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Debit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   93
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9720
            TabIndex        =   92
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "RG23-A-II"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   91
            Top             =   960
            Width           =   1095
         End
         Begin VB.Line Line6 
            X1              =   1800
            X2              =   1800
            Y1              =   720
            Y2              =   6000
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Cenvat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   90
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   89
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Hr. Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   88
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Line Line7 
            X1              =   120
            X2              =   13200
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "RG23-C-II"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   87
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Cenvat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   86
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   85
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Hr. Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   84
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Line Line8 
            X1              =   120
            X2              =   13200
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "PLA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   83
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Cenvat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   82
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   81
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Hr. Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   80
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Line Line10 
            X1              =   120
            X2              =   13200
            Y1              =   5760
            Y2              =   5760
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   79
            Top             =   6120
            Width           =   1695
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Cenvat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   78
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   77
            Top             =   6240
            Width           =   1695
         End
         Begin VB.Line Line11 
            X1              =   1800
            X2              =   13200
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line12 
            X1              =   1800
            X2              =   13200
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line13 
            X1              =   1800
            X2              =   13200
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line14 
            X1              =   1800
            X2              =   13200
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line15 
            X1              =   1800
            X2              =   13200
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Line Line16 
            X1              =   1800
            X2              =   13200
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Line Line17 
            X1              =   1800
            X2              =   13200
            Y1              =   6120
            Y2              =   6120
         End
         Begin VB.Shape Shape4 
            Height          =   1695
            Left            =   120
            Top             =   6600
            Width           =   10095
         End
         Begin VB.Line Line18 
            X1              =   1800
            X2              =   10200
            Y1              =   7080
            Y2              =   7080
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Sale of Current Month"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   76
            Top             =   6840
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Cenvat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   75
            Top             =   6720
            Width           =   1815
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   74
            Top             =   6720
            Width           =   1695
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Hr. Edu. Cess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6960
            TabIndex        =   73
            Top             =   6720
            Width           =   1695
         End
         Begin VB.Line Line19 
            X1              =   1800
            X2              =   1800
            Y1              =   6000
            Y2              =   8280
         End
         Begin VB.Line Line20 
            X1              =   3480
            X2              =   3480
            Y1              =   6000
            Y2              =   6600
         End
         Begin VB.Line Line21 
            X1              =   5400
            X2              =   5400
            Y1              =   6000
            Y2              =   6600
         End
         Begin VB.Line Line22 
            X1              =   11400
            X2              =   11400
            Y1              =   120
            Y2              =   7200
         End
         Begin VB.Line Line23 
            X1              =   7440
            X2              =   7440
            Y1              =   6000
            Y2              =   6600
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Duty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8880
            TabIndex        =   72
            Top             =   6720
            Width           =   1095
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cur. Month Deb."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11640
            TabIndex        =   71
            Top             =   240
            Width           =   1335
         End
         Begin VB.Line Line24 
            X1              =   11400
            X2              =   13200
            Y1              =   7200
            Y2              =   7200
         End
         Begin VB.Line Line25 
            X1              =   14760
            X2              =   14760
            Y1              =   6720
            Y2              =   7320
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Total of Cur. Month Db."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   10320
            TabIndex        =   70
            Top             =   6720
            Width           =   1095
         End
         Begin VB.Label RG23AOPNCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   69
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label RG23COPNCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   68
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PLAOPNCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   67
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label STXOPNCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   66
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label RG23ACRDCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   65
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label RG23CCRDCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   64
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PLACRDCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   63
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label SRVCRDCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   62
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label RG23ADBCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   61
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label RG23CDBCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   60
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PLADBCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   59
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label SRVDBCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   58
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label RG23ABALCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   57
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label RG23CBALCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   56
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PLABALCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   55
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label SRVBALCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   54
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label RG23AOPNEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   53
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label RG23COPNEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   52
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label PLAOPNEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   51
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label STXOPNEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   50
            Top             =   6240
            Width           =   1695
         End
         Begin VB.Label RG23ACRDEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   49
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label RG23CCRDEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   48
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label PLACRDEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   47
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label STXCRDEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   46
            Top             =   6240
            Width           =   1695
         End
         Begin VB.Label RG23ADBEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   45
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label RG23CDBEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   44
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label PLADBEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   43
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label STXDBEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   42
            Top             =   6240
            Width           =   1695
         End
         Begin VB.Label RG23ABALEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   41
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label RG23CBALEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   40
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label PLABALEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   39
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label STXBALEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   38
            Top             =   6240
            Width           =   1695
         End
         Begin VB.Label RG23AOPNHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   37
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label RG23COPNHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   36
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label PLAOPNHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   35
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Label RG23ACRDHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   34
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label RG23CCRDHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   33
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label RG23ADBHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   32
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label RG23CDBHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   31
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label PLADBHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   30
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Label RG23ABALHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   29
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label RG23CBALHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   28
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label PLABALHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   27
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Label SALCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   25
            Top             =   7320
            Width           =   1455
         End
         Begin VB.Label SALEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   23
            Top             =   7320
            Width           =   1455
         End
         Begin VB.Label SALHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6840
            TabIndex        =   21
            Top             =   7320
            Width           =   1335
         End
         Begin VB.Label SALTOTDTY 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8280
            TabIndex        =   18
            Top             =   7320
            Width           =   1695
         End
         Begin VB.Label curtotdty 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11520
            TabIndex        =   16
            Top             =   6840
            Width           =   1575
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Bal. To be Adjusted"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   7800
            Width           =   1455
         End
         Begin VB.Label ADJBALDTY 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8280
            TabIndex        =   12
            Top             =   7920
            Width           =   1695
         End
         Begin VB.Label ADJBALHEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6840
            TabIndex        =   10
            Top             =   7920
            Width           =   1335
         End
         Begin VB.Label ADJBALEDCS 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   8
            Top             =   7920
            Width           =   1455
         End
         Begin VB.Label ADJBALCNT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   5
            Top             =   7920
            Width           =   1455
         End
      End
      Begin FramePlusCtl.FramePlus FramePlus1 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   99
         Top             =   360
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   13573
         BackgroundPictureAlignment=   5
         BorderStyle     =   10
         BackColorGradient=   8438015
         BackColor       =   16777215
         Style           =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "n"
         Begin MSFlexGridLib.MSFlexGrid RG1FLEX 
            Height          =   6735
            Left            =   120
            TabIndex        =   101
            Top             =   840
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   11880
            _Version        =   393216
            Cols            =   11
            FixedCols       =   9
            WordWrap        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Shape Shape2 
            Height          =   615
            Left            =   9720
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label LBLHEADING1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "RG-1 Detail"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9840
            TabIndex        =   100
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin WelchButton.lvButtons_H btngen 
      Height          =   495
      Left            =   9000
      TabIndex        =   97
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Generat"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_er1.frx":1214
      cBack           =   -2147483633
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ER-1 For the Month of"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   98
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frm_er1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cur_cenvat As Double
Dim cur_educess As Double
Dim cur_hedcess As Double

Private Sub ADJBALCNT_Change()
  Call CAL_BALADJDTY
End Sub

Private Sub ADJBALDTY_Change()
  Call CAL_BALADJDTY
End Sub

Private Sub ADJBALEDCS_Change()
  Call CAL_BALADJDTY
End Sub

Private Sub ADJBALHEDCS_Change()
  Call CAL_BALADJDTY
End Sub

Private Sub btngen_Click()
  Dim MM As mMonth
  Dim YY As Integer
  
  Dim RS As New ADODB.Recordset
  Set RS = New ADODB.Recordset
  
  
  Select Case cmbmnt.ListIndex
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8
    MM = cmbmnt.ListIndex + 4
   Case 9
    MM = 1
   Case 10
    MM = 2
   Case 11
    MM = 3
  End Select
  Select Case MM
    Case 4, 5, 6, 7, 8, 9, 10, 11, 12
      YY = Year(FSDT)
    Case 1, 2, 3
      YY = Year(FEDT)
  End Select
  Dim start_dt As Date
  Dim end_dt As Date
  
  start_dt = GetMinDate(MM, YY)
  end_dt = GetMaxDate(MM, YY)
  
  If RS.State = 1 Then RS.Close
  RS.Open "select BSTS from billmain where comp='" & compPth & "' and unit='" & UNCD & "' and recstat<>'D' and date>='" & Format(start_dt, "MM/DD/YYYY") & "' and date<='" & Format(end_dt, "MM/DD/YYYY") & "' AND BSTS='P'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    MsgBox "Audit is Pending. Can Not Generate ER-1"
    Unload Me
    Exit Sub
  End If
  
  Call genexcisedtl
  Call GENRG1
  rg23acurcnt.SetFocus
  
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub



Private Sub Form_Activate()
  If Allow_view_only = "Y" Then
     Unload Me
     Exit Sub
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
  cmbmnt.Clear
  cmbmnt.AddItem "April"
  cmbmnt.AddItem "May"
  cmbmnt.AddItem "June"
  cmbmnt.AddItem "July"
  cmbmnt.AddItem "August"
  cmbmnt.AddItem "September"
  cmbmnt.AddItem "October"
  cmbmnt.AddItem "November"
  cmbmnt.AddItem "December"
  cmbmnt.AddItem "January"
  cmbmnt.AddItem "February"
  cmbmnt.AddItem "March"
  cmbmnt.ListIndex = 0
  Call CenterChild(frm_Main, Me)
  Me.Left = Me.Left - 100
  Call SETFLX
End Sub
Private Sub genexcisedtl()
  Dim MM As mMonth
  Dim YY As Integer
  
  Dim RS As New ADODB.Recordset
  Set RS = New ADODB.Recordset
  
  
  Select Case cmbmnt.ListIndex
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8
    MM = cmbmnt.ListIndex + 4
   Case 9
    MM = 1
   Case 10
    MM = 2
   Case 11
    MM = 3
  End Select
  Select Case MM
    Case 4, 5, 6, 7, 8, 9, 10, 11, 12
      YY = Year(FSDT)
    Case 1, 2, 3
      YY = Year(FEDT)
  End Select
  Dim start_dt As Date
  Dim end_dt As Date
  
  start_dt = GetMinDate(MM, YY)
  end_dt = GetMaxDate(MM, YY)
  

    RG23AOPNCNT = 0
    RG23COPNCNT = 0
    PLAOPNCNT = 0
    STXOPNCNT = 0
  
    RG23AOPNEDCS = 0
    RG23COPNEDCS = 0
    PLAOPNEDCS = 0
    STXOPNEDCS = 0
    
    RG23AOPNHEDCS = 0
    RG23COPNHEDCS = 0
    PLAOPNHEDCS = 0


  'Opening of srtatging date rg23-A-II
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS,  isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' AND VTYP<>'EXD' and recstat<>'D' AND TTYP='RG23-A'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23AOPNCNT = RG23AOPNCNT + RS!CENVAT
    RG23AOPNPCESS = RG23AOPNPCESS + RS!PCESS
    RG23AOPNEDCS = RG23AOPNEDCS + RS!EDUCESS
    RG23AOPNHEDCS = RG23AOPNHEDCS + RS!HEDCESS
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' AND VTYP='EXD' and recstat<>'D' AND TTYP='RG23-A'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23AOPNCNT = RG23AOPNCNT - RS!CENVAT
    RG23AOPNPCESS = RG23AOPNPCESS - RS!PCESS
    RG23AOPNEDCS = RG23AOPNEDCS - RS!EDUCESS
    RG23AOPNHEDCS = RG23AOPNHEDCS - RS!HEDCESS
  End If
  
  
  'Credit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='RG23-A' AND VTYP<>'EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23ACRDCNT = RS!CENVAT
    RG23ACRDPCESS = RS!PCESS
    RG23ACRDEDCS = RS!EDUCESS
    RG23ACRDHEDCS = RS!HEDCESS
   Else
    RG23ACRDCNT = 0
    RG23ACRDPCESS = 0
    RG23ACRDEDCS = 0
    RG23ACRDHEDCS = 0
  End If

  'Other Debit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='RG23-A' AND VTYP='EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23ADBCNT = RS!CENVAT
    RG23ADBPCESS = RS!PCESS
    RG23ADBEDCS = RS!EDUCESS
    RG23ADBHEDCS = RS!HEDCESS
   Else
    RG23ADBCNT = 0
    RG23ADBPCESS = 0
    RG23ADBEDCS = 0
    RG23ADBHEDCS = 0
  End If

  'Balance of the Month
   RG23ABALCNT = Val(RG23AOPNCNT) + RG23ACRDCNT - RG23ADBCNT
   RG23ABALPCESS = Val(RG23AOPNPCESS) + RG23ACRDPCESS - RG23ADBPCESS
   RG23ABALEDCS = Val(RG23AOPNEDCS) + RG23ACRDEDCS - RG23ADBEDCS
   RG23ABALHEDCS = Val(RG23AOPNHEDCS) + RG23ACRDHEDCS - RG23ADBHEDCS

'--------------------------------------------------------------------------------------------------
  'Opening of srtatging date rg23-C-II
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND VTYP<>'EXD' AND TTYP='RG23-C'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23COPNCNT = RG23COPNCNT + RS!CENVAT
    RG23COPNEDCS = RG23COPNEDCS + RS!EDUCESS
    RG23COPNHEDCS = RG23COPNHEDCS + RS!HEDCESS
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND VTYP='EXD' AND TTYP='RG23-C'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23COPNCNT = RG23COPNCNT - RS!CENVAT
    RG23COPNEDCS = RG23COPNEDCS - RS!EDUCESS
    RG23COPNHEDCS = RG23COPNHEDCS - RS!HEDCESS
  End If
  
  'Credit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='RG23-C' AND VTYP<>'EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23CCRDCNT = RS!CENVAT
    RG23CCRDEDCS = RS!EDUCESS
    RG23CCRDHEDCS = RS!HEDCESS
   Else
    RG23CCRDCNT = 0
    RG23CCRDEDCS = 0
    RG23CCRDHEDCS = 0
  End If

  'Other Debit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='RG23-C' AND VTYP='EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    RG23CDBCNT = RS!CENVAT
    RG23CDBEDCS = RS!EDUCESS
    RG23CDBHEDCS = RS!HEDCESS
   Else
    RG23CDBCNT = 0
    RG23CDBEDCS = 0
    RG23CDBHEDCS = 0
  End If

  'Balance of the Month
   RG23CBALCNT = Val(RG23COPNCNT) + RG23CCRDCNT - RG23CDBCNT
   RG23CBALEDCS = Val(RG23COPNEDCS) + RG23CCRDEDCS - RG23CDBEDCS
   RG23CBALHEDCS = Val(RG23COPNHEDCS) + RG23CCRDHEDCS - RG23CDBHEDCS


'--------------------------------------------------------------------------------------------------
  'Opening of srtatging date PLA
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' and vtyp<>'EXD' and recstat<>'D' AND TTYP='PLAREG'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    PLAOPNCNT = PLAOPNCNT + RS!CENVAT
    PLAOPNPCESS = PLAOPNPCESS + RS!PCESS
    PLAOPNEDCS = PLAOPNEDCS + RS!EDUCESS
    PLAOPNHEDCS = PLAOPNHEDCS + RS!HEDCESS
  End If
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' and vtyp='EXD' and recstat<>'D' AND TTYP='PLAREG'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    PLAOPNCNT = PLAOPNCNT - RS!CENVAT
    PLAOPNPCESS = PLAOPNPCESS - RS!PCESS
    PLAOPNEDCS = PLAOPNEDCS - RS!EDUCESS
    PLAOPNHEDCS = PLAOPNHEDCS - RS!HEDCESS
  End If
  
  'Credit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='PLAREG' AND VTYP<>'EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    PLACRDCNT = RS!CENVAT
    PLACRDPCESS = RS!PCESS
    PLACRDEDCS = RS!EDUCESS
    PLACRDHEDCS = RS!HEDCESS
   Else
    PLACRDCNT = 0
    PLACRDPCESS = 0
    PLACRDEDCS = 0
    PLACRDHEDCS = 0
  End If

  'Other Debit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, ISNULL(SUM(CESS),0) AS PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='PLAREG' AND VTYP='EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    PLADBCNT = RS!CENVAT
    PLADBPCESS = RS!PCESS
    PLADBEDCS = RS!EDUCESS
    PLADBHEDCS = RS!HEDCESS
   Else
    PLADBCNT = 0
    PLADBEDCS = 0
    PLADBHEDCS = 0
  End If

  'Balance of the Month
   PLABALCNT = Val(PLAOPNCNT) + PLACRDCNT - PLADBCNT
   PLABALPCESS = Val(PLAOPNPCESS) + PLACRDPCESS - PLADBPCESS
   PLABALEDCS = Val(PLAOPNEDCS) + PLACRDEDCS - PLADBEDCS
   PLABALHEDCS = Val(PLAOPNHEDCS) + PLACRDHEDCS - PLADBHEDCS


'Opening of srtatging date SERVICE TAX
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+a_duty),0) as cenvat, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' AND VTYP<>'EXD' and recstat<>'D' AND TTYP='SERVICE TAX'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    STXOPNCNT = STXOPNCNT + RS!CENVAT
    STXOPNEDCS = STXOPNEDCS + RS!EDUCESS + RS!HEDCESS
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+SRVTAX+a_duty),0) as cenvat, isnull(sum(educess+STAX_ED_CESS),0) as educess , " & _
          "isnull(sum(h_ed_cess+STAX_HED_CESS),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date<'" & Format(start_dt, "mm/dd/yyyy") & "' AND VTYP='EXD' and recstat<>'D' AND TTYP='SERVICE TAX'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    STXOPNCNT = STXOPNCNT - RS!CENVAT
    STXOPNEDCS = STXOPNEDCS - RS!EDUCESS - RS!HEDCESS
  End If
  
  
  
  'Credit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(cenvat+SRVTAX+a_duty),0) as cenvat, isnull(sum(EDUCESS+STAX_ED_CESS),0) as educess , " & _
          "isnull(sum(H_ED_CESS+STAX_HED_CESS),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='SERVICE TAX' AND VTYP<>'EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    SRVCRDCNT = RS!CENVAT
    STXCRDEDCS = RS!EDUCESS + RS!HEDCESS
   Else
    SRVCRDCNT = 0
    STXCRDEDCS = 0
  End If

  'Other Debit During the Month
  If RS.State = 1 Then RS.Close
  RS.Open "select isnull(sum(CENVAT+SRVTAX+a_duty),0) as cenvat, isnull(sum(EDUCESS+STAX_ED_CESS),0) as educess , " & _
          "isnull(sum(H_ED_CESS+STAX_HED_CESS),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' AND TTYP='SERVICE TAX' AND VTYP='EXD'", CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    SRVDBCNT = RS!CENVAT
    STXDBEDCS = RS!EDUCESS + RS!HEDCESS
   Else
    SRVDBCNT = 0
    STXDBEDCS = 0
  End If

  'Balance of the Month
   SRVBALCNT = Val(STXOPNCNT) + SRVCRDCNT - SRVDBCNT
   STXBALEDCS = Val(STXOPNEDCS) + STXCRDEDCS - STXDBEDCS



'-----------Formaing For Display ---------------------
   RG23AOPNCNT = Format(RG23AOPNCNT, "############.00")
   RG23ACRDCNT = Format(RG23ACRDCNT, "############.00")
   RG23ADBCNT = Format(RG23ADBCNT, "############.00")
   RG23ABALCNT = Format(RG23ABALCNT, "############.00")
   
   
   RG23AOPNPCESS = Format(RG23AOPNPCESS, "############.00")
   RG23ACRDPCESS = Format(RG23ACRDPCESS, "############.00")
   RG23ADBPCESS = Format(RG23ADBPCESS, "############.00")
   RG23ABALPCESS = Format(RG23ABALPCESS, "############.00")
   
   
   RG23AOPNEDCS = Format(RG23AOPNEDCS, "############.00")
   RG23ACRDEDCS = Format(RG23ACRDEDCS, "############.00")
   RG23ADBEDCS = Format(RG23ADBEDCS, "############.00")
   RG23ABALEDCS = Format(RG23ABALEDCS, "############.00")

   RG23AOPNHEDCS = Format(RG23AOPNHEDCS, "############.00")
   RG23ACRDHEDCS = Format(RG23ACRDHEDCS, "############.00")
   RG23ADBHEDCS = Format(RG23ADBHEDCS, "############.00")
   RG23ABALHEDCS = Format(RG23ABALHEDCS, "############.00")
   
   
   RG23COPNCNT = Format(RG23COPNCNT, "############.00")
   RG23CCRDCNT = Format(RG23CCRDCNT, "############.00")
   RG23CDBCNT = Format(RG23CDBCNT, "############.00")
   RG23CBALCNT = Format(RG23CBALCNT, "############.00")
   
   RG23COPNEDCS = Format(RG23COPNEDCS, "############.00")
   RG23CCRDEDCS = Format(RG23CCRDEDCS, "############.00")
   RG23CDBEDCS = Format(RG23CDBEDCS, "############.00")
   RG23CBALEDCS = Format(RG23CBALEDCS, "############.00")

   RG23COPNHEDCS = Format(RG23COPNHEDCS, "############.00")
   RG23CCRDHEDCS = Format(RG23CCRDHEDCS, "############.00")
   RG23CDBHEDCS = Format(RG23CDBHEDCS, "############.00")
   RG23CBALHEDCS = Format(RG23CBALHEDCS, "############.00")
   

   PLAOPNCNT = Format(PLAOPNCNT, "############.00")
   PLACRDCNT = Format(PLACRDCNT, "############.00")
   PLADBCNT = Format(PLADBCNT, "############.00")
   PLABALCNT = Format(PLABALCNT, "############.00")
   
   PLAOPNPCESS = Format(PLAOPNPCESS, "############.00")
   PLACRDPCESS = Format(PLACRDPCESS, "############.00")
   PLADBPCESS = Format(PLADBPCESS, "############.00")
   PLABALPCESS = Format(PLABALPCESS, "############.00")
   
   PLAOPNEDCS = Format(PLAOPNEDCS, "############.00")
   PLACRDEDCS = Format(PLACRDEDCS, "############.00")
   PLADBEDCS = Format(PLADBEDCS, "############.00")
   PLABALEDCS = Format(PLABALEDCS, "############.00")
   
   PLAOPNHEDCS = Format(PLAOPNHEDCS, "############.00")
   PLACRDHEDCS = Format(PLACRDHEDCS, "############.00")
   PLADBHEDCS = Format(PLADBHEDCS, "############.00")
   PLABALHEDCS = Format(PLABALHEDCS, "############.00")
   
   
   STXOPNCNT = Format(STXOPNCNT, "############.00")
   SRVCRDCNT = Format(SRVCRDCNT, "############.00")
   SRVDBCNT = Format(SRVDBCNT, "############.00")
   SRVBALCNT = Format(SRVBALCNT, "############.00")
   
   STXOPNEDCS = Format(STXOPNEDCS, "############.00")
   STXCRDEDCS = Format(STXCRDEDCS, "############.00")
   STXDBEDCS = Format(STXDBEDCS, "############.00")
   STXBALEDCS = Format(STXBALEDCS, "############.00")
   

   
   
   
'Sale During the Month
If RS.State = 1 Then RS.Close
RS.Open "SELECT ISNULL(SUM(CENVAT+a_duty),0) AS CENVAT,ISNULL(SUM(CESS),0) AS PCESS,ISNULL(SUM(EDUCESS),0) AS EDUCESS,ISNULL(SUM(H_ED_CESS),0) AS HEDCESS FROM EGPMAN " & _
        " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND (VTYP='SAL' OR VTYP='DBN') AND RECSTAT<>'D' AND DATE>='" & Format(start_dt, "MM/DD/YYYY") & "' AND DATE<='" & Format(end_dt, "MM/DD/YYYY") & "' ", CN, adOpenDynamic, adLockOptimistic
        
If Not RS.EOF Then
  SALCNT = RS!CENVAT
  SALPCESS = RS!PCESS
  SALEDCS = RS!EDUCESS
  SALHEDCS = RS!HEDCESS
 Else
  SALCNT = 0
  SALPCESS = 0
  SALEDCS = 0
  SALHEDCS = 0
End If
SALCNT = Format(SALCNT, "##########.00")
SALPCESS = Format(SALPCESS, "##########.00")
SALEDCS = Format(SALEDCS, "##########.00")
SALHEDCS = Format(SALHEDCS, "##########.00")
SALTOTDTY = Format((Val(SALCNT) + Val(SALPCESS) + Val(SALEDCS) + Val(SALHEDCS)), "###########.00")

ADJBALCNT = Format(SALCNT, "##########.00")
ADJBALPCESS = Format(SALPCESS, "##########.00")
ADJBALEDCS = Format(SALEDCS, "##########.00")
ADJBALHEDCS = Format(SALHEDCS, "##########.00")
ADJBALDTY = Format((Val(SALCNT) + Val(SALPCESS) + Val(SALEDCS) + Val(SALHEDCS)), "###########.00")





End Sub

Private Sub rg23acurcnt_GotFocus()
  rg23acurcnt.BackColor = RGB(BRED, BGREEN, BBLUE)
  rg23acurcnt.SelStart = 0
  rg23acurcnt.SelLength = Len(rg23acurcnt)
End Sub

Private Sub rg23acurcnt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub rg23acurcnt_LostFocus()
  rg23acurcnt.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub

Private Sub rg23acurcnt_Validate(CANCEL As Boolean)
  If Val(rg23acurcnt) > RG23ABALCNT Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALCNT = Format(SALCNT - (Val(rg23acurcnt) + Val(rg23ccurcnt) + Val(placurcnt) + Val(stxcurcnt)), "###########.00")
  End If
End Sub
Private Sub rg23acurEDCS_Validate(CANCEL As Boolean)
  If Val(rg23acuredcs) > RG23ABALEDCS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALEDCS = Format(SALEDCS - (Val(rg23acuredcs) + Val(rg23ccuredcs) + Val(placuredcs) + Val(stxcuredcs)), "###########.00")
  End If
End Sub
Private Sub rg23acurhedcs_Validate(CANCEL As Boolean)
  If Val(rg23acurhedcs) > RG23ABALHEDCS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALHEDCS = Format(SALHEDCS - (Val(rg23acurhedcs) + Val(rg23ccurhedcs) + Val(placurhedcs)), "###########.00")
  End If
End Sub
'-----------
Private Sub rg23acuredcs_GotFocus()
  rg23acuredcs.BackColor = RGB(BRED, BGREEN, BBLUE)
  rg23acuredcs.SelStart = 0
  rg23acuredcs.SelLength = Len(rg23acuredcs)
End Sub

Private Sub rg23acuredcs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub rg23acuredcs_LostFocus()
  rg23acuredcs.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub
'----------------
Private Sub rg23acurhedcs_GotFocus()
  rg23acurhedcs.BackColor = RGB(BRED, BGREEN, BBLUE)
  rg23acurhedcs.SelStart = 0
  rg23acurhedcs.SelLength = Len(rg23acurhedcs)
End Sub

Private Sub rg23acurhedcs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub rg23acurhedcs_LostFocus()
  rg23acurhedcs.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub
'------- For RG23-C ---------------------------
Private Sub rg23ccurcnt_GotFocus()
  rg23ccurcnt.BackColor = RGB(BRED, BGREEN, BBLUE)
  rg23ccurcnt.SelStart = 0
  rg23ccurcnt.SelLength = Len(rg23ccurcnt)
End Sub

Private Sub rg23ccurcnt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub rg23ccurcnt_LostFocus()
  rg23ccurcnt.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub

Private Sub rg23ccurcnt_Validate(CANCEL As Boolean)
  If Val(rg23ccurcnt) > RG23CBALCNT Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALCNT = Format(SALCNT - (Val(rg23acurcnt) + Val(rg23ccurcnt) + Val(placurcnt) + Val(stxcurcnt)), "###########.00")
  End If
End Sub
Private Sub rg23ccurEDCS_Validate(CANCEL As Boolean)
  If Val(rg23ccuredcs) > RG23CBALEDCS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALEDCS = Format(SALEDCS - (Val(rg23acuredcs) + Val(rg23ccuredcs) + Val(placuredcs) + Val(stxcuredcs)), "###########.00")
  End If
End Sub
Private Sub rg23ccurhedcs_Validate(CANCEL As Boolean)
  If Val(rg23ccurhedcs) > RG23CBALHEDCS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALHEDCS = Format(SALHEDCS - (Val(rg23acurhedcs) + Val(rg23ccurhedcs) + Val(placurhedcs)), "###########.00")
  End If
End Sub
'-----------
Private Sub rg23ccuredcs_GotFocus()
  rg23ccuredcs.BackColor = RGB(BRED, BGREEN, BBLUE)
  rg23ccuredcs.SelStart = 0
  rg23ccuredcs.SelLength = Len(rg23ccuredcs)
End Sub

Private Sub rg23ccuredcs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub rg23ccuredcs_LostFocus()
  rg23ccuredcs.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub
'----------------
Private Sub rg23ccurhedcs_GotFocus()
  rg23ccurhedcs.BackColor = RGB(BRED, BGREEN, BBLUE)
  rg23ccurhedcs.SelStart = 0
  rg23ccurhedcs.SelLength = Len(rg23ccurhedcs)
End Sub

Private Sub rg23ccurhedcs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub rg23ccurhedcs_LostFocus()
  rg23ccurhedcs.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub

'------- For PLA ---------------------------
Private Sub placurcnt_GotFocus()
  placurcnt.BackColor = RGB(BRED, BGREEN, BBLUE)
  placurcnt.SelStart = 0
  placurcnt.SelLength = Len(placurcnt)
End Sub

Private Sub placurcnt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub placurcnt_LostFocus()
  placurcnt.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub

Private Sub placurcnt_Validate(CANCEL As Boolean)
  If Val(placurcnt) > PLABALCNT Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALCNT = Format(SALCNT - (Val(rg23acurcnt) + Val(rg23ccurcnt) + Val(placurcnt) + Val(stxcurcnt)), "###########.00")
  End If
End Sub
Private Sub placurEDCS_Validate(CANCEL As Boolean)
  If Val(placuredcs) > PLABALEDCS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALEDCS = Format(SALEDCS - (Val(rg23acuredcs) + Val(rg23ccuredcs) + Val(placuredcs) + Val(stxcuredcs)), "###########.00")
  End If
End Sub
Private Sub placurhedcs_Validate(CANCEL As Boolean)
  If Val(placurhedcs) > PLABALHEDCS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALHEDCS = Format(SALHEDCS - (Val(rg23acurhedcs) + Val(rg23ccurhedcs) + Val(placurhedcs)), "###########.00")
  End If
End Sub
'-----------
Private Sub placuredcs_GotFocus()
  placuredcs.BackColor = RGB(BRED, BGREEN, BBLUE)
  placuredcs.SelStart = 0
  placuredcs.SelLength = Len(placuredcs)
End Sub

Private Sub placuredcs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub placuredcs_LostFocus()
  placuredcs.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub
'----------------
Private Sub placurhedcs_GotFocus()
  placurhedcs.BackColor = RGB(BRED, BGREEN, BBLUE)
  placurhedcs.SelStart = 0
  placurhedcs.SelLength = Len(placurhedcs)
End Sub

Private Sub placurhedcs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub placurhedcs_LostFocus()
  placurhedcs.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub


'------- For SRVTAX  ---------------------------
Private Sub stxcurcnt_GotFocus()
  stxcurcnt.BackColor = RGB(BRED, BGREEN, BBLUE)
  stxcurcnt.SelStart = 0
  stxcurcnt.SelLength = Len(stxcurcnt)
End Sub

Private Sub stxcurcnt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub stxcurcnt_LostFocus()
  stxcurcnt.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub

Private Sub stxcurcnt_Validate(CANCEL As Boolean)
  If Val(stxcurcnt) > SRVBALCNT Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALCNT = Format(SALCNT - (Val(rg23acurcnt) + Val(rg23ccurcnt) + Val(placurcnt) + Val(stxcurcnt)), "###########.00")
  End If
End Sub
Private Sub stxcuredcs_Validate(CANCEL As Boolean)
  If Val(stxcuredcs) > STXBALEDCS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALEDCS = Format(SALEDCS - (Val(rg23acuredcs) + Val(rg23ccuredcs) + Val(placuredcs) + Val(stxcuredcs)), "###########.00")
  End If
End Sub

'-----------
Private Sub stxcuredcs_GotFocus()
  stxcuredcs.BackColor = RGB(BRED, BGREEN, BBLUE)
  stxcuredcs.SelStart = 0
  stxcuredcs.SelLength = Len(stxcuredcs)
End Sub

Private Sub stxcuredcs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub

Private Sub stxcuredcs_LostFocus()
  stxcuredcs.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub
Private Sub CAL_BALADJDTY()
  ADJBALDTY = Format(Val(ADJBALCNT) + Val(ADJBALEDCS) + Val(ADJBALHEDCS), "##########.00")
End Sub
Private Sub CAL_PAIDDTY()
  curtotdty = Format(Val(rg23acurcnt) + Val(rg23acurpcess) + Val(rg23acuredcs) + Val(rg23acurhedcs) + Val(rg23ccurcnt) + Val(rg23ccuredcs) + Val(rg23ccurhedcs) + Val(placurcnt) + Val(placurpcess) + Val(placuredcs) + Val(placurhedcs) + Val(stxcurcnt) + Val(stxcuredcs), "##########.00")
End Sub

Private Sub SETFLX()
  RG1FLEX.Clear
  RG1FLEX.Rows = 1
  RG1FLEX.TextMatrix(0, 0) = "Chapter No."
  RG1FLEX.TextMatrix(0, 1) = "Description"
  RG1FLEX.TextMatrix(0, 2) = "Opening Bal"
  RG1FLEX.TextMatrix(0, 3) = "Manufacture"
  RG1FLEX.TextMatrix(0, 4) = "Clearance"
  RG1FLEX.TextMatrix(0, 5) = "Closing Bal"
  RG1FLEX.TextMatrix(0, 6) = "Cenvat"
  RG1FLEX.TextMatrix(0, 7) = "Edu. Cess"
  RG1FLEX.TextMatrix(0, 8) = "Hr. Edu. Cess"
  RG1FLEX.TextMatrix(0, 9) = "Total Duty"
  
  RG1FLEX.ColWidth(0) = 2500
  RG1FLEX.ColWidth(1) = 2500
  RG1FLEX.ColWidth(2) = 1100
  RG1FLEX.ColWidth(3) = 1100
  RG1FLEX.ColWidth(4) = 1100
  RG1FLEX.ColWidth(5) = 1100
  RG1FLEX.ColWidth(6) = 1100
  RG1FLEX.ColWidth(7) = 1100
  RG1FLEX.ColWidth(8) = 1100
  RG1FLEX.ColWidth(9) = 1100
End Sub


Private Sub GENRG1()
  Dim MM As mMonth
  Dim YY As Integer
  
  Dim RS As New ADODB.Recordset
  Set RS = New ADODB.Recordset
  
  
  Select Case cmbmnt.ListIndex
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8
    MM = cmbmnt.ListIndex + 4
   Case 9
    MM = 1
   Case 10
    MM = 2
   Case 11
    MM = 3
  End Select
  Select Case MM
    Case 4, 5, 6, 7, 8, 9, 10, 11, 12
      YY = Year(FSDT)
    Case 1, 2, 3
      YY = Year(FEDT)
  End Select
  Dim start_dt As Date
  Dim end_dt As Date
  
  start_dt = GetMinDate(MM, YY)
  end_dt = GetMaxDate(MM, YY)
    
  
  Dim PPFDATA As New ADODB.Recordset
  Set PPFDATA = New ADODB.Recordset
  
  Dim DPFDATA As New ADODB.Recordset
  Set DPFDATA = New ADODB.Recordset
  
  Dim SQL As String
  
  SQL = Empty
  On Error Resume Next
  'Opening PPF Division Wise
  CN.Execute "DROP VIEW RG1OPNPPF"
  
  SQL = "CREATE VIEW RG1OPNPPF AS SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,DIVMST.EXCOMMODITY AS EXCOM,DIVMST.CHAPTERNO AS CHAP, " & _
      "ISNULL(SUM(NTWGT),0) AS OPNPPFWT,0 AS OPNDPF,0 AS INWPPF,0 AS INWDPF FROM BOXREGISTER INNER JOIN DIVMST ON " & _
      "BOXREGISTER.COMP=DIVMST.COMP AND " & _
      "BOXREGISTER.UNIT=DIVMST.UNIT AND BOXREGISTER.DVCD=DIVMST.CODE WHERE VBDT<'" & Format(start_dt, "MM/DD/YYYY") & "' AND DBCD<>'000006' " & _
      "AND BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & "' AND BOXREGISTER.RECSTAT='A' " & _
      "Group By BOXREGISTER.COMP , BOXREGISTER.unit, BOXREGISTER.DVCD, DIVMST.EXCOMMODITY, DIVMST.CHAPTERNO "
  
  CN.Execute SQL
  'Opening PPF Division Wise Wastage
  
  CN.Execute "DROP VIEW RG1OPNWST"
  
  SQL = "CREATE VIEW RG1OPNWST AS SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,UNTCFG.WEXCO AS EXCOM,UNTCFG.WCHAP AS CHAP, " & _
      "ISNULL(SUM(NTWGT),0) AS OPNPPFWT,0 AS OPNDPF,0 AS INWPPF,0 AS INWDPF FROM BOXREGISTER INNER JOIN UNTCFG ON " & _
      "UNTCFG.COMP=BOXREGISTER.COMP AND UNTCFG.UNIT=BOXREGISTER.UNIT  INNER JOIN DIVMST ON " & _
      "BOXREGISTER.COMP=DIVMST.COMP AND " & _
      "BOXREGISTER.UNIT=DIVMST.UNIT AND BOXREGISTER.DVCD=DIVMST.CODE WHERE VBDT<'" & Format(start_dt, "MM/DD/YYYY") & "' AND DBCD='000006' " & _
      "AND BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & "' AND BOXREGISTER.RECSTAT='A' " & _
      "Group By BOXREGISTER.COMP , BOXREGISTER.unit, BOXREGISTER.DVCD,UNTCFG.WEXCO,UNTCFG.WCHAP"
  
  CN.Execute SQL
  'Period Produciton of finish Goods
  
  CN.Execute "DROP VIEW RG1PRDPPF"
  
  SQL = "CREATE VIEW RG1PRDPPF AS SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,DIVMST.EXCOMMODITY AS EXCOM,DIVMST.CHAPTERNO AS CHAP, " & _
      "0 AS OPNPPFWT,0 AS OPNDPF,ISNULL(SUM(NTWGT),0) AS INWPPF,0 AS INWDPF FROM BOXREGISTER INNER JOIN DIVMST ON " & _
      "BOXREGISTER.COMP=DIVMST.COMP AND " & _
      "BOXREGISTER.UNIT=DIVMST.UNIT AND BOXREGISTER.DVCD=DIVMST.CODE WHERE VBDT>='" & Format(start_dt, "MM/DD/YYYY") & "' AND VBDT<='" & Format(end_dt, "MM/DD/YYYY") & "' AND DBCD<>'000006' " & _
      "AND BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & "' AND BOXREGISTER.RECSTAT='A' " & _
      "Group By BOXREGISTER.COMP , BOXREGISTER.unit, BOXREGISTER.DVCD, DIVMST.EXCOMMODITY, DIVMST.CHAPTERNO "
  
   CN.Execute SQL
  'Period Produciton of Wastage
  
  CN.Execute "DROP VIEW RG1PRDWST"
  
  SQL = "CREATE VIEW RG1PRDWST AS SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,UNTCFG.WEXCO AS EXCOM,UNTCFG.WCHAP AS CHAP, " & _
      "0 AS OPNPPFWT,0 AS OPNDPF,ISNULL(SUM(NTWGT),0) AS INWPPF,0 AS INWDPF FROM BOXREGISTER " & _
      "INNER JOIN UNTCFG ON UNTCFG.COMP=BOXREGISTER.COMP AND UNTCFG.UNIT=BOXREGISTER.UNIT " & _
      "INNER JOIN DIVMST ON " & _
      "BOXREGISTER.COMP=DIVMST.COMP AND " & _
      "BOXREGISTER.UNIT=DIVMST.UNIT AND BOXREGISTER.DVCD=DIVMST.CODE WHERE VBDT>='" & Format(start_dt, "MM/DD/YYYY") & "' AND VBDT<='" & Format(end_dt, "MM/DD/YYYY") & "' AND DBCD='000006' " & _
      "AND BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & "' AND BOXREGISTER.RECSTAT='A' " & _
      "Group By BOXREGISTER.COMP , BOXREGISTER.unit, BOXREGISTER.DVCD, DIVMST.EXCOMMODITY, DIVMST.CHAPTERNO,UNTCFG.WEXCO,UNTCFG.WCHAP "
  
  
  
  
  CN.Execute SQL
  '----------------------------
  
  'Opening DPF Division Wise
  CN.Execute "DROP VIEW RG1OPNDPF"
  
  SQL = "CREATE VIEW RG1OPNDPF AS SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,DIVMST.EXCOMMODITY AS EXCOM,DIVMST.CHAPTERNO AS CHAP, " & _
      "0 OPNPPFWT,ISNULL(SUM(QNTY),0) AS OPNDPF,0 AS INWPPF,0 AS INWDPF FROM SPTRAN INNER JOIN DIVMST ON " & _
      "SPTRAN.COMP=DIVMST.COMP AND " & _
      "SPTRAN.UNIT=DIVMST.UNIT AND SPTRAN.DVCD=DIVMST.CODE WHERE DATE<'" & Format(start_dt, "MM/DD/YYYY") & "' AND DBCD<>'000005' " & _
      "AND SPTRAN.VTYP='DPF' AND SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & "' AND SPTRAN.RECSTAT='A' " & _
      "Group By SPTRAN.COMP , SPTRAN.unit, SPTRAN.DVCD, DIVMST.EXCOMMODITY, DIVMST.CHAPTERNO "
  
  CN.Execute SQL
  'Opening DPF Division Wise Wastage
  
  CN.Execute "DROP VIEW RG1OPNDPFWST"
  
  SQL = "CREATE VIEW RG1OPNDPFWST AS SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,UNTCFG.WEXCO AS EXCOM,UNTCFG.WCHAP AS CHAP, " & _
      "0 AS OPNPPFWT,ISNULL(SUM(QNTY),0) AS OPNDPF,0 AS INWPPF,0 AS INWDPF FROM SPTRAN " & _
      "INNER JOIN UNTCFG ON UNTCFG.COMP=SPTRAN.COMP AND UNTCFG.UNIT=SPTRAN.UNIT " & _
      "INNER JOIN DIVMST ON " & _
      "SPTRAN.COMP=DIVMST.COMP AND " & _
      "SPTRAN.UNIT=DIVMST.UNIT AND SPTRAN.DVCD=DIVMST.CODE WHERE DATE<'" & Format(start_dt, "MM/DD/YYYY") & "' AND DBCD='000005' " & _
      "AND SPTRAN.VTYP='DPF' AND SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & "' AND SPTRAN.RECSTAT='A' " & _
      "Group By SPTRAN.COMP , SPTRAN.unit, SPTRAN.DVCD, UNTCFG.WEXCO,UNTCFG.WCHAP"
  
  CN.Execute SQL
  'Period Produciton of finish Goods
  
  CN.Execute "DROP VIEW RG1PRDDPF"
  
  SQL = "CREATE VIEW RG1PRDDPF AS SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,DIVMST.EXCOMMODITY AS EXCOM,DIVMST.CHAPTERNO AS CHAP, " & _
      "0 AS OPNPPFWT,0 AS OPNDPF,0 AS INWPPF,ISNULL(SUM(QNTY),0) AS INWDPF FROM SPTRAN INNER JOIN DIVMST ON " & _
      "SPTRAN.COMP=DIVMST.COMP AND " & _
      "SPTRAN.UNIT=DIVMST.UNIT AND SPTRAN.DVCD=DIVMST.CODE WHERE DATE>='" & Format(start_dt, "MM/DD/YYYY") & "' AND DATE<='" & Format(end_dt, "MM/DD/YYYY") & "' AND DBCD<>'000005' " & _
      "AND SPTRAN.VTYP='DPF' AND SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & "' AND SPTRAN.RECSTAT='A' " & _
      "Group By SPTRAN.COMP , SPTRAN.unit, SPTRAN.DVCD, DIVMST.EXCOMMODITY, DIVMST.CHAPTERNO "
    
   CN.Execute SQL
  'Period Produciton of Wastage
  
  CN.Execute "DROP VIEW RG1PRDDPFWST"
  
  SQL = "CREATE VIEW RG1PRDDPFWST AS SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,UNTCFG.WEXCO AS EXCOM,UNTCFG.WCHAP AS CHAP, " & _
      "0 AS OPNPPFWT,0 AS OPNDPF,0 AS INWPPF,ISNULL(SUM(QNTY),0) AS INWDPF FROM SPTRAN " & _
      "INNER JOIN UNTCFG ON UNTCFG.COMP=SPTRAN.COMP AND UNTCFG.UNIT=SPTRAN.UNIT " & _
      "INNER JOIN DIVMST ON " & _
      "SPTRAN.COMP=DIVMST.COMP AND " & _
      "SPTRAN.UNIT=DIVMST.UNIT AND SPTRAN.DVCD=DIVMST.CODE WHERE DATE>='" & Format(start_dt, "MM/DD/YYYY") & "' AND DATE<='" & Format(end_dt, "MM/DD/YYYY") & "' AND DBCD='000005' " & _
      "AND SPTRAN.VTYP='DPF' AND SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & "' AND SPTRAN.RECSTAT='A' " & _
      "Group By SPTRAN.COMP , SPTRAN.unit, SPTRAN.DVCD, DIVMST.EXCOMMODITY, DIVMST.CHAPTERNO, UNTCFG.WCHAP,UNTCFG.WEXCO "
  
   CN.Execute SQL
    
    
   CN.Execute "DROP VIEW RG1DATA_TMP"
  
   SQL = "CREATE VIEW RG1DATA_TMP AS SELECT * FROM RG1OPNPPF Union SELECT * FROM RG1OPNWST Union SELECT * FROM RG1PRDPPF Union SELECT * FROM RG1PRDWST " & _
         " Union SELECT * FROM RG1OPNDPF Union SELECT * FROM RG1OPNDPFWST Union SELECT * FROM RG1PRDDPF Union SELECT * FROM RG1PRDDPFWST "
         
   CN.Execute SQL
   
   
   CN.Execute "DROP VIEW RG1DATA"
   SQL = "CREATE VIEW RG1DATA AS SELECT COMP,UNIT,EXCOM,CHAP,DVCD,ISNULL(SUM(OPNPPFWT)-SUM(OPNDPF),0) AS OPN,SUM(INWPPF) AS PPF,SUM(INWDPF) AS DPF FROM " & _
         "RG1DATA_TMP GROUP BY COMP,UNIT,EXCOM,CHAP,DVCD "
         
   CN.Execute SQL
   
   If PPFDATA.State = 1 Then PPFDATA.Close
   PPFDATA.Open "SELECT * FROM RG1DATA", CN, adOpenDynamic, adLockOptimistic
   Call SETFLX
   Do While Not PPFDATA.EOF
    RG1FLEX.Rows = RG1FLEX.Rows + 1
    RG1FLEX.TextMatrix(RG1FLEX.Rows - 1, 0) = PPFDATA!chap & ""
    RG1FLEX.TextMatrix(RG1FLEX.Rows - 1, 1) = PPFDATA!EXCOM & ""
    RG1FLEX.TextMatrix(RG1FLEX.Rows - 1, 2) = Format(PPFDATA!OPN, "############.000")
    RG1FLEX.TextMatrix(RG1FLEX.Rows - 1, 3) = Format(PPFDATA!PPF, "############.000")
    RG1FLEX.TextMatrix(RG1FLEX.Rows - 1, 4) = Format(PPFDATA!DPF, "############.000")
    RG1FLEX.TextMatrix(RG1FLEX.Rows - 1, 5) = Format(PPFDATA!OPN + PPFDATA!PPF - PPFDATA!DPF, "############.000")
    PPFDATA.MoveNext
   Loop
   
   
   If PPFDATA.State = 1 Then PPFDATA.Close
   PPFDATA.Open "SELECT CHAP,EXCO,ISNULL(SUM(CENVAT),0) AS CENVAT,ISNULL(SUM(CESS),0) AS PCESS,ISNULL(SUM(EDUCESS),0) AS EDUCESS,ISNULL(SUM(H_ED_CESS),0) AS H_ED_CESS FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DATE>='" & Format(start_dt, "MM/DD/YYYY") & "' AND DATE<='" & Format(end_dt, "MM/DD/YYYY") & "' AND RECSTAT='A' GROUP BY CHAP,EXCO", CN, adOpenDynamic, adLockOptimistic
   Dim IRW
   IRW = 0
   For IRW = 1 To RG1FLEX.Rows - 1
    PPFDATA.MoveFirst
    Do While Not PPFDATA.EOF
      If Trim(PPFDATA!chap) = Trim(RG1FLEX.TextMatrix(IRW, 0)) Then
        RG1FLEX.TextMatrix(IRW, 6) = Format(PPFDATA!CENVAT, "##########.00")
        RG1FLEX.TextMatrix(IRW, 7) = Format(PPFDATA!PCESS, "##########.00")
        RG1FLEX.TextMatrix(IRW, 8) = Format(PPFDATA!EDUCESS, "##########.00")
        RG1FLEX.TextMatrix(IRW, 9) = Format(PPFDATA!H_ED_CESS, "##########.00")
        RG1FLEX.TextMatrix(IRW, 10) = Format(PPFDATA!CENVAT + PPFDATA!PCESS + PPFDATA!EDUCESS + PPFDATA!H_ED_CESS, "###########.00")
      End If
      PPFDATA.MoveNext
    Loop
   Next
   
End Sub
Private Sub cmdSave_Click()
  Dim MM As mMonth
  Dim YY As Integer
  
  Dim RS As New ADODB.Recordset
  Set RS = New ADODB.Recordset
  
  
  Select Case cmbmnt.ListIndex
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8
    MM = cmbmnt.ListIndex + 4
   Case 9
    MM = 1
   Case 10
    MM = 2
   Case 11
    MM = 3
  End Select
  Select Case MM
    Case 4, 5, 6, 7, 8, 9, 10, 11, 12
      YY = Year(FSDT)
    Case 1, 2, 3
      YY = Year(FEDT)
  End Select
  Dim start_dt As Date
  Dim end_dt As Date
  
  start_dt = GetMinDate(MM, YY)
  end_dt = GetMaxDate(MM, YY)
  CN.Execute "DELETE FROM ER1 WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DATE='" & Format(end_dt, "MM/DD/YYYY") & "'"
  CN.Execute "delete from egpman where comp='" & compPth & "' and unit='" & UNCD & "' and vtyp='EXD' AND (DBCD='RG23-A' OR DBCD='RG23-C' OR DBCD='PLAREG' OR DBCD='SRVREG') AND DATE='" & Format(end_dt, "YYYY/MM/DD") & "'"
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  Dim IRW As Double
  IRW = 0
  For IRW = 1 To RG1FLEX.Rows - 1
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM ER1 WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!Date = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!chap = Mid(Trim(RG1FLEX.TextMatrix(IRW, 0)), 1, 50)
    SAVDAT!exco = Mid(Trim(RG1FLEX.TextMatrix(IRW, 1)), 1, 50)
    SAVDAT!DVCD = "STOCK"
    SAVDAT!OPNQ = Val(RG1FLEX.TextMatrix(IRW, 2))
    SAVDAT!INWQ = Val(RG1FLEX.TextMatrix(IRW, 3))
    SAVDAT!SALQ = Val(RG1FLEX.TextMatrix(IRW, 4))
    SAVDAT!BALQ = Val(RG1FLEX.TextMatrix(IRW, 5))
    SAVDAT!CENVAT = Val(RG1FLEX.TextMatrix(IRW, 6))
    SAVDAT!CESS = Val(RG1FLEX.TextMatrix(IRW, 7))
    SAVDAT!EDUCESS = Val(RG1FLEX.TextMatrix(IRW, 8))
    SAVDAT!HEDCESS = Val(RG1FLEX.TextMatrix(IRW, 9))
    SAVDAT!ASSV = Round(((Val(RG1FLEX.TextMatrix(IRW, 10)) * 100) / 10.3), 0)
    SAVDAT.Update
  Next
  
  
  Dim DLRCRDCEN As Double
  Dim DLRCRDPCESS As Double
  Dim DLRCRDCES As Double
  Dim DLRCRDHCS As Double
  
  Dim IMPCRDCEN As Double
  Dim IMPCRDPCESS As Double
  Dim IMPCRDCES As Double
  Dim IMPCRDHCS As Double
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "select isnull(sum(cenvat),0) as cenvat,isnull(sum(CESS),0) as PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' " & _
          "AND TTYP='RG23-A' AND VTYP='EXC' AND EXTRA4='1st Stage Dealer'", CN, adOpenDynamic, adLockOptimistic
  DLRCRDCEN = 0
  DLRCRDPCESS = 0
  DLRCRDCES = 0
  DLRCRDHCS = 0
  
  IMPCRDCEN = 0
  IMPCRDPCESS = 0
  IMPCRDCES = 0
  IMPCRDHCS = 0
  
  If Not SAVDAT.EOF Then
    DLRCRDCEN = SAVDAT!CENVAT
    DLRCRDPCESS = SAVDAT!PCESS
    DLRCRDCES = SAVDAT!EDUCESS
    DLRCRDHCS = SAVDAT!HEDCESS
  End If
  
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "select isnull(sum(cenvat),0) as cenvat,isnull(sum(CESS),0) as PCESS, isnull(sum(educess),0) as educess , " & _
          "isnull(sum(h_ed_cess),0) as hedcess from egpman where comp='" & compPth & "' and " & _
          "unit='" & UNCD & "' and  date>='" & Format(start_dt, "mm/dd/yyyy") & "' " & _
          "and date<='" & Format(end_dt, "mm/dd/yyyy") & "' and recstat<>'D' " & _
          "AND TTYP='RG23-A' AND VTYP='EXC' AND EXTRA4='Importer'", CN, adOpenDynamic, adLockOptimistic
  
  If Not SAVDAT.EOF Then
    IMPCRDCEN = SAVDAT!CENVAT
    IMPCRDPCESS = SAVDAT!PCESS
    IMPCRDCES = SAVDAT!EDUCESS
    IMPCRDHCS = SAVDAT!HEDCESS
  End If
  
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM ER1 WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!UNIT = UNCD
  SAVDAT!Date = Format(end_dt, "YYYY/MM/DD")
  SAVDAT!DVCD = "EXCDTL"
  
  SAVDAT!RG23AOPNCEN = Val(RG23AOPNCNT)
  SAVDAT!RG23ACRDCEN = Val(RG23ACRDCNT)
  SAVDAT!RG23ADEBCEN = Val(RG23ADBCNT) + Val(rg23acurcnt)
  SAVDAT!RG23ABALCEN = SAVDAT!RG23AOPNCEN + SAVDAT!RG23ACRDCEN - SAVDAT!RG23ADEBCEN
  
  SAVDAT!RG23AOPNPCESS = Val(RG23AOPNPCESS)
  SAVDAT!RG23ACRDPCESS = Val(RG23ACRDPCESS)
  SAVDAT!RG23ADEBPCESS = Val(RG23ADBPCESS) + Val(rg23acurpcess)
  SAVDAT!RG23ABALPCESS = SAVDAT!RG23AOPNPCESS + SAVDAT!RG23ACRDPCESS - SAVDAT!RG23ADEBPCESS
  
  SAVDAT!RG23AOPNEDCS = Val(RG23AOPNEDCS)
  SAVDAT!RG23ACRDEDCS = Val(RG23ACRDEDCS)
  SAVDAT!RG23ADEBEDCS = Val(RG23ADBEDCS) + Val(rg23acuredcs)
  SAVDAT!RG23ABALEDCS = SAVDAT!RG23AOPNEDCS + SAVDAT!RG23ACRDEDCS - SAVDAT!RG23ADEBEDCS
  
  SAVDAT!RG23AOPNHEDCS = Val(RG23AOPNHEDCS)
  SAVDAT!RG23ACRDHEDCS = Val(RG23ACRDHEDCS)
  SAVDAT!RG23ADEBhEDCS = Val(RG23ADBHEDCS) + Val(rg23acurhedcs)
  SAVDAT!RG23ABALHEDCS = SAVDAT!RG23AOPNHEDCS + SAVDAT!RG23ACRDHEDCS - SAVDAT!RG23ADEBhEDCS
  
  'FOR-c
  SAVDAT!RG23COPNCEN = Val(RG23COPNCNT)
  SAVDAT!RG23CCRDCEN = Val(RG23CCRDCNT)
  SAVDAT!RG23CDEBCEN = Val(RG23CDBCNT) + Val(rg23ccurcnt)
  SAVDAT!RG23CBALCEN = SAVDAT!RG23COPNCEN + SAVDAT!RG23CCRDCEN - SAVDAT!RG23CDEBCEN
  
  SAVDAT!RG23COPNEDCS = Val(RG23COPNEDCS)
  SAVDAT!RG23CCRDEDCS = Val(RG23CCRDEDCS)
  SAVDAT!RG23CDEBEDCS = Val(RG23CDBEDCS) + Val(rg23ccuredcs)
  SAVDAT!RG23CBALEDCS = SAVDAT!RG23COPNEDCS + SAVDAT!RG23CCRDEDCS - SAVDAT!RG23CDEBEDCS
  
  SAVDAT!RG23COPNHEDCS = Val(RG23COPNHEDCS)
  SAVDAT!RG23CCRDHEDCS = Val(RG23CCRDHEDCS)
  SAVDAT!RG23CDEBhEDCS = Val(RG23CDBHEDCS) + Val(rg23ccurhedcs)
  SAVDAT!RG23CBALHEDCS = SAVDAT!RG23COPNHEDCS + SAVDAT!RG23CCRDHEDCS - SAVDAT!RG23CDEBhEDCS
  
  'P.L.A
  
  SAVDAT!PLAOPNCEN = Val(PLAOPNCNT)
  SAVDAT!PLACRDCEN = Val(PLACRDCNT)
  SAVDAT!PLADEBCEN = Val(PLADBCNT) + Val(placurcnt)
  SAVDAT!PLABALCEN = SAVDAT!PLAOPNCEN + SAVDAT!PLACRDCEN - SAVDAT!PLADEBCEN
  
  SAVDAT!PLAOPNPCESS = Val(PLAOPNPCESS)
  SAVDAT!PLACRDPCESS = Val(PLACRDPCESS)
  SAVDAT!PLADEBPCESS = Val(PLADBPCESS) + Val(placurpcess)
  SAVDAT!PLABALCEN = SAVDAT!PLAOPNCEN + SAVDAT!PLACRDPCESS + SAVDAT!PLACRDPCESS - SAVDAT!PLADEBPCESS
  
  
  SAVDAT!PLAOPNEDCS = Val(PLAOPNEDCS)
  SAVDAT!PLACRDEDCS = Val(PLACRDEDCS)
  SAVDAT!PLADEBEDCS = Val(PLADBEDCS) + Val(placuredcs)
  SAVDAT!PLABALEDCS = SAVDAT!PLAOPNEDCS + SAVDAT!PLACRDEDCS - SAVDAT!PLADEBEDCS
  
  SAVDAT!PLAOPNHEDCS = Val(PLAOPNHEDCS)
  SAVDAT!PLACRDHEDCS = Val(PLACRDHEDCS)
  SAVDAT!PLADEBhEDCS = Val(PLADBHEDCS) + Val(placurhedcs)
  SAVDAT!PLABALHEDCS = SAVDAT!PLAOPNHEDCS + SAVDAT!PLACRDHEDCS - SAVDAT!PLADEBhEDCS
  
  'Service Tax
  
  SAVDAT!SRVOPNCEN = Val(STXOPNCNT)
  SAVDAT!SRVCRDCEN = Val(SRVCRDCNT)
  SAVDAT!SRVDEBCEN = Val(SRVDBCNT) + Val(stxcurcnt)
  SAVDAT!SRVBALCEN = SAVDAT!PLAOPNCEN + SAVDAT!PLACRDCEN - SAVDAT!PLADEBCEN
  
  SAVDAT!SRVOPNEDCS = Val(PLAOPNEDCS)
  SAVDAT!SRVCRDEDCS = Val(PLACRDEDCS)
  SAVDAT!SRVDEBEDCS = Val(PLADBEDCS) + Val(stxcurcnt)
  SAVDAT!SRVBALEDCS = SAVDAT!SRVOPNEDCS + SAVDAT!SRVCRDEDCS - SAVDAT!SRVDEBEDCS
  
  SAVDAT!DLRCRDCEN = DLRCRDCEN
  SAVDAT!DLRCRDPCESS = DLRCRDPCESS
  SAVDAT!DLRCRDECS = DLRCRDCES
  SAVDAT!DLRCRDHCS = DLRCRDHCS
  
  
  SAVDAT!IMPCRDCEN = IMPCRDCEN
  SAVDAT!IMPCRDPCESS = IMPCRDPCESS
  SAVDAT!IMPCRDECS = IMPCRDCES
  SAVDAT!IMPCRDHCS = IMPCRDHCS
  
  SAVDAT.Update
  Dim mnt As Double
  Dim yr As Double
  mnt = MM
  yr = YY
  'Add Records In EGP MAN For Debit Entry
  If Val(rg23acurcnt) <> 0 And Val(rg23acuredcs) <> 0 And Val(rg23acurhedcs) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXD"
    SAVDAT!SRNO = CStr(end_dt)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!dbcd = "RG23-A"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!chln = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!CHDT = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!TTYP = "RG23-A"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(rg23acurcnt)
    SAVDAT!CESS = Val(rg23acurpcess)
    SAVDAT!EDUCESS = Val(rg23acuredcs)
    SAVDAT!H_ED_CESS = Val(rg23acurhedcs)
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "ER-1 For the Month : " + cmbmnt
    SAVDAT.Update
  End If
  
  
  
  If Val(rg23ccurcnt) <> 0 And Val(rg23ccuredcs) <> 0 And Val(rg23ccurhedcs) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXD"
    SAVDAT!SRNO = CStr(end_dt)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!dbcd = "RG23-C"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!chln = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!CHDT = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!TTYP = "RG23-C"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(rg23ccurcnt)
    SAVDAT!EDUCESS = Val(rg23ccuredcs)
    SAVDAT!H_ED_CESS = Val(rg23ccurhedcs)
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "ER-1 For the Month : " + cmbmnt
    SAVDAT.Update
  End If
  
  
  If Val(placurcnt) <> 0 And Val(placuredcs) <> 0 And Val(placurhedcs) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXD"
    SAVDAT!SRNO = CStr(end_dt)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!dbcd = "PLAREG"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!chln = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!CHDT = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!TTYP = "PLAREG"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(placurcnt)
    SAVDAT!CESS = Val(placurpcess)
    SAVDAT!EDUCESS = Val(placuredcs)
    SAVDAT!H_ED_CESS = Val(placurhedcs)
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "ER-1 For the Month : " + cmbmnt
    SAVDAT.Update
  End If
  
  
  If Val(stxcurcnt) <> 0 And Val(stxcuredcs) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXD"
    SAVDAT!SRNO = CStr(end_dt)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!dbcd = "SRVREG"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!chln = nstr(mnt, 2, 0) + "-" + nstr(yr, 4, 0)
    SAVDAT!CHDT = Format(end_dt, "YYYY/MM/DD")
    SAVDAT!TTYP = "SERVICE TAX"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(stxcurcnt)
    SAVDAT!EDUCESS = Val(stxcuredcs)
    SAVDAT!H_ED_CESS = 0
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "ER-1 For the Month : " + cmbmnt
    SAVDAT.Update
  End If
  
  MsgBox "UPDATE SUCCESSFUL"
  Call ClsData(Me)
End Sub


Private Sub rg23acurpcess_GotFocus()
  rg23acurpcess.BackColor = RGB(BRED, BGREEN, BBLUE)
  rg23acurpcess.SelStart = 0
  rg23acurpcess.SelLength = Len(rg23acurpcess)
End Sub
Private Sub rg23acurpcess_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub
Private Sub rg23acurpcess_LostFocus()
  rg23acurpcess.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub
Private Sub rg23acurpcess_Validate(CANCEL As Boolean)
  If Val(rg23acurpcess) > RG23ABALPCESS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALPCESS = Format(SALPCESS - (Val(rg23acurpcess) + Val(placurpcess)), "###########.00")
  End If
End Sub
Private Sub placurpcess_GotFocus()
  placurpcess.BackColor = RGB(BRED, BGREEN, BBLUE)
  placurpcess.SelStart = 0
  placurpcess.SelLength = Len(placurpcess)
End Sub
Private Sub placurpcess_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
  Call CAL_PAIDDTY
End Sub
Private Sub placurpcess_LostFocus()
  placurpcess.BackColor = vbWhite
  Call CAL_PAIDDTY
End Sub
Private Sub placurpcess_Validate(CANCEL As Boolean)
  If Val(placurpcess) > PLABALPCESS Then
    MsgBox "Invalid Figure"
    CANCEL = True
  End If
  If CANCEL = False Then
    ADJBALPCESS = Format(SALPCESS - (Val(rg23acurpcess) + Val(placurpcess)), "###########.00")
  End If
End Sub

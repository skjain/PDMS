VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmJobDispatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box Dispatch Without Order"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   11385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7683.554
   ScaleMode       =   0  'User
   ScaleWidth      =   30395.35
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   50
      Top             =   9480
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrTool 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1800
         Top             =   240
      End
      Begin VB.Label lblToolTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   120
         TabIndex        =   51
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   120
      Top             =   9120
   End
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   7755
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13679
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   12632319
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
      Begin VB.TextBox txtVHCL 
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   7320
         Width           =   1095
      End
      Begin VB.TextBox txtTransport 
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   7320
         Width           =   3495
      End
      Begin VB.TextBox txtLRNO 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   22
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CheckBox chkReturnable 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Is Cops Returnable ?"
         Height          =   495
         Left            =   40
         TabIndex        =   18
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox TXTGRNNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   6720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TXTRMRK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8040
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cmbPackingType 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "frmJobDispatch.frx":0000
         Left            =   2040
         List            =   "frmJobDispatch.frx":0002
         TabIndex        =   6
         Tag             =   "0"
         Text            =   "cmbPackingType"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox TXTRATE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         MaxLength       =   200
         TabIndex        =   15
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1110
         Width           =   3615
      End
      Begin VB.TextBox txtPCOD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtLTNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtCONSINEE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1110
         Width           =   3135
      End
      Begin VB.TextBox TXTGRAD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox TXTITEM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   285
         Left            =   9360
         TabIndex        =   7
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57278465
         CurrentDate     =   39347
      End
      Begin MSComctlLib.ListView lstBox 
         Height          =   3255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   10485760
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Box No."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cops"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Net Wt."
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Twist"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Gross Wt."
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Tare Wt."
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Pkg Date"
            Object.Width           =   2207
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   5293
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "PK_STCOD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ISRETURNABLE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TOP"
            Object.Width           =   0
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   9240
         TabIndex        =   3
         Top             =   6480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "E&dit"
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
         Image           =   "frmJobDispatch.frx":0004
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   12000
         TabIndex        =   4
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "&Delete"
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
         Image           =   "frmJobDispatch.frx":039E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   7080
         TabIndex        =   1
         Top             =   6480
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
         Image           =   "frmJobDispatch.frx":0738
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8160
         TabIndex        =   2
         Top             =   6480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frmJobDispatch.frx":14C2
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   10320
         TabIndex        =   5
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
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
         Image           =   "frmJobDispatch.frx":1914
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   6000
         TabIndex        =   0
         Top             =   6480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "&Add"
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
         Image           =   "frmJobDispatch.frx":1D66
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   13080
         TabIndex        =   52
         Top             =   6000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Save/&Print"
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
         Image           =   "frmJobDispatch.frx":2100
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker LRDT 
         Height          =   330
         Left            =   2880
         TabIndex        =   23
         Top             =   7320
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57278465
         CurrentDate     =   38429
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F6 For Select All && F7 For De-Select All"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Tag             =   "S"
         Top             =   2040
         Width           =   7935
      End
      Begin VB.Label LBLLRDT 
         BackStyle       =   0  'Transparent
         Caption         =   "&L.R Dt."
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
         Left            =   2160
         TabIndex        =   80
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label LBLVHCL 
         BackStyle       =   0  'Transparent
         Caption         =   "&Vehicle No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   7320
         Width           =   1215
      End
      Begin VB.Label LBLTRCD 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transport"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   7320
         Width           =   1215
      End
      Begin VB.Label LBLLR 
         BackStyle       =   0  'Transparent
         Caption         =   "&L.R No."
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
         Left            =   120
         TabIndex        =   21
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Stock After Dispatch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8160
         TabIndex        =   79
         Top             =   5550
         Width           =   3015
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Selected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4680
         TabIndex        =   78
         Top             =   5550
         Width           =   3015
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Available Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1440
         TabIndex        =   77
         Top             =   5550
         Width           =   3015
      End
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   7920
         X2              =   7920
         Y1              =   5520
         Y2              =   6360
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   4560
         X2              =   4560
         Y1              =   5520
         Y2              =   6360
      End
      Begin VB.Label txtRMNCTRN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8040
         TabIndex        =   76
         Top             =   6000
         Width           =   885
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8160
         TabIndex        =   75
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label txtRMNCOPs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9000
         TabIndex        =   74
         Top             =   6000
         Width           =   900
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Wt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   10560
         TabIndex        =   73
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Cops"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   9240
         TabIndex        =   72
         Top             =   5760
         Width           =   465
      End
      Begin VB.Label txtRMNNTWT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9960
         TabIndex        =   71
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label txtTTLCTRN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   70
         Top             =   6000
         Width           =   765
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1080
         TabIndex        =   69
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label txtTTLCOPs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2280
         TabIndex        =   68
         Top             =   6000
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Net Wt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3240
         TabIndex        =   67
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cops"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2160
         TabIndex        =   66
         Top             =   5760
         Width           =   945
      End
      Begin VB.Label txtTTLNTWT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   65
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label LabelO 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "{O}"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3480
         TabIndex        =   64
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label LabelZ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "{Z}"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2160
         TabIndex        =   63
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label LBLOWGT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3540
         TabIndex        =   62
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label LBLZWGT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2100
         TabIndex        =   61
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label LBLSWGT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   660
         TabIndex        =   60
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label txtNTWT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6600
         TabIndex        =   59
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cops"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5520
         TabIndex        =   58
         Top             =   5760
         Width           =   945
      End
      Begin VB.Label LBL0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   57
         Top             =   6720
         Width           =   435
      End
      Begin VB.Label LBLZ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   56
         Top             =   6720
         Width           =   435
      End
      Begin VB.Label LBLS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   55
         Top             =   6720
         Width           =   435
      End
      Begin VB.Label LabelS 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "{S}"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   720
         TabIndex        =   54
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   6480
         Width           =   495
      End
      Begin VB.Label LBLGRN 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GRN No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Tag             =   "S"
         Top             =   6480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LBLHEAD 
         BackStyle       =   0  'Transparent
         Caption         =   "       Challan No ."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   49
         Tag             =   "0"
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label LBLCHLN 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9360
         TabIndex        =   48
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Dispatch :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label LBLCHDT 
         BackStyle       =   0  'Transparent
         Caption         =   "       Challan Date :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   46
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label LBLDIV 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   45
         Top             =   120
         Width           =   3255
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape BORDER2 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   5040
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label LBLHEADING2 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Challan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carton Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3600
         TabIndex        =   42
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4680
         TabIndex        =   41
         Tag             =   "S"
         Top             =   915
         Width           =   2055
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7920
         TabIndex        =   40
         Tag             =   "S"
         Top             =   915
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Tag             =   "S"
         Top             =   915
         Width           =   1455
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LotNo."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Tag             =   "S"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LBLCFG2 
         BackStyle       =   0  'Transparent
         Caption         =   "SubGrade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Tag             =   "S"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7080
         TabIndex        =   36
         Tag             =   "S"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1215
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   11175
      End
      Begin VB.Label LBLCFG 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Tag             =   "S"
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Tag             =   "S"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   8640
         TabIndex        =   33
         Tag             =   "S"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Net Wt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6600
         TabIndex        =   32
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label txtCOPs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5640
         TabIndex        =   31
         Top             =   6000
         Width           =   900
      End
      Begin VB.Label lblNTWT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4800
         TabIndex        =   30
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label txtCTRN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   29
         Top             =   6000
         Width           =   885
      End
      Begin VB.Shape BORDER3 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3975
      End
      Begin VB.Shape ShapeSZO 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   720
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Width           =   4455
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   705
         Left            =   4560
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmJobDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim INIT As Boolean
Dim GRD_SHD_REQ As Boolean
Public CHLNTYP As String
Dim ALLOWEDITDEL As Boolean
Dim DIVCODE As String
Dim SAVEFLAG As Boolean
Dim EDITFLAG As Boolean
Dim ORDREQ As Boolean
Dim TTLCOPS As Long
Dim TTLBOXES As Long
Dim TTLQTY As Double, TTLGRSWGT As Double, TTLTAREWGT As Double
Dim SPARTY As String, SCONSINEE As String, SADD As String, grad As String, SUBGRD As String, SGRD As String, SITEM As String
Dim M_BRCD As String
Public VTCD As String
Public ORDN As String
Public DONO As String
Public M_DBCD As String
Dim ADD_SRNO As Long
Dim M_VHCD As String, M_TRCD As String
Public chln As String

Private Sub cmbPackingType_Click()
   LBLGRN.Visible = False
   TXTGRNNO.Visible = False
If InStr(1, UCase(cmbPackingType.Text), "JOB CHALLAN") <> 0 Then
   Me.Caption = "Box Dispatch (Job Challan) "
   LBLHEAD = "Job Challan No ."
   LBLCHDT = "Job Challan Date :"
   LBLGRN.Visible = True
   TXTGRNNO.Visible = True
ElseIf InStr(1, UCase(cmbPackingType.Text), "CAPTIVE") <> 0 Then
   Me.Caption = "Box Dispatch (Captive Challan) "
   LBLHEAD = "Captive Challan No ."
   LBLCHDT = "Captive Challan Date :"
ElseIf InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then
   Me.Caption = "Box Dispatch (Export Challan) "
   LBLHEAD = "Export Challan No ."
   LBLCHDT = "Export Challan Date :"
Else
   Me.Caption = "Box Dispatch (Sale Challan) "
   LBLHEAD = "    Challan No ."
   LBLCHDT = "    Challan Date :"
End If

If cmdAdd.Enabled = False Then
   txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
   LBLS = 0:  LBLZ = 0: LBL0 = 0
   lstBox.ListItems.Clear
   Call GenerateBoxList
End If

   TXTVBDT = Now
   Call SetGlobal
   LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
End Sub

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmdAdd_Click()
 zoomflag = False
 btn_sts (False)
 SAVEFLAG = True
 EDITFLAG = True
 If cmbPackingType.Enabled Then cmbPackingType.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Call CLEARDATA
    cmbPackingType.Enabled = True
    Call btn_sts(True)
    chkReturnable.Value = 0
    txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
    LBLS = 0:  LBLZ = 0: LBL0 = 0
    LBLSWGT = 0:  LBLZWGT = 0: LBLOWGT = 0
    
    lstBox.ListItems.Clear
    If zoomflag = True Then
        Call CMDEXIT_Click
        Exit Sub
    End If
    TXTVBDT = Now
    TXTVBDT.Enabled = True
    LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
    Call CLEARDATA
    If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 0
    If cmdExit.Enabled Then cmdExit.SetFocus
End Sub

Private Sub cmdDelete_Click()
Exit Sub
Dim SQL
ALLOWEDITDEL = True: SAVEFLAG = False: EDITFLAG = False: VTCD = Empty: chln = Empty
frmJobDispatchList.DIVCODE = DIVCODE
frmJobDispatchList.DIVNAME = LBLDIV
Call SetGlobal
frmJobDispatchList.VTCD = VTCD

btn_sts (False)
frmJobDispatchList.Show 1
  
If ALLOWEDITDEL = False Then
   MsgBox "Can't be Delete ", vbInformation
   Exit Sub
End If

If VTCD = Empty And chln = Empty Then Exit Sub

Dim AYS
   AYS = MsgBox("Are you sure to delete this Service GRN ", vbYesNo)
   If AYS = vbYes Then
      CN.BeginTrans
   'STEP:1
   If CHLNTYP = "JOB CHALLAN" Then
      SQL = "UPDATE JOBIN SET RECSTAT='D' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
      "' AND VTYP='DPF' AND DBCD='" & VTCD & "' AND VBNO = '" & LBLCHLN & "'"
      CN.Execute SQL
   End If

    SQL = "UPDATE SPTRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND DVCD ='" & DIVCODE & _
    "' AND DBCD ='" & VTCD & "'  AND VBNO ='" & LBLCHLN & "' AND VTYP='DPF' AND RECSTAT='A'"

    CN.Execute SQL

    SQL = "UPDATE PKGMAN SET RECSTAT ='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
    "' AND DVCD ='" & DIVCODE & "' AND DBCD ='" & VTCD & "' AND VTYP = 'DPF' AND PKG_STCOD='000000' AND RECSTAT='A'"

    CN.Execute SQL

    SQL = "UPDATE PKGSTK SET RECSTAT ='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
    "' AND DVCD='" & DIVCODE & "' AND DBCD='" & VTCD & "' AND VTYP='DPF' AND CHLN='" & LBLCHLN & "' AND RECSTAT='A'"

    CN.Execute SQL

    SQL = "UPDATE BOXREGISTER SET VTYP='PPF',RVBNO=NULL,RVBDT= NULL,RDBC = NULL,RVTYP = NULL WHERE COMP='" & compPth & _
    "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND RVBNO='" & LBLCHLN & "' AND RDBC = '" & VTCD & _
    "' AND RVTYP='DPF' AND VTYP='DPF'"

    CN.Execute SQL
    Call DAILYSTATUS("DPF", GetCode("ACCMST", txtpcod, "NAME", "CODE"), VTCD, Val(txtNTWT), LBLCHLN, 0, cUName, "D", Now, TXTVBDT)
CN.CommitTrans
MsgBox "Your Challan No. : " & LBLCHLN & " successfully deleted."

txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
LBLS = 0:  LBLZ = 0: LBL0 = 0
lstBox.ListItems.Clear
Call CLEARDATA
End If

Call cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("0017", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  SAVEFLAG = False
  EDITFLAG = False
  chln = Empty
  
  frmJobDispatchList.DIVCODE = DIVCODE
  frmJobDispatchList.DIVNAME = LBLDIV
  
  Call SetGlobal
  
  frmJobDispatchList.VTCD = VTCD
  frmJobDispatchList.Show 1
     
  If IsSaleExist Then
     MsgBox "Sale Bill Exist Against this Challan."
     txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
     LBLS = 0:  LBLZ = 0: LBL0 = 0
     lstBox.ListItems.Clear
     Call CLEARDATA
     Call cmdCancel_Click
     Exit Sub
  End If
      
  EDITFLAG = True
      
  If chln <> Empty Then
     LBLCHLN = chln
     btn_sts (False)
     cmbPackingType.Enabled = False
     If txtpcod.Enabled Then txtpcod.SetFocus
     TXTVBDT.Enabled = False
  Else
     btn_sts (True)
     cmdAdd.SetFocus
  End If
   
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim INDEX As Long
Dim FLAG As Boolean
Dim SLIP As String
Dim COPS As Double
Dim PCS As Double
Dim SQL As String
Dim RETURNCOPS As Long
Dim WDNPLY As Long
Dim PVCPLY As Long
Dim FIBPLY As Long
Dim TOPBOTTOM As Long
Dim AMOUNT As String

TOPBOTTOM = 0

If SAVEFLAG Then 'NEW
   FLAG = False
   For INDEX = 1 To lstBox.ListItems.COUNT
     If lstBox.ListItems(INDEX).Checked = True Then: FLAG = True: Exit For
   Next
    
   If FLAG = False Then Exit Sub
End If

If Not CHKSAVEDATA Then Exit Sub

For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then
     TOPBOTTOM = TOPBOTTOM + Val(lstBox.ListItems(INDEX).SubItems(10))
     RETURNCOPS = RETURNCOPS + Val(lstBox.ListItems(INDEX).SubItems(1))
  End If
Next

Call SetGlobal

Dim NSQL As String
Dim MSGS As String: MSGS = "Unit"

If SAVEFLAG Then
   LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
   
   NSQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & _
           "' AND VTYP='DPF' AND DBCD='" & VTCD & "' AND VBNO = '" & LBLCHLN & "' "
   
   If UNT_DIVSERIES_REQ = "Y" Then
      NSQL = NSQL & " AND DVCD='" & DIVCODE & "' "
      MSGS = "Division"
   End If
   
   If RS.State Then RS.Close
   RS.Open NSQL, CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      MsgBox "Challan No. " & LBLCHLN & " Already Exist. Check Last No. In " & MSGS & " Configuration", vbCritical
      Exit Sub
   End If
   RS.Close
End If

If RoundOffReq Then
   AMOUNT = TTLQTY * Val(txtRate)
   AMOUNT = nstr(Round(AMOUNT, 0), 12, 2)
Else
  AMOUNT = TTLQTY * Val(txtRate)
End If

CN.BeginTrans


If RETURNCOPS > 0 And chkReturnable.Value = 1 Then

   SQL = "DELETE FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
         "' AND DBCD='" & VTCD & "' AND VTYP='DPF' AND CHLN='" & LBLCHLN & "' AND RECSTAT='A' "
   
   CN.Execute SQL
      
   SQL = "INSERT INTO PKGSTK(COMP,UNIT,DVCD,DBCD,VTYP,CHLN,DATE,PCOD,DCOD,ADDRESS,OPER,"
   SQL = SQL & "QNTY,RECSTAT) VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & "','" & VTCD & "', 'DPF','" & LBLCHLN & _
   "','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & SPARTY & "','" & SCONSINEE & _
   "','" & SADD & "','-','" & RETURNCOPS & "','A')"

   CN.Execute SQL
   
End If

If SAVEFLAG Then ''INSERT MODE

LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)

If InStr(1, UCase(cmbPackingType.Text), "JOB CHALLAN") <> 0 Then
   SQL = "INSERT INTO JOBIN(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
   SQL = SQL & "DCOD,ADDRESS,LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
   SQL = SQL & "RECSTAT,COPS,EXTRA1,GRNNO) VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
   "','" & VTCD & "','DPF','" & LBLCHLN & "','" & LBLCHLN & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & SPARTY & "','" & SPARTY & "','" & SCONSINEE & _
   "','" & SADD & "','" & txtLTNo & "','" & SITEM & "','" & SGRD & _
   "','" & SUBGRD & "','" & TTLBOXES & "','" & TTLQTY & "'," & txtRate & "," & AMOUNT & _
   ",'Q','N','" & cUName & "','-','A','" & TTLCOPS & "','" & Trim(TXTRMRK) & "','" & Trim(TXTGRNNO) & "')"
    
   CN.Execute SQL
End If

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,BRCD,"
SQL = SQL & "DCOD,ADDRESS,LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,GWGT,TWGT,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA4,GRNNO,ISRETURNABLE,LRNO,LRDT,VEHICALNO,TRCD)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','DPF','" & VTCD & "','" & LBLCHLN & "','" & LBLCHLN & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & SPARTY & "','" & SPARTY & "','" & M_BRCD & "','" & SCONSINEE & _
"','" & SADD & "','" & txtLTNo & "','" & SITEM & "','" & SGRD & _
"','" & SUBGRD & "','" & TTLBOXES & "','" & TTLQTY & "','" & TTLGRSWGT & "','" & TTLTAREWGT & _
"'," & txtRate & "," & AMOUNT & _
",'Q','N','" & cUName & "','-','A','" & TTLCOPS & _
"','" & Trim(TXTRMRK) & "','" & Trim(TXTGRNNO) & "','" & IIf(chkReturnable.Value = 1, "Y", "N") & _
"','" & Trim(TXTLRNO) & "','" & Format(LRDT, "YYYY/MM/DD") & "','" & M_VHCD & "','" & M_TRCD & "')"

CN.Execute SQL

SQL = "INSERT INTO PKGMAN (COMP,UNIT,DVCD,DBCD,VTYP,SRNO,SRCH,DATE,SLIPNO,PKG_STCOD,"
SQL = SQL & "LOTNO,FINITMCOD,GRAD,SUBGRAD,QNTY,SYSR,[USER],OPER,RECSTAT) VALUES "
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & DIVCODE & "','" & VTCD & "','DPF',"
SQL = SQL & "'1','1','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & LBLCHLN & "','000000','" & txtLTNo & _
"','" & SITEM & "','" & SGRD & "','" & SUBGRD & "','" & TTLQTY & "','N','" & cUName & "','-','A')"

CN.Execute SQL

Dim UPSQL As String
UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & LBLCHLN & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
         "' AND VTYP='DPF' AND CODE='" & VTCD & "' AND FYCD='" & FYCD & "' "

If UNT_DIVSERIES_REQ = "Y" Then
   UPSQL = UPSQL & " AND DVCD='" & DIVCODE & "' "
End If
 
CN.Execute UPSQL

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & LBLCHLN & "',RVBDT= '" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "',RDBC = '" & VTCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
   "' AND PKG_STCOD='" & Trim(lstBox.ListItems(INDEX).SubItems(8)) & "' AND (VTYP='PPF' OR VTYP='OPN') AND " & _
   "VBNO='" & lstBox.ListItems(INDEX).Text & "'"
   CN.Execute SQL
 End If
Next INDEX

Call DAILYSTATUS("DPF", GetCode("ACCMST", txtpcod, "NAME", "CODE"), VTCD, Val(txtNTWT), LBLCHLN, 0, cUName, "N", Now, TXTVBDT)

CN.CommitTrans

Else  'EDIT UPDATE MODE

If InStr(1, UCase(cmbPackingType.Text), "JOB CHALLAN") <> 0 Then

SQL = "UPDATE JOBIN SET CHDT ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',DRAC='" & SPARTY & "',DCOD='" & SCONSINEE & _
"',ADDRESS='" & SADD & "',ICOD='" & SITEM & "',PCES='" & TTLBOXES & "',QNTY='" & TTLQTY & "',RATE=" & txtRate & _
",AMNT=" & AMOUNT & ",GRAD='" & SGRD & "',SUBGRD='" & SUBGRD & _
"',LTNO='" & txtLTNo & "',EXTRA1='" & TXTRMRK & "',COPS='" & TTLCOPS & _
"',GRNNO='" & Trim(TXTGRNNO) & "' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND VTYP='DPF' AND DBCD='" & VTCD & "' AND VBNO = '" & LBLCHLN & "'"

CN.Execute SQL
End If

SQL = "UPDATE SPTRAN SET CHDT ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',PCES='" & TTLBOXES & "',QNTY='" & TTLQTY & _
"',GWGT='" & TTLGRSWGT & "',TWGT='" & TTLTAREWGT & _
"',COPS = '" & TTLCOPS & "',PCOD = '" & SPARTY & "',DRAC = '" & SPARTY & _
"', DCOD = '" & SCONSINEE & "', ADDRESS= '" & SADD & "', LTNO= '" & txtLTNo & "', ICOD= '" & SITEM & _
"', grad= '" & SGRD & "', SUBGRD= '" & SUBGRD & "', RATE= '" & txtRate & "', AMNT = '" & AMOUNT & _
"',GRNNO='" & Trim(TXTGRNNO) & "',EXTRA4='" & Trim(TXTRMRK) & _
"',ISRETURNABLE ='" & IIf(chkReturnable.Value = 1, "Y", "N") & "',LRNO='" & TXTLRNO & "',LRDT ='" & Format(LRDT, "YYYY/MM/DD") & "',VEHICALNO= '" & M_VHCD & "',TRCD ='" & M_TRCD & "' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND DVCD ='" & DIVCODE & _
"' AND DBCD ='" & VTCD & "'  AND VBNO ='" & LBLCHLN & "' AND VTYP='DPF' AND RECSTAT='A'"

CN.Execute SQL

SQL = "UPDATE PKGMAN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',QNTY = '" & TTLQTY & _
"',LOTNO ='" & txtLTNo & "', FINITMCOD ='" & SITEM & _
"' ,grad='" & SGRD & "' ,SUBGRAD = '" & SUBGRD & "' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD ='" & DIVCODE & "' AND DBCD ='" & VTCD & "' AND VTYP = 'DPF' AND SLIPNO ='" & LBLCHLN & "'  AND PKG_STCOD='000000' AND RECSTAT='A'"

CN.Execute SQL

Dim L As Long
SQL = "UPDATE BOXREGISTER SET VTYP=PVTYP,RVBNO=NULL,RVBDT= NULL,RDBC = NULL,RVTYP = NULL WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND RVBNO='" & LBLCHLN & "' AND RDBC = '" & VTCD & _
"' AND RVTYP='DPF' AND VTYP='DPF'"

CN.Execute SQL, L

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & LBLCHLN & "',RVBDT= '" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "',RDBC = '" & VTCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
   "' AND PKG_STCOD='" & Trim(lstBox.ListItems(INDEX).SubItems(8)) & _
   "' AND (VTYP='PPF' OR VTYP='OPN') AND VBNO='" & lstBox.ListItems(INDEX).Text & "'"
   
   CN.Execute SQL, L
 End If
Next INDEX

 Call DAILYSTATUS("DPF", GetCode("ACCMST", txtpcod, "NAME", "CODE"), VTCD, Val(txtNTWT), LBLCHLN, 0, cUName, "M", Now, TXTVBDT)

CN.CommitTrans

End If

'PLY UPDATION COMMON FOR BOTH SAVE AND EDIT
If TOPBOTTOM > 0 Then
Dim NOOFPLY As Double
Dim i As Long, J As Long, K As Long
Dim RSTMP As New ADODB.Recordset
Set RSTMP = New ADODB.Recordset
If RSTMP.State = 1 Then RSTMP.Close
RSTMP.Open "SELECT * FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='" & VTCD & "' AND VTYP='DPF' AND CHLN='" & LBLCHLN & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic

If Not RSTMP.EOF Then
i = 0
  For i = 12 To lstBox.ColumnHeaders.COUNT
    J = 0
    For J = 0 To RSTMP.Fields.COUNT - 1
      If Trim(RSTMP.Fields(J).NAME) = Trim(lstBox.ColumnHeaders(i).Text) Then
            NOOFPLY = 0
         '---------------------------------------------------------------------
            For K = 1 To lstBox.ListItems.COUNT
               If lstBox.ListItems(K).Checked = True And lstBox.ListItems(K).SubItems(9) = "Y" Then
                  NOOFPLY = NOOFPLY + Val(lstBox.ListItems(K).SubItems(i - 1))
               End If
            Next K
         '---------------------------------------------------------------------
         
         RSTMP.Fields(J).Value = NOOFPLY
      End If
    Next J
  Next i
  
  RSTMP!TOPPLY = TOPBOTTOM
  RSTMP!BOTTOMPLY = TOPBOTTOM
  RSTMP.Update
  End If

If RSTMP.State = 1 Then RSTMP.Close
End If
'-------------------------------------------------

If SAVEFLAG Then
   MsgBox "Your Challan No. is : " & LBLCHLN
Else
   MsgBox "Challan No.: " & LBLCHLN & " Successfully Edited."
End If
   
 If IsOnlineChallanPrintReq Then 'IS PRINTING ONLINE CHALLAN REQUIRED ???
     
    OnlineChallanNum = LBLCHLN
    
    LOAD frmRPT_DelChallanPrint
    frmRPT_DelChallanPrint.Hide
        
    frmRPT_DelChallanPrint.cboStatus.ListIndex = 0
    frmRPT_DelChallanPrint.txtUNIT = UntNm
    frmRPT_DelChallanPrint.txtUNIT.Tag = UNCD
    frmRPT_DelChallanPrint.txtDVCD = LBLDIV.Caption
    frmRPT_DelChallanPrint.txtDVCD.Tag = DIVCODE
        
    frmRPT_DelChallanPrint.cmbDispatchType.AddItem cmbPackingType.Text
    frmRPT_DelChallanPrint.cmbDispatchType.Text = cmbPackingType
        
    frmRPT_DelChallanPrint.lstCHLN_GotFocus
    frmRPT_DelChallanPrint.opPlain.Value = True
    frmRPT_DelChallanPrint.cmdpreview_Click
    
 End If


txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
LBLS = 0:  LBLZ = 0: LBL0 = 0
lstBox.ListItems.Clear
Call CLEARDATA
Call cmdCancel_Click

TXTVBDT = Now

Exit Sub
LAST:
MsgBox ERR.Description
Exit Sub
End Sub

Private Sub cmdSavePrint_Click()
Call cmdSave_Click
End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or LBLDIV = Empty Then
  Unload Me
End If

Me.BackColor = RGB(RED, GREEN, BLUE)

If Not INIT Then
   INIT = True
   btn_sts (True)
   Call SetLRDetail
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Dim TWSTRS As New ADODB.Recordset
Dim TWSTREQ As String
SAVEFLAG = True
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
TXTADDRESS.FontBold = False
Me.Left = 50: Me.KeyPreview = True
SAVEFLAG = True
EDITFLAG = True

  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  DIVCODE = Empty
  If DIVCODE = Empty Then
    LBLDIV.Caption = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
  
  If SetIsShadeReq(DIVCODE) = "Y" Then
      GRD_SHD_REQ = True
  End If
   
  
  Call SetLabel
  Call SetPackingType
  LRDT = Now
  TXTVBDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  
  LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)

'----------------------------------------
   If TWSTRS.State = 1 Then TWSTRS.Close
   TWSTRS.Open "SELECT * FROM DIVMST WHERE COMP = '" & compPth & "' AND  UNIT = '" & UNCD & "' AND CODE = '" & DIVCODE & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TWSTRS.EOF Then
      TWSTREQ = Trim(TWSTRS!TWSTREQ & "")
   End If
   
  If TWSTREQ = "Y" Then
      ShapeSZO.Visible = True: Label1.Visible = True
      LabelS.Visible = True:  LabelZ.Visible = True:  LabelO.Visible = True
      LBLS.Visible = True:  LBLZ.Visible = True:  LBL0.Visible = True
      LBLSWGT.Visible = True: LBLZWGT.Visible = True: LBLOWGT.Visible = True
  Else
      ShapeSZO.Visible = False: Label1.Visible = False
      LabelS.Visible = False: LabelZ.Visible = False: LabelO.Visible = False
      LBLS.Visible = False: LBLZ.Visible = False: LBL0.Visible = False
      LBLSWGT.Visible = False: LBLZWGT.Visible = False: LBLOWGT.Visible = False
  End If
 '---------------------------------------
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset
Dim COUNT As Long: COUNT = 11
If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT * FROM PLYMST WHERE RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
Do While Not GETRS.EOF
    COUNT = COUNT + 1
    lstBox.ColumnHeaders.ADD COUNT, , Trim(GETRS!NAME & ""), 0, 0, 0
GETRS.MoveNext
Loop
GETRS.Close

INIT = False

End Sub

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT DISTINCT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND NAME NOT LIKE '%CAPTIVE%' AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "'  AND NAME<>''"

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic

If Not PKTYPRS.EOF Then VTCD = Trim(PKTYPRS!CODE)

Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop
If cmbPackingType.ListCount > 0 Then cmbPackingType.ListIndex = 0
End Sub

Private Sub TimerBillNo1_Timer()
Static ctr As Integer
If ctr Mod 45 = 0 And ctr <= 45 Then
   LBLHEADING1.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLHEADING2.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): BORDER1.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE): BORDER2.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE): BORDER3.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
   LBLDIV.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLHEAD.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLCHLN.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
   LBLHEADING1.ForeColor = vbRed: LBLHEADING2.ForeColor = vbRed: BORDER1.BorderColor = vbRed: BORDER2.BorderColor = vbRed: BORDER3.BorderColor = vbRed
   LBLDIV.ForeColor = vbRed: LBLHEAD.ForeColor = vbRed: LBLCHLN.ForeColor = vbRed
ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
   LBLHEADING1.ForeColor = vbBlue: LBLHEADING2.ForeColor = vbBlue: BORDER1.BorderColor = vbBlue: BORDER2.BorderColor = vbBlue: BORDER3.BorderColor = vbBlue
   LBLDIV.ForeColor = vbBlue: LBLHEAD.ForeColor = vbBlue: LBLCHLN.ForeColor = vbBlue
   ctr = 0
End If
ctr = ctr + 15
End Sub

Private Sub LRDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub lstBox_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CNT As Long

Select Case KeyCode

Case vbKeyF6

   For CNT = 1 To lstBox.ListItems.COUNT
    If lstBox.ListItems(CNT).Checked = False Then
       lstBox.ListItems(CNT).Checked = True
       txtCTRN.Caption = Val(txtCTRN.Caption) + 1
       txtNTWT.Caption = nstr(Val(txtNTWT.Caption) + Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
       txtNTWT.Caption = Trim(txtNTWT.Caption)
       txtCops.Caption = Val(txtCops.Caption) + Val(lstBox.ListItems(CNT).ListSubItems(1))
       Select Case Trim(lstBox.ListItems(CNT).ListSubItems(3))
       Case "S"
           LBLS = Val(LBLS) + 1
           LBLSWGT = nstr(Val(LBLSWGT) + Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
       Case "Z"
           LBLZ = Val(LBLZ) + 1
           LBLZWGT = nstr(Val(LBLZWGT) + Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
       Case "0"
           LBL0 = Val(LBL0) + 1
           LBLOWGT = nstr(Val(LBLOWGT) + Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
       End Select
        
    End If
   Next
   
Case vbKeyF7
   
   For CNT = 1 To lstBox.ListItems.COUNT
    If lstBox.ListItems(CNT).Checked = True Then
       lstBox.ListItems(CNT).Checked = False
        
        txtCTRN.Caption = Val(txtCTRN.Caption) - 1
        txtNTWT.Caption = nstr(Val(txtNTWT.Caption) - Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
        txtNTWT.Caption = Trim(txtNTWT.Caption)
        txtCops.Caption = Val(txtCops.Caption) - Val(lstBox.ListItems(CNT).ListSubItems(1))
        
        Select Case Trim(lstBox.ListItems(CNT).ListSubItems(3))
        Case "S"
           LBLS = Val(LBLS) - 1
           LBLSWGT = nstr(Val(LBLSWGT) + Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
        Case "Z"
           LBLZ = Val(LBLZ) - 1
           LBLZWGT = nstr(Val(LBLZWGT) + Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
        Case "0"
           LBL0 = Val(LBL0) - 1
           LBLOWGT = nstr(Val(LBLOWGT) + Val(lstBox.ListItems(CNT).ListSubItems(2)), 10, 3)
        End Select
        
    End If
   Next
   
End Select

End Sub

Private Sub TXTADDRESS_Change()
    If SAVEFLAG And TXTADDRESS <> Empty Then Call SetLastConsigneeLot
End Sub

Private Sub txtCONSINEE_Change()
  TXTADDRESS = Empty
End Sub

Private Sub txtCONSINEE_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtCONSINEE = Empty
  ElseIf KeyCode = vbKeyF2 Or txtCONSINEE = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False: Key = Empty
     txtCONSINEE = SearchList1("Select DISTINCT CODE,NAME From PADDMST", 0, Empty, "Select Consinee Name ")
     txtCONSINEE.Tag = Key
  End If
  
 Me.KeyPreview = True

End Sub

Private Sub TXTGRAD_Change()
If EDITFLAG Then
 lstBox.ListItems.Clear
 Call GenerateBoxList
End If
End Sub

Private Sub TXTGRNNO_GotFocus()
   TXTGRNNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTGRNNO_KeyDown(KeyCode As Integer, Shift As Integer)
If txtpcod = Empty Then
   Exit Sub
End If

Me.KeyPreview = False
Dim PTYCOD As String
Key = Empty
PTYCOD = GetCode("ACCMST", txtpcod, "NAME", "CODE")

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTGRNNO = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTGRNNO = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
     TXTGRNNO = SearchList("SELECT DISTINCT VBNO,VBNO FROM JOBGRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                           "' AND VTYP='IVR' AND DBCD='000002' AND PCOD='" & PTYCOD & _
                           "' AND CLRSTATUS='N' AND RECSTAT<>'D' ")
End If
Me.KeyPreview = True
End Sub

Private Sub TXTGRNNO_LostFocus()
  TXTGRNNO.BackColor = vbWhite
End Sub

Private Sub TXTITEM_Change()
If EDITFLAG Then
 lstBox.ListItems.Clear
 Call GenerateBoxList
End If
End Sub

Private Sub TXTLRNO_GotFocus()
  TXTLRNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTLRNO_LostFocus()
 TXTLRNO.BackColor = vbWhite
End Sub

Private Sub txtLTNO_Change()
If EDITFLAG Then
   txtITEM = FindItem
   lstBox.ListItems.Clear
   Call GenerateBoxList
End If
End Sub

Private Sub txtPCOD_Change()

'If SAVEFLAG And txtPCOD <> Empty Then Call SetLastPartyLot

If EDITFLAG And CHLNTYP = "JOB CHALLAN" Then
   lstBox.ListItems.Clear
   Call GenerateBoxList
End If

If SAVEFLAG Then TXTGRNNO = Empty
End Sub

Private Sub txtPCOD_GotFocus()
  txtpcod.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
 If txtpcod = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For A/C Party Master Help", "", txtpcod.Left, txtpcod.Top + txtpcod.Height + 100
  Else
      ToolTip Me, "Press {F2} For A/C Party Master Help", "", txtpcod.Left, txtpcod.Top + txtpcod.Height + 100
  End If
  picToolTip.WIDTH = picToolTip.WIDTH + 8450
End Sub

Private Sub txtConsinee_GotFocus()
  txtCONSINEE.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  
  If txtCONSINEE = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Consinee Master Help", "", txtCONSINEE.Left + 14500, txtCONSINEE.Top + txtCONSINEE.Height + 100
  Else
      ToolTip Me, "Press {F2} For Consinee Master Help", "", txtCONSINEE.Left + 6500, txtCONSINEE.Top + txtCONSINEE.Height + 100
  End If
  picToolTip.WIDTH = picToolTip.WIDTH + 8450
End Sub

Private Sub TXTADDRESS_GotFocus()
  TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
   SendKeys "{HOME}+{END}"
  
  If TXTADDRESS = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Address Master Help", "", TXTADDRESS.Left + 9900, TXTADDRESS.Top + TXTADDRESS.Height + 100
  Else
      ToolTip Me, "Press {F2} For Address Master Help", "", TXTADDRESS.Left + 9900, TXTADDRESS.Top + TXTADDRESS.Height + 100
  End If
  picToolTip.WIDTH = picToolTip.WIDTH + 8450
  
  Dim TEMPRS As New ADODB.Recordset
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT *FROM PADDMST WHERE NAME='" & txtCONSINEE & "' AND ADDR='" & Trim(TXTADDRESS) & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
        ADD_SRNO = TEMPRS!SRNO
   End If
   
   If SAVEFLAG And TXTADDRESS <> Empty Then Call SetLastConsigneeLot
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   TXTADDRESS.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Or (TXTADDRESS = Empty And KeyCode = vbKeyReturn) Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & txtCONSINEE & "'", 0, Empty, "Select Consignee Address from List")
   End If
   
   Dim TEMPRS As New ADODB.Recordset
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT *FROM PADDMST WHERE CODE='" & txtCONSINEE.Tag & "' AND ADDR='" & Trim(TXTADDRESS) & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
        ADD_SRNO = TEMPRS!SRNO
   End If
   
End Sub

Private Sub txtItem_GotFocus()
    txtITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTGRAD_GotFocus()
  TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If (TXTGRAD = Empty And KeyCode = vbKeyReturn) Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = True
    TXTGRAD = SearchList1("SELECT DISTINCT GRAD AS GRD,GRAD FROM GRDMST", 0, TXTGRAD, "SELECT MAIN GRAD FROM LIST")
      If key_PressNew = True Then
          M_DESC = ""
          TXTGRAD = Empty
          FRM_GRDMST.Show
      End If
  End If
End Sub

Private Sub TXTGRAD_LostFocus()
  TXTGRAD.BackColor = vbWhite
  picToolTip.Visible = False
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      txtpcod = Empty
  ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtpcod = Empty) Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtpcod = SearchList1("Select TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party ")
  End If
  
  Me.KeyPreview = True

End Sub

Private Sub txtRate_GotFocus()
txtRate.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtRate_LostFocus()
txtRate.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_Change()
If EDITFLAG Then
  lstBox.ListItems.Clear
  Call GenerateBoxList
End If
End Sub

Private Sub TXTSUBGRD_GotFocus()
  TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub


Private Sub txtltno_GotFocus()
  txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNo = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtLTNo = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   txtLTNo = SearchList("SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND ACTIVE='Y' ")
End If
   txtITEM = FindItem
Me.KeyPreview = True
End Sub


Private Sub TXTRMRK_GotFocus(): TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub

Private Sub txtPCOD_LostFocus()
 txtpcod.BackColor = vbWhite
 picToolTip.Visible = False
 
    If SAVEFLAG Then
     Dim GETRS As ADODB.Recordset
     Set GETRS = New ADODB.Recordset
  
     If GETRS.State = 1 Then GETRS.Close
     GETRS.Open "SELECT RCOD FROM ACCMST WHERE NAME='" & txtpcod & "' ", CN, adOpenDynamic, adLockOptimistic
     If Not GETRS.EOF Then
        txtCONSINEE = GetCode("PADDMST", GETRS!RCOD & "", "CODE", "NAME")
        TXTADDRESS = GetCode("PADDMST", GETRS!RCOD & "", "CODE", "ADDR")
     End If
  End If
  
End Sub
Private Sub txtConsinee_LostFocus(): txtCONSINEE.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub TXTADDRESS_LostFocus(): TXTADDRESS.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub txtItem_LostFocus(): txtITEM.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
Dim SQL As String
If TXTGRAD = Empty Then TXTGRAD.Enabled = True: TXTGRAD.SetFocus: Exit Sub

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
ElseIf KeyCode = vbKeyF2 Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT RDIFF,NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND GRAD='" & GetCode("GRDMST", TXTGRAD, "GRAD", "CODE") & "'"
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
End If
Me.KeyPreview = True
End Sub

Private Sub TXTSUBGRD_LostFocus(): TXTSUBGRD.BackColor = vbWhite:  picToolTip.Visible = False: End Sub
Private Sub txtltno_LostFocus(): txtLTNo.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub TXTRMRK_LostFocus(): TXTRMRK.BackColor = vbWhite: End Sub

Private Sub CLEARDATA()
 txtpcod = Empty: txtCONSINEE = Empty: TXTADDRESS = Empty: txtITEM = Empty: txtLTNo = Empty: TXTGRAD = Empty: TXTSUBGRD = Empty
 TXTRMRK = Empty: txtRate = Empty
 txtCTRN = Empty: txtCops = Empty: txtNTWT = Empty: LBLS = Empty: LBLZ = Empty: LBL0 = Empty
 txtRMNCTRN = Empty: txtRMNCOPs = Empty: txtRMNNTWT = Empty
 txtTTLCTRN = Empty: txtTTLCOPs = Empty: txtTTLNTWT = Empty
 'TXTLRNO = Empty: TXTVHCL = Empty: txtTransport = Empty
End Sub

Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtRate, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub lstBox_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Msg ("Press F8 For Select All F9 For De-Select All")
Call GetDetailedSelection
If Item.INDEX < lstBox.ListItems.COUNT Then lstBox.ListItems.Item(Item.INDEX + 1).Selected = True: lstBox.ListItems(Item.INDEX + 1).EnsureVisible
Exit Sub

Dim i As Integer, J As Integer, ctr As Integer

    If Item.Checked = True Then
        txtCTRN.Caption = Val(txtCTRN.Caption) + 1
        txtNTWT.Caption = nstr(Val(txtNTWT.Caption) + Val(Item.SubItems(2)), 10, 3)
        txtNTWT.Caption = Trim(txtNTWT.Caption)
        txtCops.Caption = Val(txtCops.Caption) + Val(Item.SubItems(1))
        Select Case Trim(lstBox.SelectedItem.SubItems(3))
        Case "S"
           LBLS = Val(LBLS) + 1
           LBLSWGT = nstr(Val(LBLSWGT) + Val(Item.SubItems(2)), 10, 3)
        Case "Z"
           LBLZ = Val(LBLZ) + 1
           LBLZWGT = nstr(Val(LBLZWGT) + Val(Item.SubItems(2)), 10, 3)
        Case "0"
           LBL0 = Val(LBL0) + 1
           LBLOWGT = nstr(Val(LBLOWGT) + Val(Item.SubItems(2)), 10, 3)
        End Select
    Else
        txtCTRN.Caption = Val(txtCTRN.Caption) - 1
        txtNTWT.Caption = nstr(Val(txtNTWT.Caption) - Val(Item.SubItems(2)), 10, 3)
        txtNTWT.Caption = Trim(txtNTWT.Caption)
        txtCops.Caption = Val(txtCops.Caption) - Val(Item.SubItems(1))
        'lblBCartn.Caption = Val(lblBCartn.Caption) + 1 'lblBCops.Caption = Val(lblBCops.Caption) + Val(Item.SubItems(1)) 'lblBNWGT.Caption = Val(lblBNWGT.Caption) + Val(Item.SubItems(2))
        Select Case Trim(lstBox.SelectedItem.SubItems(3))
        Case "S"
           LBLS = Val(LBLS) - 1
           LBLSWGT = nstr(Val(LBLSWGT) + Val(Item.SubItems(2)), 10, 3)
        Case "Z"
           LBLZ = Val(LBLZ) - 1
           LBLZWGT = nstr(Val(LBLZWGT) + Val(Item.SubItems(2)), 10, 3)
        Case "0"
           LBL0 = Val(LBL0) - 1
           LBLOWGT = nstr(Val(LBLOWGT) + Val(Item.SubItems(2)), 10, 3)
        End Select
    End If
        
    
    If Item.INDEX < lstBox.ListItems.COUNT Then lstBox.ListItems.Item(Item.INDEX + 1).Selected = True: lstBox.ListItems(Item.INDEX + 1).EnsureVisible
End Sub

Public Sub btn_sts(Yes As Boolean)
 cmdSave.Enabled = Not Yes: cmdCancel.Enabled = Not Yes: cmdAdd.Enabled = Yes: cmdEdit.Enabled = Yes
 cmdDelete.Enabled = Yes
 txtpcod.Enabled = Not Yes: txtCONSINEE.Enabled = Not Yes: TXTADDRESS.Enabled = Not Yes: txtITEM.Enabled = Not Yes: txtLTNo.Enabled = Not Yes: TXTGRAD.Enabled = Not Yes: TXTSUBGRD.Enabled = Not Yes
 txtRate.Enabled = Not Yes: TXTRMRK.Enabled = Not Yes
End Sub


Private Function FindItem() As String
Dim FICD As String
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset


If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   FICD = Trim(FINDRS!FICD & "")
Else
   FICD = Empty
   Exit Function
End If
FINDRS.Close

If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   FindItem = Trim(FINDRS!NAME & "")
Else
   FindItem = Empty
   Exit Function
End If
FINDRS.Close

End Function

Private Sub GenerateBoxList()
If txtITEM = Empty Or TXTGRAD = Empty Or txtLTNo = Empty Then Exit Sub 'Or TXTSUBGRD = Empty
If CHLNTYP = "JOB CHALLAN" And txtpcod = Empty Then txtpcod.Enabled = True: txtpcod.SetFocus: Exit Sub

If GRD_SHD_REQ Then
   If TXTSUBGRD = Empty Then Exit Sub
End If

SGRD = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")

Dim SQL As String
Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset

SUBGRD = Empty

If TXTSUBGRD <> Empty Then
    If IsTwistRequired Then
       SUBGRD = Trim(TXTSUBGRD)
    Else
       Dim TYP As String
       TYP = LabelType(DIVCODE, UNCD)
       If TYP = "SG" Then
          If RSDATA.State = 1 Then RSDATA.Close
          RSDATA.Open "SELECT SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
          "' AND DVCD = '" & DIVCODE & "' AND GRAD = '" & SGRD & "' AND NAME = '" & TXTSUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
          If Not RSDATA.EOF Then
             SUBGRD = Trim(RSDATA!SUBGRD & "")
          End If
          RSDATA.Close
       ElseIf TYP = "GS" Then
          
          If RSDATA.State = 1 Then RSDATA.Close
          RSDATA.Open "SELECT SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
          "' AND NAME = '" & TXTSUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
          If Not RSDATA.EOF Then
             SUBGRD = Trim(RSDATA!SUBGRD & "")
          End If
          RSDATA.Close
       End If
    End If
End If

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND NAME = '" & txtITEM & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SITEM = Trim(RSDATA!CODE & "")
End If
RSDATA.Close

'INITIAL SET TOTAL BOX COPES
 txtTTLCOPs = 0
 txtTTLCTRN = 0
 txtTTLNTWT = 0
 txtRMNCOPs = 0
 txtRMNCTRN = 0
 txtRMNNTWT = 0
'===============================================

SQL = "SELECT * FROM BOXREGISTER WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & DIVCODE & _
"'AND BOXREGISTER.LOTNO ='" & txtLTNo & "' AND BOXREGISTER.ICOD = '" & SITEM & _
"' AND BOXREGISTER.GRAD ='" & SGRD & "' AND BOXREGISTER.RECSTAT<>'D' "

'AND (VTYP='PPF' OR VTYP='OPN')
SQL = SQL & " AND ( (VTYP IN ('PPF','OPN')) OR (VTYP='DPF' AND RDBC='" & VTCD & "' AND RVBNO='" & LBLCHLN & "') ) "

'if sugrade exist
If SUBGRD <> Empty Then
   SQL = SQL & " AND BOXREGISTER.SUBGRD = '" & SUBGRD & "' "
End If

'DO DATE ARE LESS THEN OR EQUAL TO CHALLAN DATE
SQL = SQL & " AND BOXREGISTER.VBDT <= '" & Format(TXTVBDT.Value, "MM/DD/YYYY") & "' "
'---------------------------------------------------

If InStr(1, UCase(cmbPackingType.Text), "JOB CHALLAN") <> 0 Then
   SQL = SQL & "AND DBCD='000005' ORDER BY VBNO"
ElseIf InStr(1, UCase(cmbPackingType.Text), "CAPTIVE") <> 0 Then
   SQL = SQL & "AND DBCD='000001' ORDER BY VBNO"
ElseIf InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then
   SQL = SQL & "AND DBCD='000002' ORDER BY VBNO"
Else
   SQL = SQL & "AND DBCD IN ('000003','000004') ORDER BY VBNO"
End If

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   MsgBox "Boxes are not available for this criteria."
   TXTGRAD.Enabled = True: TXTGRAD.SetFocus
   Exit Sub
End If
  txtCTRN.Caption = 0
  txtNTWT.Caption = 0
  txtNTWT.Caption = 0
  txtCops.Caption = 0
  LBLS = 0
  LBLZ = 0
  LBL0 = 0
  LBLSWGT = 0
  LBLZWGT = 0
  LBLOWGT = 0
  
  lstBox.ListItems.Clear
  Dim Item
  Do While Not RSDATA.EOF
   Set Item = lstBox.ListItems.ADD
   Item.Text = RSDATA!VBNO
   Item.SubItems(1) = RSDATA!COPS
   Item.SubItems(2) = nstr(RSDATA!NTWGT, 9, 3)
   Item.SubItems(2) = Trim(Item.SubItems(2))
   If Trim(RSDATA!SUBGRD) = "S" Or Trim(RSDATA!SUBGRD) = "Z" Or Trim(RSDATA!SUBGRD) = "O" Then
     Item.SubItems(3) = Trim(RSDATA!SUBGRD)
     If lstBox.SelectedItem.ListSubItems.COUNT = 2 Then lstBox.ColumnHeaders(4).Text = "Twist"
   Else
     Item.SubItems(3) = FindSubGradeName(Trim(RSDATA!SUBGRD & ""))
     If lstBox.ListItems.COUNT = 1 Then
        If Not GRD_SHD_REQ Then
           lstBox.ColumnHeaders(4).Text = "SubGrade"
        Else
           lstBox.ColumnHeaders(4).Text = "Shade"
           lstBox.ColumnHeaders(4).WIDTH = 0
        End If
     End If
   End If
   
   Item.SubItems(4) = nstr(RSDATA!GRSWGT, 9, 3)
   Item.SubItems(4) = Trim(Item.SubItems(4) & "")
   Item.SubItems(5) = nstr(RSDATA!TRWGT, 9, 3)
   Item.SubItems(5) = Trim(Item.SubItems(5) & "")
   Item.SubItems(6) = Format(RSDATA!VBDT, "DD/MM/YYYY")
   Item.SubItems(7) = Trim(RSDATA!RMRK & "")
   Item.SubItems(8) = Trim(RSDATA!PKG_STCOD & "")
   Item.SubItems(9) = Trim(RSDATA!ISRETURNABLE & "")
   Item.SubItems(10) = Trim(RSDATA!Top & "")
   
   Dim i As Double, J As Double
   i = 0
   For i = 12 To lstBox.ColumnHeaders.COUNT
      J = 0
      For J = 0 To RSDATA.Fields.COUNT - 1
        If Trim(RSDATA.Fields(J).NAME) = Trim(lstBox.ColumnHeaders(i).Text) Then
            Item.SubItems(i - 1) = Val(RSDATA.Fields(J).Value)
        End If
      Next
   Next
   
     
     txtTTLCOPs = Val(txtTTLCOPs) + Val(RSDATA!COPS)
     txtTTLCTRN = Val(txtTTLCTRN) + 1
     txtTTLNTWT = Val(txtTTLNTWT) + Val(RSDATA!NTWGT)
           
     RSDATA.MoveNext
  Loop
  RSDATA.Close
  
  TXTVBDT.Enabled = False
  
End Sub

Private Sub SetGlobal()
Dim INDEX As Long

TTLQTY = 0: TTLBOXES = 0: TTLCOPS = 0
TTLGRSWGT = 0: TTLTAREWGT = 0
Dim DBCDRS As ADODB.Recordset
Set DBCDRS = New ADODB.Recordset
If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND NAME = '" & cmbPackingType.Text & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   VTCD = Trim(DBCDRS!CODE & "")
Else
   VTCD = Empty
End If
DBCDRS.Close

For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then
     TTLCOPS = TTLCOPS + Val(lstBox.ListItems(INDEX).SubItems(1))
     TTLBOXES = TTLBOXES + 1
     TTLQTY = TTLQTY + Val(lstBox.ListItems(INDEX).SubItems(2))
     TTLGRSWGT = TTLGRSWGT + Val(lstBox.ListItems(INDEX).SubItems(4))
     TTLTAREWGT = TTLTAREWGT + Val(lstBox.ListItems(INDEX).SubItems(5))
  End If
Next

SPARTY = GetCode("ACCMST", txtpcod, "NAME", "CODE")
M_BRCD = GetCode("ACCMST", txtpcod, "NAME", "BRCD")

SGRD = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")

Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT CODE,SRNO FROM PADDMST WHERE NAME='" & txtCONSINEE & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
  SCONSINEE = RSDATA!CODE & ""
  SADD = RSDATA!SRNO & ""
Else
  SCONSINEE = Empty
  SADD = Empty
End If
RSDATA.Close

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND NAME = '" & txtITEM & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SITEM = Trim(RSDATA!CODE & "")
End If
RSDATA.Close

If GRD_SHD_REQ Then
   QRY = "SELECT SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
         "' AND NAME = '" & TXTSUBGRD & "'"
Else
   QRY = "SELECT SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
         "' AND DVCD = '" & DIVCODE & "' AND GRAD = " & Val(SGRD) & " AND NAME = '" & TXTSUBGRD & "'"
End If

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open QRY, CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SUBGRD = Trim(RSDATA!SUBGRD & "")
End If
RSDATA.Close

'TRANSPORT CODE
If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT * FROM TRANSPORTMST WHERE NAME ='" & txtTransport & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   M_TRCD = Trim(RSDATA!CODE & "")
Else
   M_TRCD = Empty
End If
RSDATA.Close

'VEHICLE CODE
If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT * FROM VHCLMST WHERE NAME ='" & TXTVHCL & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   M_VHCD = Trim(RSDATA!CODE & "")
Else
   M_VHCD = Empty
End If
RSDATA.Close

End Sub

Private Function CHKSAVEDATA() As Boolean
  CHKSAVEDATA = True
  
  Dim CHKRS As ADODB.Recordset
  Set CHKRS = New ADODB.Recordset
    
  If SAVEFLAG And (Trim(txtpcod) = Empty Or Trim(txtCONSINEE) = Empty Or Trim(TXTADDRESS) = Empty Or Trim(txtLTNo) = Empty Or Trim(txtITEM) = Empty Or Trim(TXTGRAD) = Empty Or Trim(TXTSUBGRD) = Empty Or Trim(txtRate) = Empty) Then
     If txtpcod = Empty Then txtpcod.Enabled = True: txtpcod.SetFocus: CHKSAVEDATA = False: Exit Function
     If txtCONSINEE = Empty Then txtCONSINEE.Enabled = True: txtCONSINEE.SetFocus: CHKSAVEDATA = False: Exit Function
     If TXTADDRESS = Empty Then TXTADDRESS.Enabled = True: TXTADDRESS.SetFocus: CHKSAVEDATA = False: Exit Function
     If txtLTNo = Empty Then txtLTNo.Enabled = True: txtLTNo.SetFocus: CHKSAVEDATA = False: Exit Function
     If txtITEM = Empty Then txtITEM.Enabled = True: txtITEM.SetFocus: CHKSAVEDATA = False: Exit Function
     If TXTGRAD = Empty Then TXTGRAD.Enabled = True: TXTGRAD.SetFocus: CHKSAVEDATA = False: Exit Function
     'If TXTSUBGRD = Empty And TXTSUBGRD.Enabled And TXTSUBGRD.Visible Then TXTSUBGRD.SetFocus: CHKSAVEDATA = False: Exit Function
     If txtRate = Empty Then txtRate.Enabled = True: txtRate.SetFocus: CHKSAVEDATA = False: Exit Function
  End If
  
  
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE FROM ACCMST WHERE NAME='" & txtpcod & "'", CN, adOpenDynamic, adLockOptimistic
  If CHKRS.EOF Then
     MsgBox "Party Not Properly Defined"
     CHKSAVEDATA = False
     If txtpcod.Enabled Then txtpcod.SetFocus
     Exit Function
  Else
     SPARTY = Trim(CHKRS!CODE & "")
  End If
  
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE,SRNO FROM PADDMST WHERE NAME='" & txtCONSINEE & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
  If CHKRS.EOF Then
     MsgBox "Consignee Not Properly Defined"
     CHKSAVEDATA = False
     If txtCONSINEE.Enabled Then txtCONSINEE.SetFocus
     Exit Function
  Else
     SCONSINEE = CHKRS!CODE & ""
     SADD = CHKRS!SRNO & ""
  End If
  
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
  "' AND DVCD = '" & DIVCODE & "' AND NAME = '" & txtITEM & "'", CN, adOpenDynamic, adLockOptimistic
  If CHKRS.EOF Then
     MsgBox "Finish Item Not Properly Defined"
     CHKSAVEDATA = False
     If txtITEM.Enabled Then txtITEM.SetFocus
     Exit Function
  Else
     SITEM = Trim(CHKRS!CODE & "")
  End If
  CHKRS.Close

  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE FROM GRDMST WHERE GRAD='" & TXTGRAD & "'", CN, adOpenDynamic, adLockOptimistic
  If CHKRS.EOF Then
     MsgBox "Grade Not Properly Defined"
     CHKSAVEDATA = False
     If TXTGRAD.Enabled Then TXTGRAD.SetFocus
     Exit Function
  Else
     SGRD = Trim(CHKRS!CODE & "")
  End If
        
  If SAVEFLAG = True And InStr(1, UCase(cmbPackingType.Text), "JOB CHALLAN") <> 0 Then
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT CLRSTATUS FROM JOBGRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='IVR' " & _
             "AND DBCD='000002' AND RECSTAT<>'D' AND VBNO='" & Trim(TXTGRNNO) & "'", CN, adOpenDynamic, adLockOptimistic
     If Not RS.EOF Then
        If Trim(RS!CLRSTATUS & "") = "Y" Then
           MsgBox "GRN IS CLEARED", vbCritical
           TXTGRNNO.Enabled = True
           CHKSAVEDATA = False
           Exit Function
        End If
     Else
        TXTGRNNO.Enabled = True
        CHKSAVEDATA = False
        Exit Function
     End If
  End If
    
    Dim M_BOXES As String, X As Long
    For X = 1 To lstBox.ListItems.COUNT
      If lstBox.ListItems(X).Checked = True Then
         If M_BOXES <> Empty Then M_BOXES = M_BOXES & ","
         M_BOXES = M_BOXES & "'" & Trim(lstBox.ListItems(X).Text) & "'"
      End If
    Next
    
    If SAVEFLAG Then
       LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
    End If
    
    If M_BOXES <> Empty Then 'NEW
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT RVBNO,VBNO FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='DPF' " & _
                "AND RECSTAT<>'D' AND DVCD='" & DIVCODE & "' AND VBNO NOT IN (SELECT VBNO FROM BOXREGISTER WHERE COMP='" & compPth & _
                "' AND UNIT='" & UNCD & "' AND VTYP='DPF' AND RECSTAT<>'D' AND RDBC='" & VTCD & _
                "' AND RVBNO='" & LBLCHLN & "') AND VBNO IN (" & M_BOXES & ")", CN, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
           MsgBox "Box No. " & Trim(RS!VBNO & "") & " Already Dispatched In Challan " & Trim(RS!RVBNO & "") & ".Click Ok To Refresh List", vbCritical, "Multi Location Dispatch"
           lstBox.ListItems.Clear
           Call GenerateBoxList
           If lstBox.Enabled Then lstBox.SetFocus
           CHKSAVEDATA = False
           Exit Function
        End If
    End If
End Function

Private Function IsSaleExist() As Boolean
'default
IsSaleExist = False
'-----------------------------------
Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
Dim SQL As String

  'CODE TO CHECK SALE BILL EXIST
  SQL = "SELECT TOP 1 VBNO FROM SPTRAN "
  SQL = SQL & "WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND DVCD ='" & DIVCODE & "' AND DBCD ='" & VTCD & _
  "'  AND VBNO ='" & chln & "' AND VTYP='DPF' AND RECSTAT='A' AND SVBN IS NOT NULL"
   
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not CHKRS.EOF Then
     IsSaleExist = True
  End If
  '---------------------------------
End Function

Private Sub SetLabel()
   Dim TYP As String
   TYP = LabelType(DIVCODE, UNCD)
   
   If TYP = "SD" Then
     LBLCFG2.Visible = False
     TXTSUBGRD.Visible = False
     LBLCFG.Caption = LabelDisplay(DIVCODE, UNCD)
   ElseIf TYP = "GD" Then
   
      If IsTwistRequired Then
         LBLCFG2.Visible = True
         LBLCFG2.Caption = "Twist"
         lstBox.ColumnHeaders(4).Text = "Twist"
         TXTSUBGRD.Visible = True
         LBLCFG.Caption = LabelDisplay(DIVCODE, UNCD)
         Exit Sub
      End If
      
      LBLCFG2.Visible = False
      TXTSUBGRD.Visible = False
      LBLCFG.Caption = LabelDisplay(DIVCODE, UNCD)
   ElseIf TYP = "GS" Then
     LBLCFG2.Visible = True
     TXTSUBGRD.Visible = True
     LBLCFG2.Caption = "Shade"
     lstBox.ColumnHeaders(4).Text = "Shade"
   Else
      lstBox.ColumnHeaders(4).Text = "SubGrade"
   End If
End Sub

Private Function IsTwistRequired() As Boolean
IsTwistRequired = False

Dim DISPRS As ADODB.Recordset
Set DISPRS = New ADODB.Recordset
        
If DISPRS.State = 1 Then DISPRS.Close
DISPRS.Open "SELECT TWSTREQ FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND CODE='" & DIVCODE & "' AND TWSTREQ='Y'", CN, adOpenDynamic, adLockOptimistic
      If Not DISPRS.EOF Then
         IsTwistRequired = True
         Exit Function
      End If
      DISPRS.Close
End Function

Private Function FindSubGradeName(SGCODE As String) As String
FindSubGradeName = "."
Dim QRY As String

Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRD_SHD_REQ Then
  QRY = "SELECT NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND SUBGRD='" & SGCODE & "'"
Else
  QRY = "SELECT NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND  DVCD='" & DIVCODE & "' AND GRAD='" & SGRD & "' AND SUBGRD='" & SGCODE & "'"
End If

If GRRS.State = 1 Then GRRS.Close
GRRS.Open QRY, CN, adOpenDynamic, adLockOptimistic

If Not GRRS.EOF Then
   FindSubGradeName = Trim(GRRS!NAME & "")
   GRRS.Close
   Exit Function
End If

End Function

Private Sub GetDetailedSelection()
Dim INDEX As Long
txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
LBLS = 0: LBLZ = 0: LBL0 = 0
LBLSWGT = 0: LBLZWGT = 0: LBLOWGT = 0

txtRMNCOPs = 0
txtRMNCTRN = 0
txtRMNNTWT = 0


For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then
     txtCTRN.Caption = Val(txtCTRN.Caption) + 1
     txtNTWT.Caption = nstr(Val(txtNTWT.Caption) + Val(lstBox.ListItems(INDEX).SubItems(2)), 10, 3)
     txtCops.Caption = Val(txtCops.Caption) + Val(lstBox.ListItems(INDEX).SubItems(1))
     
     Select Case Trim(lstBox.ListItems(INDEX).SubItems(3))
     Case "S"
        LBLS = Val(LBLS) + 1
        LBLSWGT = nstr(Val(LBLSWGT) + Val(lstBox.ListItems(INDEX).SubItems(2)), 10, 3)
     Case "Z"
        LBLZ = Val(LBLZ) + 1
        LBLZWGT = nstr(Val(LBLZWGT) + Val(lstBox.ListItems(INDEX).SubItems(2)), 10, 3)
     Case "O"
        LBL0 = Val(LBL0) + 1
        LBLOWGT = nstr(Val(LBLOWGT) + Val(lstBox.ListItems(INDEX).SubItems(2)), 10, 3)
     End Select
  End If
Next

txtRMNCOPs = Val(txtTTLCOPs) - Val(txtCops)
txtRMNCTRN = Val(txtTTLCTRN) - Val(txtCTRN)
txtRMNNTWT = Val(txtTTLNTWT) - Val(txtNTWT)

End Sub

Private Function RoundOffReq() As Boolean
RoundOffReq = False

Dim RFRS As ADODB.Recordset
Set RFRS = New ADODB.Recordset
        
If RFRS.State = 1 Then RFRS.Close
RFRS.Open "SELECT ITEMRO FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ITEMRO ='Y'", CN, adOpenDynamic, adLockOptimistic
      If Not RFRS.EOF Then
         RoundOffReq = True
         Exit Function
      End If
      RFRS.Close
End Function

Private Sub SetLastPartyLot()

    Dim LOTRS As ADODB.Recordset
    Set LOTRS = New ADODB.Recordset
            
    If LOTRS.State = 1 Then LOTRS.Close
    LOTRS.Open "SELECT LTNO FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND DVCD='" & DIVCODE & "' AND PCOD ='" & GetCode("ACCMST", txtpcod, "NAME", "CODE") & _
               "' AND VTYP='DPF' AND RECSTAT<>'D' AND DATE <= '" & Format(TXTVBDT.Value, "MM/DD/YYYY") & "' ORDER BY DATE DESC", CN, adOpenDynamic, adLockOptimistic
    If Not LOTRS.EOF Then
       txtLTNo = Trim(LOTRS!ltno & "")
    End If
    LOTRS.Close

End Sub

Private Sub SetLastConsigneeLot()

    Dim LOTRS As ADODB.Recordset
    Set LOTRS = New ADODB.Recordset
            
    If LOTRS.State = 1 Then LOTRS.Close
    LOTRS.Open "SELECT LTNO FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND DVCD='" & DIVCODE & "' AND DCOD ='" & GetCode("PADDMST", txtCONSINEE, "NAME", "CODE") & "' AND ADDRESS=" & ADD_SRNO & _
               " AND VTYP='DPF' AND RECSTAT<>'D' AND DATE <= '" & Format(TXTVBDT.Value, "MM/DD/YYYY") & "' ORDER BY DATE DESC", CN, adOpenDynamic, adLockOptimistic
    If Not LOTRS.EOF Then
       txtLTNo = Trim(LOTRS!ltno & "")
    End If
    LOTRS.Close

End Sub

Private Sub txtTransport_GotFocus(): txtTransport.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtTransport_LostFocus(): txtTransport.BackColor = vbWhite: End Sub

Private Sub TXTVHCL_GotFocus(): TXTVHCL.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTVHCL_LostFocus(): TXTVHCL.BackColor = vbWhite: End Sub

Private Sub txtVHCL_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
   
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTVHCL = Empty
  ElseIf KeyCode = vbKeyF2 Then
     M_DESC = Empty:   NEW_VISIBLE = True
     TXTVHCL = SearchList1("Select DISTINCT CODE,NAME From VHCLMST WHERE RECSTAT='A'", 0, Empty, "Select Vehicle From List. ")
     
     If key_PressNew Then
        LOAD frmVehicleMaster
     Else
        TXTVHCL.Tag = Key
     End If
     
     txtTransport.Tag = GetCode("VHCLMST", TXTVHCL.Tag, "CODE", "TRCD")
     txtTransport = GetCode("TRANSPORTMST", txtTransport.Tag, "CODE", "NAME")
  End If
  
 Me.KeyPreview = True
End Sub

Private Sub txtTransport_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(txtTransport) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtTransport.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM TRANSPORTMST WHERE RECSTAT<>'D'", 0, txtTransport.Text, "SELECT TRANSPORT FROM LIST")
        If key_PressNew = True Then
          M_DESC = ""
          txtTransport = Empty
          frmTransportMaster.Show
        Else
          txtTransport.Tag = Key
        End If
    ElseIf KeyCode = vbKeyDelete Then
       txtTransport = Empty
       txtTransport.Tag = Empty
    End If
    
Me.KeyPreview = True
    
End Sub

Private Sub SetLRDetail()
Dim RFRS As ADODB.Recordset
Set RFRS = New ADODB.Recordset
Dim FLAG As Boolean

If RFRS.State = 1 Then RFRS.Close
RFRS.Open "SELECT LRONCHLN FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND LRONCHLN ='Y'", CN, adOpenDynamic, adLockOptimistic
If Not RFRS.EOF Then
   FLAG = True
   Me.Height = 8040
   If RFRS.State = 1 Then RFRS.Close
   RFRS.Open "SELECT LRNO,LRDT,VEHICALNO,TRANSPORTMST.NAME AS TRANSPORT,VHCLMST.NAME AS VHCL FROM SPTRAN " & _
             "LEFT JOIN TRANSPORTMST ON TRANSPORTMST.CODE=SPTRAN.TRCD " & _
             "LEFT JOIN VHCLMST ON VHCLMST.CODE=SPTRAN.VEHICALNO " & _
             "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND DBCD='" & VTCD & _
             "' AND DATE>='" & Format(FSDT, "MM/dd/yyyy") & _
             "' AND DATE<='" & Format(FEDT, "MM/dd/yyyy") & "' ORDER BY DATE DESC", CN, adOpenDynamic, adLockOptimistic
   If Not RFRS.EOF Then
      TXTLRNO = Trim(RFRS!LRNO & "")
      If Not IsNull(RFRS!LRDT) Then
         LRDT = Format(RFRS!LRDT & "", "DD/MM/YYYY")
      Else
         LRDT = Format(TXTVBDT, "DD/MM/YYYY")
      End If
      TXTVHCL = Trim(RFRS!VHCL & "")
      txtTransport = Trim(RFRS!TRANSPORT & "")
   End If
   RFRS.Close
Else
   FLAG = False
   Me.Height = 7515
   RFRS.Close
End If

   LBLLRDT.Enabled = FLAG
   LRDT.Enabled = FLAG
   LBLLR.Enabled = FLAG
   TXTLRNO.Enabled = FLAG
   LBLTRCD.Enabled = FLAG
   txtTransport.Enabled = FLAG
   LBLVHCL.Enabled = FLAG
   TXTVHCL.Enabled = FLAG
   
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmBoxDispatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box/Carton/Pallet Dispatch Module : Based on Delivery Order"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   1110
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8391.647
   ScaleMode       =   0  'User
   ScaleWidth      =   12494.53
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   120
      Top             =   8520
   End
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   61
      Top             =   8640
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrTool 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1800
         Top             =   120
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
         TabIndex        =   62
         Top             =   0
         Width           =   120
      End
   End
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   8475
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   14949
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
      Begin VB.TextBox txtLRNO 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   26
         Top             =   8040
         Width           =   1215
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
         TabIndex        =   32
         Top             =   8040
         Width           =   3495
      End
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
         TabIndex        =   30
         Top             =   8040
         Width           =   1095
      End
      Begin VB.TextBox txtFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   91
         Top             =   7320
         Width           =   6165
      End
      Begin VB.TextBox TXTORDN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TXTAGENT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox TXTITEM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox TXTDONO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TXTGRAD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCONSINEE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1110
         Width           =   3135
      End
      Begin VB.TextBox txtQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox M_DORAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   20
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtLTNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtPCOD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1110
         Width           =   3735
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1110
         Width           =   3615
      End
      Begin VB.ComboBox M_RTTX 
         Height          =   315
         ItemData        =   "frmBoxDispatch.frx":0000
         Left            =   9480
         List            =   "frmBoxDispatch.frx":000A
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "M_RTTX"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox M_ARAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   21
         Top             =   2280
         Width           =   975
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
         ItemData        =   "frmBoxDispatch.frx":002B
         Left            =   2040
         List            =   "frmBoxDispatch.frx":002D
         TabIndex        =   5
         Tag             =   "0"
         Text            =   "cmbPackingType"
         Top             =   480
         Width           =   2655
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
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   22
         Top             =   2280
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   285
         Left            =   9240
         TabIndex        =   6
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
         Format          =   57344001
         CurrentDate     =   39347
      End
      Begin MSMask.MaskEdBox dtDate 
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Top             =   2280
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lstBox 
         Height          =   2775
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4895
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
         NumItems        =   12
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
            Text            =   "Packing Station Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ISRETURNABLE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TOP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "PLTNO"
            Object.Width           =   0
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   7800
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
         Image           =   "frmBoxDispatch.frx":002F
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   5640
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
         Image           =   "frmBoxDispatch.frx":03C9
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   6720
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
         Image           =   "frmBoxDispatch.frx":1153
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   8880
         TabIndex        =   4
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
         Image           =   "frmBoxDispatch.frx":15A5
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   4590
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
         Image           =   "frmBoxDispatch.frx":19F7
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   9840
         TabIndex        =   65
         Top             =   6480
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
         Image           =   "frmBoxDispatch.frx":1D91
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H btnLoadFile 
         Height          =   375
         Left            =   7680
         TabIndex        =   92
         Top             =   7320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "..."
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
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdImport 
         Height          =   375
         Left            =   8280
         TabIndex        =   93
         Top             =   7320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Import Data"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   10440
         Top             =   7320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker LRDT 
         Height          =   330
         Left            =   2880
         TabIndex        =   28
         Top             =   8040
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
         Format          =   57344001
         CurrentDate     =   38429
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
         TabIndex        =   25
         Top             =   8040
         Width           =   1095
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
         TabIndex        =   31
         Top             =   8040
         Width           =   1215
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
         TabIndex        =   29
         Top             =   8040
         Width           =   1215
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
         TabIndex        =   27
         Top             =   8040
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   11295
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Import File "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Tag             =   "S"
         Top             =   7320
         Width           =   1215
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
         TabIndex        =   90
         Top             =   5970
         Width           =   1335
      End
      Begin VB.Label Label20 
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
         TabIndex        =   89
         Top             =   5730
         Width           =   465
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
         TabIndex        =   88
         Top             =   5730
         Width           =   735
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
         TabIndex        =   87
         Top             =   5970
         Width           =   900
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
         TabIndex        =   86
         Top             =   5730
         Width           =   615
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
         TabIndex        =   85
         Top             =   5970
         Width           =   885
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
         TabIndex        =   84
         Top             =   5520
         Width           =   3015
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
         Left            =   3120
         TabIndex        =   83
         Top             =   5970
         Width           =   1215
      End
      Begin VB.Label Label18 
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
         Left            =   2040
         TabIndex        =   82
         Top             =   5730
         Width           =   945
      End
      Begin VB.Label Label17 
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
         Left            =   3120
         TabIndex        =   81
         Top             =   5730
         Width           =   1215
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
         Left            =   2160
         TabIndex        =   80
         Top             =   5970
         Width           =   900
      End
      Begin VB.Label Label2 
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
         Left            =   960
         TabIndex        =   79
         Top             =   5730
         Width           =   1095
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
         Left            =   1320
         TabIndex        =   78
         Top             =   5970
         Width           =   765
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
         Left            =   1320
         TabIndex        =   77
         Top             =   5520
         Width           =   3015
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   4500
         X2              =   4500
         Y1              =   6360
         Y2              =   5520
      End
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   7920
         X2              =   7920
         Y1              =   5520
         Y2              =   6360
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
         Left            =   4560
         TabIndex        =   76
         Top             =   5520
         Width           =   3015
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
         TabIndex        =   75
         Top             =   6480
         Width           =   495
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
         Left            =   600
         TabIndex        =   74
         Top             =   6480
         Width           =   855
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
         Left            =   120
         TabIndex        =   73
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
         Left            =   1560
         TabIndex        =   72
         Top             =   6720
         Width           =   435
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
         Left            =   3000
         TabIndex        =   71
         Top             =   6720
         Width           =   435
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
         Left            =   540
         TabIndex        =   70
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
         Left            =   1980
         TabIndex        =   69
         Top             =   6720
         Width           =   975
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
         Left            =   3420
         TabIndex        =   68
         Top             =   6720
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
         Left            =   2040
         TabIndex        =   67
         Top             =   6480
         Width           =   855
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
         Left            =   3360
         TabIndex        =   66
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label LBLORDER 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No."
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
         Left            =   720
         TabIndex        =   64
         Tag             =   "S"
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
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
         TabIndex        =   63
         Tag             =   "S"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   705
         Left            =   4500
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Width           =   6735
      End
      Begin VB.Shape ShapeSZO 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   705
         Left            =   75
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Width           =   4440
      End
      Begin VB.Shape BORDER3 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3495
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
         TabIndex        =   60
         Top             =   6000
         Width           =   885
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
         TabIndex        =   58
         Top             =   6000
         Width           =   900
      End
      Begin VB.Label Label13 
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
         Left            =   5760
         TabIndex        =   57
         Top             =   5760
         Width           =   540
      End
      Begin VB.Label Label10 
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
         Left            =   6840
         TabIndex        =   56
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
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
         TabIndex        =   55
         Top             =   5760
         Width           =   645
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
         Left            =   9000
         TabIndex        =   54
         Tag             =   "S"
         Top             =   2040
         Width           =   1215
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
         Left            =   2760
         TabIndex        =   53
         Tag             =   "S"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label LBLDO 
         BackStyle       =   0  'Transparent
         Caption         =   "D.O. No."
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
         Left            =   2160
         TabIndex        =   52
         Tag             =   "S"
         Top             =   2040
         Width           =   855
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
         Left            =   7200
         TabIndex        =   51
         Tag             =   "S"
         Top             =   1440
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1815
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   11175
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DO Qnty."
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
         TabIndex        =   50
         Tag             =   "S"
         Top             =   2040
         Width           =   1095
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
         Left            =   6000
         TabIndex        =   49
         Tag             =   "S"
         Top             =   2040
         Width           =   735
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
         Left            =   8400
         TabIndex        =   48
         Tag             =   "S"
         Top             =   1440
         Width           =   855
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
         Left            =   5520
         TabIndex        =   47
         Tag             =   "S"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LBLDODT 
         BackStyle       =   0  'Transparent
         Caption         =   "D.O. Date"
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
         Left            =   3480
         TabIndex        =   46
         Tag             =   "S"
         Top             =   2040
         Width           =   975
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
         TabIndex        =   45
         Tag             =   "S"
         Top             =   915
         Width           =   1455
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
         TabIndex        =   44
         Tag             =   "S"
         Top             =   915
         Width           =   2175
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
         TabIndex        =   43
         Tag             =   "S"
         Top             =   915
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Retail/Tax Invoice"
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
         Left            =   9480
         TabIndex        =   42
         Tag             =   "S"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ass. Rate"
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
         Left            =   6840
         TabIndex        =   41
         Tag             =   "S"
         Top             =   2040
         Width           =   1215
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
         TabIndex        =   40
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
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
         Left            =   5280
         TabIndex        =   39
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape BORDER2 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   5040
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   2175
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
         TabIndex        =   38
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   4575
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
         TabIndex        =   37
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Challan Da&te :"
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
         Left            =   7680
         TabIndex        =   36
         Top             =   480
         Width           =   1455
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
         TabIndex        =   35
         Top             =   480
         Width           =   1815
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
         Left            =   9000
         TabIndex        =   34
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label LBLHEAD 
         BackStyle       =   0  'Transparent
         Caption         =   "Challan No ."
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
         Left            =   7680
         TabIndex        =   33
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmBoxDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CSVTable As String
Public PLTREQ As String
Dim DIVCODE As String
Dim DIV_ACCESS_QTY As Double
Dim SAVEFLAG As Boolean
Dim ORDREQ As Boolean
Dim TTLCOPS As Long
Dim TTLBOXES As Long
Dim TTLQTY As Double, TTLGRSWGT As Double, TTLTAREWGT As Double
Dim SPARTY As String, SCONSINEE As String, SADD As String, grad As String, SUBGRD As String
Public VTCD As String
Public PKG_DBCD As String
Public ORDN As String
Public DONO As String
Public M_DBCD As String
Dim AMOUNT As String
Public chln As String
Dim M_VHCD As String, M_TRCD As String

Private Sub btnLoadFile_Click()
  Dim PARAM As String
  PARAM = " Z:\BACKDATA E:\BACKDATA "
  
  Shell App.PATH & "\REPORTS\SCANNER.BAT", vbHide
  'Shell App.PATH & "\REPORTS\SCANNER.BAT " & PARAM, vbHide
  'Shell App.PATH & "\REPORTS\PRINTDOC.BAT " & TXTFLE.FILENAME & "  prn", vbHide
        
  CommonDialog1.FILTER = "TXT File (*.txt)|*.*"
  CommonDialog1.InitDir = txtFile
  CommonDialog1.ShowOpen
  If CommonDialog1.FILENAME <> Empty Then txtFile = CommonDialog1.FILENAME Else txtFile = Empty
  If txtFile = Empty Then
     If txtFile.Enabled = True Then
        If txtFile.Visible = True Then
           txtFile.SetFocus
        End If
     End If
  End If
End Sub

Private Sub cmbPackingType_Click()
If cmdAdd.Enabled = False Then
   lstBox.ListItems.Clear
   Call CLEARDATA
   Call cmdCancel_Click
   TXTVBDT = Now
Else
   Call SetGlobal
   LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
End If
End Sub

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmdAdd_Click()
 zoomflag = False
 btn_sts (False)
 cmbPackingType.SetFocus
 SAVEFLAG = True
 PKG_DBCD = Empty
 If InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then      'PARTY REQUIRED in CASE OF JOB
    PKG_DBCD = "000002"
 End If
 
 frmDOList.PKG_DBCD = PKG_DBCD
 frmDOList.DIVCODE = DIVCODE
 frmDOList.Show 1
 
 Call SETPLTREQ
 
 If ORDN = Empty Or DONO = Empty Or M_DBCD = Empty Then
  Call cmdCancel_Click
 Else
  TXTVBDT.Enabled = False
 End If
End Sub

Private Sub cmdCancel_Click()
    Call ClsData(Me)
    Call btn_sts(True)
    lstBox.ListItems.Clear
    If zoomflag = True Then
        Call CMDEXIT_Click
        Exit Sub
    End If
    LBLS = 0:  LBLZ = 0: LBL0 = 0
    LBLSWGT = 0:  LBLZWGT = 0: LBLOWGT = 0
    txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
     txtRMNCTRN = Empty: txtRMNCOPs = Empty: txtRMNNTWT = Empty
     txtTTLCTRN = Empty: txtTTLCOPs = Empty: txtTTLNTWT = Empty
    TXTVBDT = Now
    TXTVBDT.Enabled = True
    LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
    If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 0
    If cmdAdd.Enabled Then cmdAdd.SetFocus
    
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("0017", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  SAVEFLAG = False
  chln = Empty
    
 PKG_DBCD = Empty
 If InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then    'PARTY REQUIRED in CASE OF JOB
    PKG_DBCD = "000002"
 End If
 frmEditDispatch.PKG_DBCD = PKG_DBCD
 Call SetGlobal
 frmEditDispatch.VTCD = VTCD
  
  frmEditDispatch.DIVCODE = DIVCODE
  frmEditDispatch.DIVNAME = LBLDIV
  frmEditDispatch.Show 1
  
  If chln <> Empty Then
     Call GetDetailedSelection
  End If
  
  If IsSaleExist Then
     MsgBox "Sale Bill Exist Against this Challan."
     lstBox.ListItems.Clear
     Call CLEARDATA
     Call cmdCancel_Click
     Exit Sub
  End If
    
  If chln <> Empty Then
     Call GetDetailedSelection
     TXTVBDT.Enabled = False
     LBLCHLN = chln
     btn_sts (False)
     cmbPackingType.SetFocus
  Else
     btn_sts (True)
     cmdAdd.SetFocus
  End If
  
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub cmdImport_Click()
  On Error Resume Next
  CSVTable = UCase(cUName) + "ROLLCSV"
  
  CN.Execute "DROP TABLE " & CSVTable
  CN.Execute "CREATE TABLE " & CSVTable & " (ROLLNO VARCHAR(40)) "
                            
  On Error GoTo LAST
   CN.Execute "BULK INSERT " & CSVTable & " FROM '" & txtFile.Text & "' WITH (firstrow=1,FIELDTERMINATOR=',',ROWTERMINATOR='\n')"
   
   
   MsgBox "Import File Successfully "
  Exit Sub
LAST:
  MsgBox "Error In file : " + txtFile + " Description : " + ERR.Description
  Exit Sub
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim RSTMP As ADODB.Recordset
Set RSTMP = New ADODB.Recordset
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
Dim BOXSTR As String

TOPBOTTOM = 0

FLAG = False
For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then: FLAG = True: Exit For
Next

If FLAG = False Then Exit Sub

Call SetGlobal

For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then
     TOPBOTTOM = TOPBOTTOM + Val(lstBox.ListItems(INDEX).SubItems(10))
     RETURNCOPS = RETURNCOPS + Val(lstBox.ListItems(INDEX).SubItems(1))
  End If
Next

If (Val(TTLQTY) > Val(txtQty) + DIV_ACCESS_QTY) Then
   MsgBox "Net Weight Exceed From DO Qty"
   Exit Sub
End If

Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT * FROM ORDTRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND VTYP='DOS' AND RECSTAT='A' AND DOSTAT='Y' AND ORDN='" & Trim(TXTORDN) & _
"' AND DONO='" & Trim(txtDONO) & "' AND DBCD='" & M_DBCD & "'"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

If Val(TTLQTY) > (Val(txtQty) + OrderBalanceQty(Trim(RECSET!ORDN), Trim(RECSET!ICOD), Trim(RECSET!grad))) Then
   MsgBox "Net Weight Exceed From Order Qty. Please Check Division Access Limit."
   Exit Sub
End If


If SAVEFLAG Then
   Dim NSQL As String
   Dim MSGS As String: MSGS = "Unit"
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
   AMOUNT = TTLQTY * Val(M_ARAT)
   AMOUNT = nstr(Round(AMOUNT, 0), 12, 2)
Else
  AMOUNT = TTLQTY * Val(M_ARAT)
End If


CN.BeginTrans

If SAVEFLAG Then ''INSERT MODE

LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)

SQL = "INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO,GRAD,SUBGRD,QNTY,RATE,ARAT,"
SQL = SQL & "ORDN,OSRC,BRMK,PRDL,ICOD,RTCD,TXRT,TXCD,SLIP,SLIPDATE,RDBC,DELQNTY,DFLG) VALUES ('" & compPth & "','" & UNCD & _
"','" & DIVCODE & "','DPF','" & M_DBCD & "','" & txtDONO & "','" & Format(RECSET!DODT, "YYYY/MM/DD") & _
"','" & RECSET!PCOD & "','" & RECSET!DCOD & "','" & RECSET!SRCH & "','" & RECSET!BRCD & "','" & RECSET!ltno & _
"','" & RECSET!grad & "','" & RECSET!SUBGRD & "','" & TTLQTY & "'," & RECSET!RATE & "," & RECSET!ARAT & _
",'" & RECSET!ORDN & "','1','" & Trim(TXTRMRK) & "','','" & RECSET!ICOD & "','" & RECSET!RTCD & "','" & RECSET!TXRT & "','" & RECSET!TXCD & _
"','" & LBLCHLN & "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & VTCD & "','" & RECSET!QNTY & "','Y')"

CN.Execute SQL

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,BRCD,"
SQL = SQL & "DCOD,ADDRESS,LTNO,ICOD,TXRT,GRAD,SUBGRD,PCES,QNTY,GWGT,TWGT,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA1,EXTRA2,EXTRA3,TXCD,RTCD,ARAT,LRNO,LRDT,VEHICALNO,TRCD)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','DPF','" & VTCD & "','" & LBLCHLN & "','" & LBLCHLN & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & RECSET!PCOD & "','" & RECSET!PCOD & _
"','" & RECSET!BRCD & "','" & RECSET!DCOD & "','" & RECSET!SRCH & "','" & RECSET!ltno & _
"','" & RECSET!ICOD & "','" & RECSET!TXRT & "','" & RECSET!grad & "','" & RECSET!SUBGRD & _
"','" & TTLBOXES & "','" & TTLQTY & "','" & TTLGRSWGT & "','" & TTLTAREWGT & _
"'," & RECSET!ARAT & "," & AMOUNT & _
",'Q','N','" & cUName & "','-','A','" & TTLCOPS & "','" & RECSET!ORDN & _
"','" & txtDONO & "','" & M_DBCD & "','" & Trim(RECSET!TXCD & "") & "','" & Trim(RECSET!RTCD & "") & _
"'," & RECSET!RATE & ",'" & Trim(TXTLRNO) & "','" & Format(LRDT, "YYYY/MM/DD") & "','" & M_VHCD & "','" & M_TRCD & "')"

CN.Execute SQL

SQL = "INSERT INTO PKGMAN (COMP,UNIT,DVCD,DBCD,VTYP,SRNO,SRCH,DATE,SLIPNO,PKG_STCOD,"
SQL = SQL & "LOTNO,FINITMCOD,GRAD,SUBGRAD,QNTY,SYSR,[USER],OPER,RECSTAT) VALUES "
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & DIVCODE & "','" & VTCD & "','DPF',"
SQL = SQL & "'1','1','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & LBLCHLN & "','000000',"
SQL = SQL & "'" & txtLTNo & "','" & RECSET!ICOD & _
"','" & RECSET!grad & "','" & RECSET!SUBGRD & "','" & TTLQTY & _
"','N','" & cUName & "','-','A')"

CN.Execute SQL

Dim UPSQL As String
UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & LBLCHLN & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
         "' AND VTYP='DPF' AND CODE='" & VTCD & "' AND FYCD='" & FYCD & "' "

If UNT_DIVSERIES_REQ = "Y" Then
   UPSQL = UPSQL & " AND DVCD='" & DIVCODE & "' "
End If
 
CN.Execute UPSQL

SQL = "UPDATE ORDTRN SET DFLG ='Y',SLIP='" & LBLCHLN & "',SLIPDATE='" & Format(TXTVBDT, "YYYY/MM/DD") & _
"',RDBC='" & VTCD & "',DELQNTY='" & TTLQTY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='" & M_DBCD & "' AND VTYP='DOS' AND "
SQL = SQL & "DFLG<>'Y' AND RECSTAT='A' AND DOSTAT='Y' AND ORDN = '" & Trim(RECSET!ORDN & "") & _
"' AND DONO  = '" & Trim(txtDONO) & "'"

CN.Execute SQL

SQL = "UPDATE ORDMAN SET LDSPDAT='" & Format(TXTVBDT, "MM/DD/YYYY") & _
"',DISPATCHQTY = DISPATCHQTY + " & TTLQTY & ",DOQTY = DOQTY - " & Val(txtQty) & " WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DCOD='" & DIVCODE & "' AND DBCD='" & M_DBCD & _
"' AND ORDN = '" & RECSET!ORDN & "' AND ICOD = '" & RECSET!ICOD & "' AND TRCD='" & RECSET!grad & "'"

CN.Execute SQL

If (Trim(RECSET!ISRETURNABLE & "") = "Y") Or (TOPBOTTOM > 0) Then

   SQL = "INSERT INTO PKGSTK(COMP,UNIT,DVCD,DBCD,VTYP,CHLN,DATE,PCOD,BRCD,DCOD,ADDRESS,OPER,"
   SQL = SQL & "QNTY,RECSTAT) VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & "','" & VTCD & _
   "', 'DPF','" & LBLCHLN & _
   "','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & RECSET!PCOD & "','" & RECSET!BRCD & _
   "','" & RECSET!DCOD & "','" & Trim(RECSET!SRCH) & "','-','" & RETURNCOPS & "','A')"
   
   CN.Execute SQL
   
End If

BOXSTR = Empty

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   If BOXSTR <> Empty Then BOXSTR = BOXSTR & ","
   BOXSTR = BOXSTR & "'" & lstBox.ListItems(INDEX).Text & "'"
 End If
Next INDEX

   SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & LBLCHLN & "',RVBDT= '" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "',RDBC = '" & VTCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
   "' AND VBNO IN (" & BOXSTR & ") AND (VTYP='PPF' OR VTYP='OPN') "
   
   CN.Execute SQL

CN.CommitTrans

Else  'EDIT UPDATE MODE

SQL = "UPDATE ORDTRN SET QNTY='" & TTLQTY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND RECSTAT='A' AND SLIP='" & LBLCHLN & _
"' AND RDBC = '" & VTCD & "' AND DBCD = '" & M_DBCD & "'"

CN.Execute SQL

SQL = "UPDATE ORDTRN SET DELQNTY='" & TTLQTY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='" & M_DBCD & "' AND VTYP='DOS' AND "
SQL = SQL & "DFLG='Y' AND RECSTAT='A' AND DOSTAT='Y' AND SLIP='" & LBLCHLN & _
"' AND RDBC = '" & VTCD & "'"

CN.Execute SQL

'COPS    'OK
SQL = "UPDATE SPTRAN SET PCES='" & TTLBOXES & "',QNTY='" & TTLQTY & "',AMNT= '" & AMOUNT & "',GWGT='" & TTLGRSWGT & "',TWGT='" & TTLTAREWGT & _
"' ,COPS = '" & TTLCOPS & "',LRNO='" & TXTLRNO & "',LRDT ='" & Format(LRDT, "YYYY/MM/DD") & "',VEHICALNO= '" & M_VHCD & "',TRCD ='" & M_TRCD & "' "
SQL = SQL & "WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND DVCD ='" & DIVCODE & "' AND DBCD ='" & VTCD & _
"'  AND VBNO ='" & LBLCHLN & "' AND VTYP='DPF' AND RECSTAT='A'"

CN.Execute SQL

SQL = "UPDATE PKGMAN SET QNTY = '" & TTLQTY & "' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD ='" & DIVCODE & "' AND DBCD ='" & VTCD & "' AND VTYP = 'DPF' AND SLIPNO ='" & LBLCHLN & "' AND PKG_STCOD='000000' AND RECSTAT='A'"

CN.Execute SQL

'CHECK DISPATCH QUANTITY  'OK
SQL = "UPDATE ORDMAN SET DISPATCHQTY = DISPATCHQTY - " & Val(txtNTWT.Tag) & " + " & TTLQTY & " WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DCOD='" & DIVCODE & "' AND DBCD='" & M_DBCD & _
"' AND ORDN = '" & TXTORDN & "' AND ICOD = '" & RECSET!ICOD & "' AND TRCD='" & RECSET!grad & "'"

CN.Execute SQL

If Trim(RECSET!ISRETURNABLE & "") = "Y" Then

SQL = "UPDATE PKGSTK SET QNTY = " & RETURNCOPS & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='" & VTCD & "' AND VTYP='DPF' AND CHLN='" & LBLCHLN & "' AND RECSTAT='A'"

CN.Execute SQL

End If

SQL = "UPDATE BOXREGISTER SET VTYP=PVTYP,RVBNO=NULL,RVBDT= NULL,RDBC = NULL,RVTYP = NULL WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND RVBNO='" & LBLCHLN & "' AND RDBC = '" & VTCD & _
"' AND RVTYP='DPF' AND VTYP='DPF'"

CN.Execute SQL

BOXSTR = Empty

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   If BOXSTR <> Empty Then BOXSTR = BOXSTR & ","
   BOXSTR = BOXSTR & "'" & lstBox.ListItems(INDEX).Text & "'"
 End If
Next INDEX

   SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & LBLCHLN & "',RVBDT= '" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "',RDBC = '" & VTCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
   "' AND (VTYP='PPF' OR VTYP='OPN') AND VBNO IN (" & BOXSTR & ") "
   
   CN.Execute SQL

'---------------------------------
'DAILYSTATUS ENTRY
If SAVEFLAG = True Then
  Call DAILYSTATUS("DPF", GetCode("ACCMST", txtpcod, "NAME", "CODE"), VTCD, Val(txtNTWT), LBLCHLN, 0, cUName, "N", Now, TXTVBDT)
  Else
  Call DAILYSTATUS("DPF", GetCode("ACCMST", txtpcod, "NAME", "CODE"), VTCD, Val(txtNTWT), LBLCHLN, 0, cUName, "M", Now, TXTVBDT)
 End If
'---------------------------------
CN.CommitTrans

End If

'PLY UPDATION COMMON FOR BOTH SAVE AND EDIT
If TOPBOTTOM > 0 Then
Dim NOOFPLY As Double
Dim i As Long, J As Long, K As Long
If RSTMP.State = 1 Then RSTMP.Close
RSTMP.Open "SELECT * FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND DBCD='" & VTCD & "' AND VTYP='DPF' AND CHLN='" & LBLCHLN & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic

If Not RSTMP.EOF Then
i = 0
  For i = 13 To lstBox.ColumnHeaders.COUNT
    J = 0
    For J = 0 To RSTMP.Fields.COUNT - 1
      If Trim(RSTMP.Fields(J).NAME) = Trim(lstBox.ColumnHeaders(i).Text) Then
            NOOFPLY = 0
         '---------------------------------------------------------------------
            For K = 1 To lstBox.ListItems.COUNT
               If lstBox.ListItems(K).Checked = True Then
                  NOOFPLY = NOOFPLY + Val(lstBox.ListItems(K).SubItems(i - 1))
               End If
            Next
         '---------------------------------------------------------------------
         
         RSTMP.Fields(J).Value = NOOFPLY
      End If
    Next
  Next
  
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

lstBox.ListItems.Clear
Call CLEARDATA
Call cmdCancel_Click

TXTVBDT = Now

Exit Sub
LAST:
MsgBox ERR.Description
Resume
CN.RollbackTrans
Exit Sub
End Sub

Private Sub cmdSavePrint_Click()
Exit Sub
        Call cmdSave_Click
   
   'FOR ONLINE CHALLAN PRINTING
   If IsOnlineChallanPrintReq Then
   
        LOAD frmRPT_DelChallanPrint
        frmRPT_DelChallanPrint.Hide
        
        frmRPT_DelChallanPrint.cboStatus.ListIndex = 0
        frmRPT_DelChallanPrint.txtUNIT.Tag = UNCD
        frmRPT_DelChallanPrint.txtDVCD.Tag = DIVCOD
        frmRPT_DelChallanPrint.cmbDispatchType.AddItem cmbPackingType.Text
        frmRPT_DelChallanPrint.txtUNIT = UntNm
        frmRPT_DelChallanPrint.txtDVCD = LBLDIV.Caption
        frmRPT_DelChallanPrint.cmbDispatchType.Text = cmbPackingType
        frmRPT_DelChallanPrint.lstCHLN_GotFocus
        frmRPT_DelChallanPrint.opPrePrinted.Value = True
        frmRPT_DelChallanPrint.cmdpreview_Click
    End If
    
End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or LBLDIV.Caption = Empty Then
   Unload Me
End If
Call FindDivAccessQty
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)

PLTREQ = "N"

TXTADDRESS.FontBold = False
Me.Left = 50: Me.KeyPreview = True
SAVEFLAG = True
  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  DIVCODE = Empty
  If DIVCODE = Empty Then
    LBLDIV.Caption = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
   
  Call SetLabel
  Call SetPackingType
  
  TXTVBDT = Now
  LRDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
    
  LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
  
  
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

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT SCANIMP FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GETRS.EOF Then
   txtFile = Trim(GETRS!SCANIMP & "")
End If
GETRS.Close

Call SetLRDetail

'lstBox.ColumnHeaders(11).Text
End Sub

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT DISTINCT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND CODE NOT IN ('000003','000004') AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "' AND NAME<>''", CN, adOpenDynamic, adLockOptimistic

If Not PKTYPRS.EOF Then VTCD = Trim(PKTYPRS!CODE)
Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!NAME)
PKTYPRS.MoveNext
Loop
If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 0
End Sub

Private Sub TimerBillNo2_Timer()
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

Private Sub txtFile_GotFocus()
  txtFile.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
     txtFile = Empty
     CSVTable = Empty
  End If
End Sub

Private Sub txtFile_LostFocus()
  txtFile.BackColor = vbWhite
End Sub

Private Sub TXTLRNO_GotFocus()
  TXTLRNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTLRNO_LostFocus()
 TXTLRNO.BackColor = vbWhite
End Sub

Private Sub LRDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub txtTransport_GotFocus(): txtTransport.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtTransport_LostFocus(): txtTransport.BackColor = vbWhite: End Sub

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

Private Sub txtPCOD_GotFocus(): txtpcod.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtConsinee_GotFocus(): txtCONSINEE.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTADDRESS_GotFocus(): TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtAgent_GotFocus(): TXTAGENT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtItem_GotFocus(): txtITEM.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTGRAD_GotFocus(): TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTSUBGRD_GotFocus(): TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtltno_GotFocus(): txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtDONO_GotFocus(): txtDONO.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTORDN_GotFocus(): TXTORDN.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub dtDate_KeyDown(KeyCode As Integer, Shift As Integer): dtDate.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTQTY_GotFocus(): txtQty.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub M_DORAT_GotFocus(): M_DORAT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub M_ARAT_GotFocus(): M_ARAT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTRMRK_GotFocus(): TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub

Private Sub txtPCOD_LostFocus(): txtpcod.BackColor = vbWhite: End Sub
Private Sub txtConsinee_LostFocus(): txtCONSINEE.BackColor = vbWhite: End Sub
Private Sub TXTADDRESS_LostFocus(): TXTADDRESS.BackColor = vbWhite: End Sub
Private Sub txtAgent_LostFocus(): TXTAGENT.BackColor = vbWhite: End Sub
Private Sub txtItem_LostFocus(): txtITEM.BackColor = vbWhite: End Sub
Private Sub TXTGRAD_LostFocus(): TXTGRAD.BackColor = vbWhite: End Sub
Private Sub TXTSUBGRD_LostFocus(): TXTSUBGRD.BackColor = vbWhite: End Sub
Private Sub txtltno_LostFocus(): txtLTNo.BackColor = vbWhite: End Sub
Private Sub TXTDONO_LostFocus(): txtDONO.BackColor = vbWhite: End Sub
Private Sub TXTORDN_LostFocus(): TXTORDN.BackColor = vbWhite: End Sub
Private Sub dtDate_LostFocus(): dtDate.BackColor = vbWhite: End Sub
Private Sub TXTQTY_LostFocus(): txtQty.BackColor = vbWhite: End Sub
Private Sub M_DORAT_LostFocus(): M_DORAT.BackColor = vbWhite: End Sub
Private Sub M_ARAT_LostFocus(): M_ARAT.BackColor = vbWhite: End Sub
Private Sub TXTRMRK_LostFocus(): TXTRMRK.BackColor = vbWhite: End Sub
Private Sub CLEARDATA()
 txtpcod = Empty: txtCONSINEE = Empty: TXTADDRESS = Empty: txtITEM = Empty: txtLTNo = Empty: TXTGRAD = Empty: TXTSUBGRD = Empty
 M_RTTX = Empty: txtDONO = Empty: txtQty = Empty: M_DORAT = Empty: M_ARAT = Empty: TXTRMRK = Empty
 txtCTRN = Empty: txtCops = Empty: txtNTWT = Empty
 TXTVBDT.Enabled = True
 TXTVBDT = Format(Now, "DD/MM/YYYY")
 
End Sub
Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub
Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub SetGlobal()
Dim INDEX As Long
TTLQTY = 0: TTLBOXES = 0: TTLCOPS = 0: TTLGRSWGT = 0: TTLTAREWGT = 0

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


'TRANSPORT CODE
If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT * FROM TRANSPORTMST WHERE NAME ='" & txtTransport & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   M_TRCD = Trim(DBCDRS!CODE & "")
Else
   M_TRCD = Empty
End If
DBCDRS.Close

'VEHICLE CODE
If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT * FROM VHCLMST WHERE NAME ='" & TXTVHCL & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   M_VHCD = Trim(DBCDRS!CODE & "")
Else
   M_VHCD = Empty
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
End Sub

Public Sub GetDetailedSelection()
Dim INDEX As Long
txtCTRN.Caption = 0: txtNTWT.Caption = 0: txtCops.Caption = 0
LBLS = 0: LBLZ = 0: LBL0 = 0
LBLSWGT = 0: LBLZWGT = 0: LBLOWGT = 0

txtRMNCOPs = 0
txtRMNCTRN = 0
txtRMNNTWT = 0

For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then
     
     txtNTWT.Caption = nstr(Val(txtNTWT.Caption) + Val(lstBox.ListItems(INDEX).SubItems(2)), 10, 3)
     txtCTRN.Caption = Val(txtCTRN.Caption) + 1
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
     
     If (Val(txtNTWT) > (Val(txtQty) + DIV_ACCESS_QTY)) Then
           MsgBox "Net Weight Exceed From DO Qty"
           lstBox.SetFocus
           Exit Sub
     End If
  End If
Next

txtRMNCOPs = Val(txtTTLCOPs) - Val(txtCops)
txtRMNCTRN = Val(txtTTLCTRN) - Val(txtCTRN)
txtRMNNTWT = Val(txtTTLNTWT) - Val(txtNTWT)

End Sub

Private Sub lstBox_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Msg ("Press F8 For Select All F9 For De-Select All")
    Call GetDetailedSelection
    If Item.INDEX < lstBox.ListItems.COUNT Then lstBox.ListItems.Item(Item.INDEX + 1).Selected = True: lstBox.ListItems(Item.INDEX + 1).EnsureVisible
End Sub

Public Sub btn_sts(Yes As Boolean)
 cmdSave.Enabled = Not Yes: cmdCancel.Enabled = Not Yes: cmdAdd.Enabled = Yes: cmdEdit.Enabled = Yes
 
 txtpcod.Enabled = Not Yes: txtCONSINEE.Enabled = Not Yes: TXTADDRESS.Enabled = Not Yes: txtITEM.Enabled = Not Yes: txtLTNo.Enabled = Not Yes: TXTGRAD.Enabled = Not Yes: TXTSUBGRD.Enabled = Not Yes
 M_RTTX.Enabled = Not Yes: txtDONO.Enabled = Not Yes: txtQty.Enabled = Not Yes: M_DORAT.Enabled = Not Yes: M_ARAT.Enabled = Not Yes: TXTRMRK.Enabled = Not Yes
End Sub

Private Sub FindDivAccessQty()
Dim divrs As ADODB.Recordset
Set divrs = New ADODB.Recordset

If divrs.State = 1 Then divrs.Close
divrs.Open "SELECT DOQTYLIMIT FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & DIVCODE & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not divrs.EOF Then
   DIV_ACCESS_QTY = Val(divrs!DOQTYLIMIT)
Else
   DIV_ACCESS_QTY = 0
End If

End Sub

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

Private Function OrderBalanceQty(ODNO As String, ITMCOD As String, GRCD As String) As Double
'DEFAULT
OrderBalanceQty = 0
'----------------
   Dim QUERY As String
   QUERY = "SELECT ISNULL(QNTY - DOQTY - DISPATCHQTY - CANCELQTY,0) AS BALQTY FROM ORDMAN WHERE COMP='" & compPth & _
           "' AND UNIT='" & UNCD & "' AND DCOD='" & DIVCODE & "' AND DBCD='" & M_DBCD & _
           "' AND ORDN = '" & ODNO & "' AND ICOD = '" & ITMCOD & "' AND TRCD='" & GRCD & "'"
           
   If RS.State = 1 Then RS.Close
   RS.Open QUERY, CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      OrderBalanceQty = Val(RS!BALQTY)
   End If
End Function

Private Sub SETPLTREQ()
On Error GoTo LAST
   
PLTREQ = "N"
   
If RS.State = 1 Then RS.Close
RS.Open "SELECT ISEXPORTORDER FROM SALMANMST WHERE CODE='" & M_DBCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   PLTREQ = IIf(Val(RS!ISEXPORTORDER) = 0, "N", "Y")
End If

Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Sub SetLabel()
Dim TYP As String
   TYP = LabelType(DIVCODE, UNCD)
   If TYP = "GD" Or TYP = "SD" Then
     LBLCFG2.Visible = False
     TXTSUBGRD.Visible = False
     LBLCFG.Caption = LabelDisplay(DIVCODE, UNCD)
   End If
   
   If IsTwistReq(DIVCODE) = "Y" Then
      LBLCFG2.Caption = "Twist"
      LBLCFG2.Visible = True
      TXTSUBGRD.Visible = True
      ShapeSZO.Visible = True: Label1.Visible = True
      LabelS.Visible = True:  LabelZ.Visible = True:  LabelO.Visible = True
      LBLS.Visible = True:  LBLZ.Visible = True:  LBL0.Visible = True
      LBLSWGT.Visible = True: LBLZWGT.Visible = True: LBLOWGT.Visible = True
   ElseIf SetIsShadeReq(DIVCODE) = "Y" Then
      LBLCFG2.Caption = "Shade"
      lstBox.ColumnHeaders(4).Text = "Shade"
   Else
      ShapeSZO.Visible = False: Label1.Visible = False
      LabelS.Visible = False: LabelZ.Visible = False: LabelO.Visible = False
      LBLS.Visible = False: LBLZ.Visible = False: LBL0.Visible = False
      LBLSWGT.Visible = False: LBLZWGT.Visible = False: LBLOWGT.Visible = False
   End If
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


Private Sub SetLRDetail()
Dim RFRS As ADODB.Recordset
Set RFRS = New ADODB.Recordset
Dim FLAG As Boolean

If RFRS.State = 1 Then RFRS.Close
RFRS.Open "SELECT LRONCHLN FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND LRONCHLN ='Y'", CN, adOpenDynamic, adLockOptimistic
If Not RFRS.EOF Then
   FLAG = True
   Me.Height = 8790
Else
   FLAG = False
   Me.Height = 8235
End If
RFRS.Close

   LBLLRDT.Enabled = FLAG
   LRDT.Enabled = FLAG
   LBLLR.Enabled = FLAG
   TXTLRNO.Enabled = FLAG
   LBLTRCD.Enabled = FLAG
   txtTransport.Enabled = FLAG
   LBLVHCL.Enabled = FLAG
   TXTVHCL.Enabled = FLAG
   
End Sub


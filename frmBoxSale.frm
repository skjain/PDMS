VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmBoxSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   13470
   Begin VB.Frame fraTakaDet 
      BackColor       =   &H00C0E0FF&
      Height          =   6240
      Left            =   10440
      TabIndex        =   39
      Top             =   600
      Width           =   3135
      Begin VB.TextBox txtTotalPcs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5640
         Width           =   675
      End
      Begin VB.TextBox txtTotalQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1035
      End
      Begin MSComctlLib.ListView lstRolls 
         Height          =   5295
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Box No"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NetWt."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ROW"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "GRWGT"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "TRWGT"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "RATE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ICOD"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00CFFCFE&
         BackStyle       =   0  'Transparent
         Caption         =   "Box :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00CFFCFE&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   42
         Top             =   5640
         Width           =   345
      End
   End
   Begin VB.ComboBox cmbSaleType 
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
      ItemData        =   "frmBoxSale.frx":0000
      Left            =   1440
      List            =   "frmBoxSale.frx":0002
      TabIndex        =   38
      Tag             =   "0"
      Text            =   "Select Type of Sale"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox TXTRMRK 
      Height          =   285
      Left            =   1560
      MaxLength       =   30
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   5880
      Width           =   4815
   End
   Begin VB.TextBox TXTCRDS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox TXTITOT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
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
      Left            =   4920
      TabIndex        =   37
      Text            =   "0.00"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox TXTTQTY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
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
      Left            =   2760
      TabIndex        =   36
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox TXTTPCS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
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
      Left            =   1080
      TabIndex        =   35
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmddelitm 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Remove Item"
      Height          =   315
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame frm_head 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      TabIndex        =   22
      Top             =   600
      Width           =   11175
      Begin VB.ComboBox TXTRTORTAX 
         Height          =   315
         ItemData        =   "frmBoxSale.frx":0004
         Left            =   5040
         List            =   "frmBoxSale.frx":000E
         TabIndex        =   5
         Text            =   "Select Tax Category"
         Top             =   960
         Width           =   2970
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox TXTDBAC 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
      Begin VB.TextBox TXTBRNM 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox TXTTAXNAM 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TXTVBNO 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXTCOMINV 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox TXTDLPTY 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TXTGDN 
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   3015
      End
      Begin MSMask.MaskEdBox TXTVBDT 
         Height          =   330
         Left            =   9060
         TabIndex        =   8
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
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
         Format          =   "dd/MM/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Address."
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
         Left            =   3960
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LBLDRAC 
         Caption         =   "Db A/c Name"
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
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LBLBRNM 
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
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LBLTAXNAM 
         Caption         =   "Tax Reference"
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
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LBLBILLNO 
         Caption         =   "Bill No. "
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
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LBLCOMMINV 
         Caption         =   "Comm Invoice No"
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
         TabIndex        =   28
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label LBLBILLDATE 
         Caption         =   "Bill Date"
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
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LBLDLPTY 
         Caption         =   "Consignee "
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
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LBLRTORTX 
         Caption         =   "Retail/Tax"
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
         Left            =   3960
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   8160
         X2              =   8160
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1695
         Left            =   0
         Top             =   120
         Width           =   10455
      End
      Begin VB.Label Label7 
         Caption         =   "GoDown"
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
         Left            =   3960
         TabIndex        =   24
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame FRMBTRM 
      Height          =   2415
      Left            =   6600
      TabIndex        =   17
      Top             =   4440
      Width           =   3735
      Begin VB.TextBox TXTBNET 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1965
         Width           =   1905
      End
      Begin VB.TextBox txtBEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox TXTADLS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2040
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flexBTRM 
         Height          =   1635
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   0
         Cols            =   3
         FixedRows       =   0
         Appearance      =   0
      End
      Begin VB.Label LBLNET 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   21
         Top             =   2040
         Width           =   1305
      End
   End
   Begin VB.Frame ITMFRM 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      TabIndex        =   16
      Top             =   2520
      Width           =   11055
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   1815
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Image           =   "frmBoxSale.frx":002F
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   3360
      TabIndex        =   44
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Image           =   "frmBoxSale.frx":03C9
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   4440
      TabIndex        =   45
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Image           =   "frmBoxSale.frx":0763
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Image           =   "frmBoxSale.frx":0AFD
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   2280
      TabIndex        =   46
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Image           =   "frmBoxSale.frx":1887
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   5520
      TabIndex        =   47
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Image           =   "frmBoxSale.frx":1CD9
      cBack           =   -2147483633
   End
   Begin VB.Label LBLCRAC 
      Caption         =   "Type of Sale"
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
      TabIndex        =   55
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   240
      TabIndex        =   54
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Days"
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
      Left            =   240
      TabIndex        =   53
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label LBLGRS 
      Alignment       =   1  'Right Justify
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
      Left            =   5400
      TabIndex        =   52
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Total Qty"
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
      Left            =   1800
      TabIndex        =   51
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Total Box"
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
      TabIndex        =   50
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Grs Amt "
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
      Left            =   4080
      TabIndex        =   49
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label LBLDIV 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "DIVISION : "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmBoxSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DIVCODE As String
Dim DIVNAME As String
Public M_DBCD As String
Dim M_OPER(0 To 15) As String
Dim M_PERC(0 To 15) As Double
Dim M_POSTCOD(0 To 15) As String
Dim M_NICK(0 To 15) As String
Dim M_POSTYESNO(0 To 15) As String
Dim M_FMLA(0 To 15) As String
Dim M_RDOF(0 To 15) As String
Dim DRAC As String
Dim PCOD As String
Dim DCOD As String
Dim ADDRESS As String
Dim BRCD As String
Dim CPCD As String
Dim ARCD As String
Dim TXCD As String
Dim TTYP As String
Dim SAVEFLAG  As Boolean
Dim itemcode As String
Dim chgFlag As Boolean
Dim calbtm As Boolean
Dim CHK_FLX As Boolean
Dim FLXROW As Long
Dim FLXCOL As Long
Dim Emptycell As Boolean

Private Sub cmbSaleType_Click()
 SendKeys "{HOME}"
 Call FindSerial
End Sub

Private Sub cmbSaleType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
      SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbSaleType_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Public Sub cmdAdd_Click()
  cmbSaleType.Enabled = True
  zoomflag = False
  SAVEFLAG = True
  btn_sts (False)
  cmddelitm.Enabled = False
  Call FindSerial
  If cmbSaleType.Enabled Then cmbSaleType.SetFocus
  TXTDBAC.Enabled = True
  TXTDBAC.SetFocus
  
End Sub

Private Sub cmdCancel_Click()
  Dim LINDEX As Long
  LINDEX = cmbSaleType.ListIndex
  ClsData (Me)
  cmbSaleType.ListIndex = LINDEX
  FLEX.Clear
  FLEX.Rows = 2
  lstRolls.ListItems.Clear
  btn_sts (True)
  Call setflexhead
  If zoomflag = True Then
    Call CMDEXIT_Click
    Exit Sub
  End If
  cmdAdd.SetFocus
  
  Dim i As Integer
  For i = 0 To flexBTRM.Rows - 1
    flexBTRM.TextMatrix(i, 2) = "0.00"
  Next
  TXTBNET.Text = "0.00"
  
  cmbSaleType.Enabled = True
End Sub

Private Sub cmdDelete_Click()
  SAVEFLAG = False
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  btn_sts (False)
  FRM_SPLISTDIRSAL.Show 1
  
    Dim AYS
    AYS = MsgBox("Are you sure to delete the invoice ", vbYesNo)
    If AYS = vbYes Then
      CN.BeginTrans
    
      Dim m_rtyp As String
      Dim m_rsrn As String
      
      Dim SAL_SALCOD As String
      Dim SAL_PTYCOD As String
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        SAL_PTYCOD = RS!CODE
      End If
      
      SAL_SALCOD = "xxxxxx"  'RS!CODE
            
      If SAVDAT.State = 1 Then SAVDAT.Close
      SAVDAT.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & _
                  "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
      Do While Not SAVDAT.EOF
          m_rtyp = SAVDAT!RTYP & ""
          m_rsrn = SAVDAT!RSRN & ""
          CN.Execute "UPDATE SPTRAN SET RTYP=NULL, RSRN=NULL WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                     "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & _
                     "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
          SAVDAT.MoveNext
      Loop
      CN.Execute "UPDATE SPTRAN set recstat='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & _
                 "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
                 
      CN.Execute "UPDATE STORETRAN set recstat='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & _
                 "' AND RECSTAT<>'D'"
                 
      CN.Execute "UPDATE BILLMAIN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & _
                 "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
                 
      CN.Execute "UPDATE EGPMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
      
      
      Call DAILYSTATUS("SAL", GetCode("ACCMST", TXTDBAC, "NAME", "CODE"), M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "D", Now, TXTVBDT)
      CN.CommitTrans
    End If
  Call cmdCancel_Click
  If zoomflag = True Then
    Call CMDEXIT_Click
    Exit Sub
  End If

End Sub

Private Sub cmddelitm_Click()

  If SAVEFLAG = False Then
     Exit Sub
  End If

  If FLEX.ROW > 1 Then
    FLEX.RemoveItem (FLEX.ROW)
    TXTTPCS.Text = 0
    TXTTQTY.Text = 0
    TXTITOT.Text = 0
    Dim i As Double
    i = 1
    For i = 1 To FLEX.Rows - 1
      FLEX.TextMatrix(i, 0) = i
      TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 2)), "######")
      TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 3)), "########.000")
      TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 6)), "########.00")
    Next
    FLEX.Refresh
    FLEX.ROW = FLEX.Rows - 1
    FLEX.COL = 1
    FLEX.SetFocus
  Else
   If FLEX.ROW = 1 Then
      MsgBox "Replace First Item From Another Item"
      Exit Sub
   End If
  End If
  cmddelitm.Enabled = False
  
End Sub

Private Sub cmddelitm_LostFocus()
  cmddelitm.Enabled = False
End Sub

Private Sub cmdEdit_Click()
  cmbSaleType.Enabled = False
  
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("0020", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  SAVEFLAG = False
  ' FRM_SPLISTDIRSAL.Show 1
    frmBoxSaleList.Show 1
    
    
   'Check for Receipt and Payment Entires
  If TXTVBNO = Empty Then
    Call cmdCancel_Click
    Exit Sub
  End If
  
  btn_sts (False)
  TXTDBAC.SetFocus
  
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub Flex_Click()

  cmddelitm.Enabled = True
 ' Call FillList(FLEX.ROW)
End Sub

Private Sub FLEX_EnterCell()
  FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
  Emptycell = True
End Sub

Private Sub FLEX_GotFocus()
  Me.KeyPreview = False
  FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.ROW
End Sub

Private Sub FLEX_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF2 And FLEX.COL = 1) Or (KeyCode = vbKeyReturn And FLEX.COL = 1) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        FLEX.TextMatrix(FLEX.ROW, 1) = SearchList1("select  TOP 20 code,name from itmmst", 0, FLEX.TextMatrix(FLEX.ROW, 1), "SELECT ITEM FROM LIST")
        If Trim(FLEX.TextMatrix(FLEX.ROW, 1)) <> Empty Then
           Call FillList(FLEX.ROW)
           lstRolls.SetFocus
        End If
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            FLEX.TextMatrix(FLEX.ROW, 3) = ""
            frm_Item.Show
        End If
        FLEX.TextMatrix(FLEX.ROW, 8) = Key
    End If
End Sub

Private Sub Flex_LeaveCell()
  Dim FLEXROW As Double
  Dim FLEXCOL As Double
  Dim i As Double
  FLEX.CellBackColor = vbWhite
  If Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "Q" Then
    FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 3)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
  ElseIf Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "P" Then
    FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 2)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
  End If
  FLEXROW = FLEX.ROW
  FLEXCOL = FLEX.COL
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 2)), "######")
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 3)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 6)), "########.00")
  Next
  FLEX.ROW = FLEXROW
  FLEX.COL = FLEXCOL
  FLEX.SetFocus
End Sub

Private Sub FLEX_LostFocus()
  Dim FLEXROW As Double
  Dim FLEXCOL As Double
  Dim i As Double
  FLEX.CellBackColor = vbWhite
  If Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "Q" Then
    FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 3)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
  ElseIf Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "P" Then
    FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 2)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
  End If
  FLEXROW = FLEX.ROW
  FLEXCOL = FLEX.COL
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 2)), "######")
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 3)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 6)), "########.00")
  Next
  FLEX.ROW = FLEXROW
  FLEX.COL = FLEXCOL
End Sub

Private Sub Form_Activate()
  'Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  DIVCOD = Me.Tag
  Dim divisionmaster As New ADODB.Recordset
  Set divisionmaster = New ADODB.Recordset
  If divisionmaster.State = 1 Then divisionmaster.Close
  divisionmaster.Open "select * from DIVMST where code='" & DIVCOD & "' AND UNIT='" & UNCD & "' and comp='" & compPth & "'", CN
  If Not divisionmaster.EOF Then
    DIVNAM = divisionmaster!NAME
    DIVCOD = divisionmaster!CODE
   Else
    DIVNAM = "??????"
  End If
  LBLDIV.Caption = "DIVISION : " + DIVNAM
  If DIVNAM = "??????" Then
    Unload Me
  End If

  FRMPARA = "SAL"
  If DIVNAM = "??????" Then
    Unload Me
  End If
  If zoomflag = True Then
    SAVEFLAG = False
    btn_sts (False)
  End If
  
  If DIVCOD = "000001" And UNT_MMS_INSTALL = "Y" Then
   '  LBLDEPT.Enabled = True: txtDept.Enabled = True
  Else
    ' LBLDEPT.Enabled = False: txtDept.Enabled = False
  End If
  
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
  If (ActiveControl.NAME = "TXTCRAC" Or ActiveControl.NAME = "FLEX" Or ActiveControl.NAME = "TXTDBAC" Or ActiveControl.NAME = "TXTDLPTY" Or ActiveControl.NAME = "TXTBRNM" Or ActiveControl.NAME = "TXTTAXNAM") Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)

TXTVBDT = Format(Now, "DD/MM/YYYY")

  flexBTRM.ColWidth(0) = 1500
  flexBTRM.ColWidth(1) = 800
  flexBTRM.ColWidth(2) = 1200
  
  Emptycell = True
  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  DIVCODE = Empty: DIVNAME = Empty
  DIVCODE = "000001"
  DIVNAME = "STORE DIVISION"
  
  Me.Tag = "000001"
  Me.Caption = "SALE MODULE ( DIVISION : " + DIVNAME + " )"
  
  TXTVBDT = Format(Now, "DD/MM/YYYY")
  Call btn_sts(True)
  Call setflexhead
  Call SetSaleType
End Sub

Private Sub setflexhead()
    FLEX.TextMatrix(0, 0) = "Sr."
    FLEX.TextMatrix(0, 1) = "Item Name"
    FLEX.TextMatrix(0, 2) = "Pcs"
    FLEX.TextMatrix(0, 3) = "Qnty"
    FLEX.TextMatrix(0, 4) = "Rate"
    FLEX.TextMatrix(0, 5) = "$."
    FLEX.TextMatrix(0, 6) = "Amount"
    FLEX.TextMatrix(0, 7) = "Remarks"
    FLEX.TextMatrix(0, 8) = "Icod"
    FLEX.TextMatrix(0, 9) = "RollDetails"
    
    FLEX.ColWidth(0) = 300
    FLEX.ColWidth(1) = 3000
    FLEX.ColWidth(2) = 600
    FLEX.ColWidth(3) = 1000
    FLEX.ColWidth(4) = 900
    FLEX.ColWidth(5) = 250
    FLEX.ColWidth(6) = 1500
    FLEX.ColWidth(7) = 3250
    FLEX.ColWidth(8) = 0
    FLEX.ColWidth(9) = 8000
    
    FLEX.ColAlignment(1) = 0
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    frm_head.Enabled = Not Yes
    ITMFRM.Enabled = Not Yes
    FRMBTRM.Enabled = Not Yes
End Sub

Private Sub FIL_Billingterm()
Dim CNTR As Byte
  flexBTRM.Clear
  flexBTRM.Rows = 0
  M_OPER(0) = ""
  M_OPER(1) = ""
  M_OPER(2) = ""
  M_OPER(3) = ""
  M_OPER(4) = ""
  M_OPER(5) = ""
  M_OPER(6) = ""
  M_OPER(7) = ""
  M_OPER(8) = ""
  M_OPER(9) = ""
  M_PERC(0) = 0
  M_PERC(1) = 0
  M_PERC(2) = 0
  M_PERC(3) = 0
  M_PERC(4) = 0
  M_PERC(5) = 0
  M_PERC(6) = 0
  M_PERC(7) = 0
  M_PERC(8) = 0
  M_PERC(9) = 0
  M_POSTCOD(0) = ""
  M_POSTCOD(1) = ""
  M_POSTCOD(2) = ""
  M_POSTCOD(3) = ""
  M_POSTCOD(4) = ""
  M_POSTCOD(5) = ""
  M_POSTCOD(6) = ""
  M_POSTCOD(7) = ""
  M_POSTCOD(8) = ""
  M_POSTCOD(9) = ""
  M_NICK(0) = ""
  M_NICK(1) = ""
  M_NICK(2) = ""
  M_NICK(3) = ""
  M_NICK(4) = ""
  M_NICK(5) = ""
  M_NICK(6) = ""
  M_NICK(7) = ""
  M_NICK(8) = ""
  M_NICK(9) = ""
  M_POSTYESNO(0) = ""
  M_POSTYESNO(1) = ""
  M_POSTYESNO(2) = ""
  M_POSTYESNO(3) = ""
  M_POSTYESNO(4) = ""
  M_POSTYESNO(5) = ""
  M_POSTYESNO(6) = ""
  M_POSTYESNO(7) = ""
  M_POSTYESNO(8) = ""
  M_POSTYESNO(9) = ""
  M_FMLA(0) = ""
  M_FMLA(1) = ""
  M_FMLA(2) = ""
  M_FMLA(3) = ""
  M_FMLA(4) = ""
  M_FMLA(5) = ""
  M_FMLA(6) = ""
  M_FMLA(7) = ""
  M_FMLA(8) = ""
  M_FMLA(9) = ""
  M_RDOF(0) = ""
  M_RDOF(1) = ""
  M_RDOF(2) = ""
  M_RDOF(3) = ""
  M_RDOF(4) = ""
  M_RDOF(5) = ""
  M_RDOF(6) = ""
  M_RDOF(7) = ""
  M_RDOF(8) = ""
  M_RDOF(9) = ""
  Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "select * from config where comp='" & compPth & "' and vtyp='SAL' AND DBCD='" & TXCD & _
          "'  AND UNIT='" & UNCD & "' order by srch", CN, adOpenKeyset, adLockPessimistic
  CNTR = 0
  Do While Not RS.EOF
   flexBTRM.Rows = flexBTRM.Rows + 1
   flexBTRM.TextMatrix(CNTR, 0) = RS!NICK & ""
   flexBTRM.TextMatrix(CNTR, 1) = Format(RS!PERC, "#######.00")
   M_OPER(CNTR) = Trim(RS!OPER)
   M_PERC(CNTR) = RS!PERC
   M_POSTCOD(CNTR) = Trim(RS!CODE)
   M_NICK(CNTR) = Trim(RS!NICK)
   M_POSTYESNO(CNTR) = Trim(RS!post)
   M_FMLA(CNTR) = Trim(RS!FMLA)
   M_RDOF(CNTR) = Trim(RS!rdof)
   RS.MoveNext
   CNTR = CNTR + 1
  Loop
  Dim TMP_FMLA(0 To 15) As String
  CNTR = 0
  For CNTR = 0 To 9
    
    M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "GROSS TOTAL", "M_STOT ")
    M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "TOTAL QUANTITY", "M_TQTY ")
    M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "TOTAL PCS", "M_TPCS ")
    If M_NICK(0) <> "" Then
        If M_OPER(0) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(0), "AMT_01 ")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(0), " -AMT_01")
        End If
    End If
    If M_NICK(1) <> "" Then
        If M_OPER(1) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(1), " +AMT_02")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(1), " -AMT_02")
        End If
    End If
    If M_NICK(2) <> "" Then
        If M_OPER(2) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(2), " +AMT_03")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(2), " -AMT_03")
        End If
    End If
    
    If M_NICK(3) <> "" Then
        If M_OPER(3) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(3), " +AMT_04")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(3), " -AMT_04")
        End If
    End If
    
    If M_NICK(4) <> "" Then
        If M_OPER(4) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(4), " +AMT_05")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(4), " -AMT_05")
        End If
    End If
    
    If M_NICK(5) <> "" Then
        If M_OPER(5) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(5), " +AMT_06")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(5), " -AMT_06")
        End If
    End If
    
    If M_NICK(6) <> "" Then
        If M_OPER(6) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(6), " +AMT_07")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(6), " -AMT_07")
        End If
    End If
    
    If M_NICK(7) <> "" Then
        If M_OPER(7) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(7), " +AMT_08")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(7), " -AMT_08")
        End If
    End If
    
    If M_NICK(8) <> "" Then
        If M_OPER(8) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(8), " +AMT_09")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(8), " -AMT_09")
        End If
    End If
  Next
  If flexBTRM.Rows > 0 Then
    'O.k
   Else
    flexBTRM.Enabled = False
  End If
End Sub

Private Sub LBLGRS_Change()
  TXTITOT = Format(LBLGRS, "#########.00")
End Sub

Private Sub TXTADDRESS_GotFocus()
TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   If TXTDLPTY = Empty And TXTDLPTY.Enabled Then TXTDLPTY.SetFocus: Exit Sub
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Or (TXTADDRESS = Empty And KeyCode = vbKeyReturn) Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & TXTDLPTY & "'", 0, Empty, "Select Consignee Address")
   End If
End Sub

Private Sub TXTADDRESS_LostFocus()
TXTADDRESS.BackColor = vbWhite
End Sub

Private Sub txtBEdit_GotFocus()
  txtBEdit.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtBEdit_LostFocus()
 txtBEdit.BackColor = vbWhite
End Sub

Private Sub TXTBRNM_GotFocus()
 TXTBRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTBRNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTBRNM.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTBRNM.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM REFMST WHERE CATA='B'", 0, TXTBRNM, "List of Agent")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTBRNM.Text = ""
            Ref_Cat = "B"
            Frm_Ref_FAS.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTBRNM = Empty
    End If
    If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
End Sub

Private Sub TXTBRNM_LostFocus()
 TXTBRNM.BackColor = vbWhite
End Sub

Private Sub TXTCOMINV_GotFocus()
TXTCOMINV.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCOMINV_LostFocus()
 TXTCOMINV.BackColor = vbWhite
End Sub

Private Sub TXTCRDS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TXTRMRK.Enabled = True
  TXTRMRK.SetFocus
End If
End Sub

Private Sub TXTCRDS_LostFocus()
 TXTCRDS.BackColor = vbWhite
End Sub

Private Sub TXTDBAC_GotFocus()
TXTDBAC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDBAC_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And (Not Trim(TXTDBAC.Text) = Empty) Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TXTCRDS_GotFocus()
  TXTCRDS.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTCRDS.SelStart = 0
  TXTCRDS.SelLength = Len(TXTCRDS)
End Sub

Private Sub TXTDBAC_Change()
  Dim RS As New ADODB.Recordset
  Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockPessimistic
  If Not RS.EOF Then
    TXTRTORTAX.Text = RS!TTYP & ""
    TXTCRDS = Val(RS!CDAY)
  End If
  
  If cmbSaleType.Text = "Commercial Tax" Then
     TXTRTORTAX.Text = "TAX INVOICE"
  ElseIf cmbSaleType.Text = "Commercial Retail" Then
     TXTRTORTAX.Text = "RETAIL INVOICE"
  End If
End Sub

Private Sub TXTDBAC_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTDBAC.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTDBAC.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM ACCMST", 0, TXTDBAC, "List of Debit A/c")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTDBAC.Text = ""
            frm_Acc.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTDBAC = Empty
    End If
    Me.KeyPreview = True
   
    Dim M_BRCD
    Dim M_TXCD
    Dim MSTDAT As New ADODB.Recordset
    Set MSTDAT = New ADODB.Recordset
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      M_BRCD = MSTDAT!BRCD & ""
      M_TXCD = MSTDAT!TXCD & ""
    End If
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM REFMST WHERE CODE='" & M_BRCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      TXTBRNM.Text = MSTDAT!NAME & ""
    End If
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM REFMST WHERE CODE='" & M_TXCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      TXTTAXNAM.Text = MSTDAT!NAME & ""
    End If
End Sub

Private Sub TXTDBAC_LostFocus()
 TXTDBAC.BackColor = vbWhite
End Sub



Private Sub TXTDLPTY_GotFocus()
 TXTDLPTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDLPTY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTDLPTY.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTDLPTY.Text = SearchList1("SELECT  DISTINCT TOP 20 CODE,NAME FROM PADDMST WHERE RECSTAT='A'", 0, TXTDLPTY, "List of Delivery Party")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "Y"
            TXTDLPTY.Text = ""
            FrmDeliveryAddress.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTDLPTY = Empty
    End If
    If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
End Sub

Private Sub TXTDLPTY_LostFocus()
  TXTDLPTY.BackColor = vbWhite
End Sub

Private Sub TXTITOT_Change()
  'TXTBNET.Text = Trim(nstr(Val(TXTITOT.Text), 10, 2))
  TXTBNET.Text = Format(FormatNumber(Val(TXTITOT.Text), 0), "##########.00")
  
  If flexBTRM.Rows > 0 Then
    flexBTRM.COL = 0
    flexBTRM.ROW = 0
  End If
  calBTRM 0
  Call calADLS
End Sub

Private Sub TXTRMRK_GotFocus()
  TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTRMRK.SelStart = 0
  TXTRMRK.SelLength = Len(TXTRMRK)
End Sub

Private Sub TXTRMRK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdSave.SetFocus
End If
End Sub

Private Sub TXTRMRK_KeyUp(KeyCode As Integer, Shift As Integer)

Dim rsNarration As Recordset

    Set rsNarration = New Recordset
    rsNarration.Open "Select * From NARRMAST Where KCod='" & KeyCode & "' And shift =" & Shift & " and Modul='SALES'", CN, adOpenDynamic, adLockOptimistic

    
    If rsNarration.EOF = False Then
        TXTRMRK = Trim(rsNarration!narr)
        TXTRMRK.SelStart = 1000
    End If
    
    rsNarration.Close
      
End Sub

Private Sub TXTRMRK_LostFocus()
 TXTRMRK.BackColor = vbWhite
End Sub

Private Sub TXTRTORTAX_GotFocus()
 TXTRTORTAX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRTORTAX_LostFocus()
 TXTRTORTAX.BackColor = vbWhite
End Sub

Private Sub TXTTAXNAM_Change()
  If TXTTAXNAM <> Empty Then
     TXCD = GetCode("TAXMST", TXTTAXNAM, "NAME", "CODE")
  Else
     TXCD = Empty
  End If
  
 Call FIL_Billingterm
 calBTRM 0
 Call calADLS
End Sub

Private Sub TXTTAXNAM_GotFocus()
  TXTTAXNAM.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTTAXNAM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or (Trim(TXTTAXNAM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTTAXNAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM TAXMST WHERE RECSTAT='A'", 0, TXTTAXNAM.Text, "SELECT TAX FROM LIST")
        If key_PressNew = True Then
            M_DESC = "": Key = "":  TXTTAXNAM.Text = ""
            FrmSaleTaxMaster.Show
        Else
            TXCD = Key
        End If
    End If
    
     If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
    
    Me.KeyPreview = True
End Sub

Private Sub TXTTAXNAM_LostFocus()
  TXTTAXNAM.BackColor = vbWhite
End Sub

Private Sub TXTTPCS_Change()
  calBTRM 0
End Sub

Private Sub TXTTQTY_Change()
  calBTRM 0
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    FLEX.COL = 1
    If FLEX.Enabled Then FLEX.SetFocus
    FLEX.CellBackColor = RGB(247, 251, 217)
  End If
End Sub

Private Sub calBTRM(ByVal ICTR As Integer)
    Dim J As Integer, iFMLA(20) As Double, subTot As Double
    Dim c_FMLA(20) As String
    Dim L As Integer
    Dim m As Integer
    Dim B() As String
    subTot = 0
    Dim a() As String, K As Integer
    J = 0
    If flexBTRM.Rows = 0 Then
      Exit Sub
    End If
    For J = flexBTRM.ROW To flexBTRM.Rows - 1
        If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then flexBTRM.TextMatrix(J, 2) = 0
    Next J

    For J = flexBTRM.ROW To flexBTRM.Rows - 1
        c_FMLA(J) = Trim(M_FMLA(J))
        If Len(c_FMLA(J)) <= 6 Then
            Select Case c_FMLA(J)
                Case "M_STOT"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTITOT.Text)) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "M_TQTY"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTTQTY.Text)), "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "M_TPCS"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTTPCS.Text)), "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_01"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(0, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_02"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(1, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_03"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(2, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_04"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(3, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_05"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(4, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_06"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(5, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_07"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(6, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_08"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(7, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_09"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(8, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
            End Select
                
            If M_RDOF(J) = "Y" Then
                flexBTRM.TextMatrix(J, 2) = Format(FormatNumber(Val(flexBTRM.TextMatrix(J, 2)), 0), "############.00")
            Else
                flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "############.00")
            End If
        Else
            c_FMLA(J) = Replace(c_FMLA(J), "M_STOT", Val(TXTITOT.Text))
            c_FMLA(J) = Replace(c_FMLA(J), "M_TQTY", Val(TXTTQTY.Text))
            c_FMLA(J) = Replace(c_FMLA(J), "M_TPCS", Val(TXTTPCS.Text))
            For K = 0 To J
                c_FMLA(J) = Replace(c_FMLA(J), "AMT_0" & K + 1, Format(flexBTRM.TextMatrix(K, 2), "##########.00"))
            Next K
            c_FMLA(J) = c_FMLA(J)
            a() = Split(c_FMLA(J), " ")
            
            Dim Y As Double
            
            Y = 0
            For K = 0 To UBound(a)
             Y = Y + Val(a(K))
            Next
                
            If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
            
              flexBTRM.TextMatrix(J, 2) = Abs(Y)
              
            End If
            
            If M_RDOF(J) = "N" Then
                If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                  flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 2)) * Val(flexBTRM.TextMatrix(J, 1))) / 100, "##########.00")
                  
                End If
            Else
                If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                  flexBTRM.TextMatrix(J, 2) = Format(FormatNumber(Val(flexBTRM.TextMatrix(J, 2)) * Val(flexBTRM.TextMatrix(J, 1)) / 100, 0), "##########.00")
                End If
            End If
            
        End If
MsubTot:
        If M_OPER(J) = "+" Then
            subTot = subTot + Val(flexBTRM.TextMatrix(J, 2))
        Else
            subTot = subTot - Val(flexBTRM.TextMatrix(J, 2))
        End If
        'TXTBNET.Text = Val(TXTITOT.Text) + subTot
        TXTBNET.Text = Format(FormatNumber(Val(TXTITOT.Text) + subTot, 0), "##########.00")
    Next J

    
End Sub

Private Sub EditKeyCode(MSHFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
   
    Dim ANS As String
    chgFlag = True
    'Standard edit control processing.
   Select Case KeyCode
    
   Case 27   ' ESC: hide, return focus to MSHFlexGrid.
      Edt.Visible = False
      MSHFlexGrid.SetFocus
    
   Case 9    ' TAB return focus to mshflexgrid.
        If FLEX.COL - 1 <> 7 And FLEX.COL - 1 <> 0 Then FLEX.TextMatrix(FLEX.ROW, FLEX.COL - 1) = 0
   Case 13    ' ENTER return focus to MSHFlexGrid.
         MSHFlexGrid.SetFocus
         If MSHFlexGrid.COL = 2 Then
            If MSHFlexGrid.ROW < MSHFlexGrid.Rows - 1 Then
               MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
               MSHFlexGrid.COL = 1
            End If
         Else
            MSHFlexGrid.COL = MSHFlexGrid.COL + 1
        End If
   Case 38      ' Up.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.ROW > MSHFlexGrid.FixedRows Then
         MSHFlexGrid.ROW = MSHFlexGrid.ROW - 1
      End If

   Case 40      ' Down.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.ROW < MSHFlexGrid.Rows - 1 Then
         MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
      End If
   End Select
   chgFlag = False
End Sub

Private Sub MSHFlexGridEdit(MSHFlexGrid As Control, Edt As Control, KeyAscii As Integer)
    chgFlag = True
    ' Use the character that was typed.
   Select Case KeyAscii
   ' A space means edit the current text.
   Case 0 To 12
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 14 To 26
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 13
      If MSHFlexGrid.COL = 2 Then
            If MSHFlexGrid.Rows <> MSHFlexGrid.ROW + 1 Then
                MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
            Else
                TXTCRDS.SetFocus
            End If
            MSHFlexGrid.COL = 1
            Exit Sub
        Else
            
            MSHFlexGrid.COL = MSHFlexGrid.COL + 1
            Exit Sub
      End If
   Case 28 To 32
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 27
        Edt.Text = Empty
        Exit Sub
   ' Anything else means replace the current text.
   Case Else
      Edt = Chr(KeyAscii)
      Edt.SelStart = 1
   End Select

   ' Show Edt at the right place.
   Edt.MOVE MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
      MSHFlexGrid.CellWidth - 8, _
      MSHFlexGrid.CellHeight - 8
   Edt.Visible = True

   ' And make it work.
   Edt.SetFocus
   chgFlag = False
End Sub
Private Sub calADLS()
    Dim P As Integer
    TXTADLS.Text = Empty
    For P = 0 To flexBTRM.Rows - 1
        If M_OPER(P) = "-" Then
            TXTADLS.Text = Format(Val(TXTADLS.Text) - Val(flexBTRM.TextMatrix(P, 2)), "############.00")
        Else
            TXTADLS.Text = Format(Val(TXTADLS.Text) + Val(flexBTRM.TextMatrix(P, 2)), "############.00")
        End If
    Next P
End Sub
Private Sub txtBEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode flexBTRM, txtBEdit, KeyCode, Shift
End Sub
Private Sub txtBEdit_KeyPress(KeyAscii As Integer)
   If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub

Private Function CHKSAVEDATA() As Boolean
  Dim CHKRS As New ADODB.Recordset
  Set CHKRS = New ADODB.Recordset
  
  
     
  'Debit A/c Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE from ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Debit A/c Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  'Delivery Party Name
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE from PADDMST WHERE NAME='" & TXTDLPTY.Text & "' AND RECSTAT='A'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Delivery Party Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  'Agent Name
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE from REFMST WHERE NAME='" & TXTBRNM.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Agent Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  'Sale Tax Catagoery
  If TXTTAXNAM.Enabled = True Then
    If CHKRS.State = 1 Then CHKRS.Close
    CHKRS.Open "SELECT CODE FROM TAXMST WHERE NAME='" & TXTTAXNAM.Text & "'", CN, adOpenKeyset, adLockPessimistic
    If CHKRS.EOF Then
       MsgBox "Tax Catagoery Name Not Define ", vbCritical
       CHKSAVEDATA = False
       Exit Function
    End If
  End If
  
  'Retail / Tax Catagoery
  If Trim(TXTRTORTAX) = Empty Then
     MsgBox "Retail/Tax Invoice Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  
  CHKSAVEDATA = True
End Function

Private Sub cmdSave_Click()
  
  On Error GoTo LAST
  If CHKSAVEDATA = False Then
    Exit Sub
  End If
  
  Call CHKFLEX
   
  If SAVEFLAG = True Then
    
  Call FindSerial
   
  Dim SAVDAT As ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
     MsgBox "Duplicate Bill No.Make Change in VOUCHER TYPE for Generation"
     cmdSave.SetFocus
     Exit Sub
  End If
  End If
  
  If SAVEFLAG = True And UNT_MMS_INSTALL = "Y" Then
     If Not IsStockSupport Then
        Exit Sub
     End If
  End If
  
  Call SAVERECSAL
  
  If SAVEFLAG = True Then
     MsgBox "Your Invoice No. is " + TXTVBNO.Text
  End If
  Call cmdCancel_Click
  If zoomflag = True Then
    Call CMDEXIT_Click
    Exit Sub
  End If
  
  Exit Sub
LAST:
  MsgBox ERR.Description
End Sub

Private Sub SAVERECSAL()
 
 On Error GoTo LAST
  Dim SQL As String
  
  Dim SAVDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Dim M_CRAC As String
  Dim M_DRAC As String
  Dim M_PCOD As String
  Dim M_DCOD As String
  Dim M_ADDR As String
  Dim M_CPCD As String
  Dim M_ARCD As String
  Dim M_TRCD As String
  Dim M_TXCD As String
  Dim M_BRCD As String
  Dim i As Double
  Dim J As Double
  Set SAVDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  
  'Credit A/c
  M_CRAC = "XXXXXX"
  
  'Debit A/c
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
    M_DRAC = SAVDAT!CODE & ""
    M_PCOD = SAVDAT!CODE & ""
    M_CPCD = SAVDAT!CPCD & ""
    M_ARCD = SAVDAT!ARCD & ""
  End If
  
  'Consignee Name and Address
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM PADDMST WHERE NAME='" & TXTDLPTY.Text & "' AND ADDR='" & TXTADDRESS & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then M_DCOD = SAVDAT!CODE & "": M_ADDR = SAVDAT!SRNO & ""
  
  'Agent
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM REFMST WHERE NAME='" & TXTBRNM.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then M_BRCD = SAVDAT!CODE & ""
  
  'Tax Catagoery
  If Not Trim(TXTTAXNAM.Text) = Empty Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT CODE FROM TAXMST WHERE NAME='" & TXTTAXNAM.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not SAVDAT.EOF Then M_TXCD = SAVDAT!CODE & ""
  End If
  
  CN.BeginTrans
  Call DELETESAL
  
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & _
              "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "SAL"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!dbcd = M_DBCD
  SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
  SAVDAT!VBNO = Trim(TXTVBNO.Text)
  SAVDAT!CVBN = Trim(TXTCOMINV.Text)
  SAVDAT!CRAC = M_CRAC
  SAVDAT!DRAC = M_DRAC
  SAVDAT!PCOD = M_PCOD
  SAVDAT!DCOD = M_DCOD
  SAVDAT!ADDRESS = M_ADDR
  SAVDAT!BRCD = M_BRCD
  SAVDAT!CPCD = M_CPCD
  SAVDAT!ARCD = M_ARCD
  SAVDAT!TXCD = M_TXCD
  SAVDAT!TAXGRP = GetCode("TAXMST", M_TXCD & "", "CODE", "GRPCOD")
  SAVDAT!TPCS = Val(TXTTPCS.Text)
  SAVDAT!TQTY = Val(TXTTQTY.Text)
  SAVDAT!ITOT = Val(TXTITOT.Text)
  SAVDAT!BADJ = Val(TXTBNET.Text) - Val(TXTITOT.Text)
  SAVDAT!BNET = Val(TXTBNET.Text)
  SAVDAT!TTYP = Trim(TXTRTORTAX.Text)
  SAVDAT!CDAY = Val(TXTCRDS)
  If SAVEFLAG = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  SAVDAT![User] = cUName & ""
  SAVDAT!DVCD = DIVCOD
  SAVDAT!unit = UNCD
  SAVDAT!TRCD = M_TRCD
  SAVDAT!RECSTAT = "A"
  SAVDAT!BRMK = Trim(TXTRMRK.Text)
  SAVDAT!EXTRA5 = "BOX"
  
  i = 0
  For i = 0 To flexBTRM.Rows - 1
    J = 0
    For J = 0 To SAVDAT.Fields.COUNT - 1
      If Trim(SAVDAT.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
        SAVDAT.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 2))
      End If
      If Trim(SAVDAT.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
        SAVDAT.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 1))
      End If
    Next
  Next
  
  Dim K As Double
  K = 1
  SAVDAT.Update
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  i = 1
  For i = 1 To FLEX.Rows - 1
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "SAL"
    SAVDAT!SRNO = ""
    SAVDAT!SRCH = i
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = TXTVBNO.Text
    SAVDAT!CHDT = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!CRAC = M_CRAC
    SAVDAT!DRAC = M_DRAC
    SAVDAT!PCOD = M_PCOD
    SAVDAT!DCOD = M_DCOD
    SAVDAT!ICOD = GetCode("ITMMST", FLEX.TextMatrix(i, 1), "NAME", "CODE")
    SAVDAT!PCES = Val(FLEX.TextMatrix(i, 2))
    SAVDAT!QNTY = Val(FLEX.TextMatrix(i, 3))
    SAVDAT!RATE = Val(FLEX.TextMatrix(i, 4))
    SAVDAT!AMNT = Val(FLEX.TextMatrix(i, 6))
    SAVDAT!QORP = Mid(FLEX.TextMatrix(i, 5), 1, 1)
    SAVDAT!GP_REMARKS = Trim(FLEX.TextMatrix(i, 7))
    SAVDAT![User] = cUName
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    
    If Trim(SAVDAT!ICOD & "") <> Empty Then
       SAVDAT!OPER = "-"
    Else
       SAVDAT!OPER = "*"
    End If
    SAVDAT!DVCD = DIVCOD
    SAVDAT!unit = UNCD
    SAVDAT!RTYP = "SAL"
    SAVDAT!RSRN = ""
    SAVDAT!RSRC = i
    SAVDAT!SDBC = ""
    SAVDAT!SVBN = TXTVBNO
    SAVDAT!TXCD = M_TXCD
    SAVDAT!RTCD = GetCode("TAXMST", M_TXCD, "CODE", "RATE_CODE")
    SAVDAT!RECSTAT = "A"
    SAVDAT.Update
  Next
   
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  i = 1
  For i = 1 To FLEX.Rows - 1
    
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "SAL"
    SAVDAT!SRNO = ""
    SAVDAT!SRCH = i
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = TXTVBNO.Text
    SAVDAT!CHDT = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!CRAC = M_CRAC
    SAVDAT!DRAC = M_DRAC
    SAVDAT!PCOD = M_PCOD
    SAVDAT!DCOD = M_DCOD
    SAVDAT!ICOD = GetCode("ITMMST", FLEX.TextMatrix(i, 1), "NAME", "CODE")
    SAVDAT!PCES = Val(FLEX.TextMatrix(i, 2))
    SAVDAT!QNTY = Val(FLEX.TextMatrix(i, 3))
    SAVDAT!RATE = FindFIFORate(FLEX.TextMatrix(i, 1), Val(FLEX.TextMatrix(i, 3)))
    SAVDAT!AMNT = Val(SAVDAT!QNTY) * Val(SAVDAT!RATE)
    SAVDAT!QORP = Mid(FLEX.TextMatrix(i, 5), 1, 1)
    SAVDAT![User] = cUName
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    
    If Trim(SAVDAT!ICOD & "") <> Empty Then
       SAVDAT!OPER = "-"
    Else
       SAVDAT!OPER = "*"
    End If
    
    SAVDAT!DVCD = DIVCOD
    SAVDAT!unit = UNCD
    SAVDAT!grad = ""
    SAVDAT!ltno = ""
    SAVDAT!COPS = 0
    SAVDAT!TWST = ""
    SAVDAT!RTYP = "SAL"
    SAVDAT!RSRN = ""
    SAVDAT!RSRC = i
    
    SAVDAT!RECSTAT = "A"
    SAVDAT!EXTRA1 = "BOX"
    SAVDAT!EXTRA5 = Left(FLEX.TextMatrix(i, 9), 249)
    SAVDAT.Update
  Next
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  'Dim SQL As String
  
  Call SetRollBalQty
  
  If MSTDAT.State = 1 Then MSTDAT.Close
  MSTDAT.Open "SELECT ISNULL(SUM(PCES),0) AS TPCS,ISNULL(SUM(QNTY),0) AS TQTY FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If Not MSTDAT.EOF Then
    CN.Execute "UPDATE BILLMAIN SET TPCS='" & MSTDAT!TPCS & "', TQTY='" & MSTDAT!TQTY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
    CN.Execute "UPDATE BILLMAIN SET EXTRA2='" & FLEX.TextMatrix(1, 8) & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  End If

  'update last bill no in DAYBOK
  If SAVEFLAG = True Then
    SQL = "UPDATE SERIALMASTER SET [SRNO]='" & TXTVBNO & "',LEDT = '" & Format(TXTVBDT, "YYYY/MM/DD") & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
          "' AND VTYP='SAL' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "'"
    CN.Execute SQL
    
  End If
  
  'Update spmain for ramt,dbna,crna,retg
  Dim REC_AMT As Double
  Dim DBN_AMT As Double
  Dim CRN_AMT As Double
  Dim RET_AMT As Double
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT ISNULL(SUM(RAMT),0) AS RAMT, ISNULL(SUM(DBNA),0) AS DBNA, ISNULL(SUM(CRNA),0) AS CRNA, ISNULL(SUM(RETG),0) AS RETG FROM RPTRAN WHERE COMP='" & compPth & "' AND BSR1='SAL' AND BSR2=''", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
    CN.Execute "UPDATE BILLMAIN SET RAMT='" & SAVDAT!RAMT & "',DBNA='" & SAVDAT!DBNA & "',CRNA='" & SAVDAT!CRNA & "',RETG='" & SAVDAT!RETG & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  End If
  
   If SAVEFLAG = True Then
      Call SetFIFOConsumption
   End If
   
   'Save Data IN EGPMan for VAT REports
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!unit = UNCD
  SAVDAT!VTYP = "SAL"
  SAVDAT!SRNO = TXTVBNO
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
  SAVDAT!dbcd = M_DBCD
  SAVDAT!CRAC = "XXXXXX"
  SAVDAT!DRAC = M_PCOD
  SAVDAT!VBNO = Trim(TXTVBNO)
  SAVDAT!chln = Trim(TXTVBNO)
  SAVDAT!CHDT = Format(TXTVBDT, "YYYY/MM/DD")
  SAVDAT!DCOD = M_DCOD
  SAVDAT!ADDRESS = M_ADDR
  SAVDAT!BRCD = M_BRCD
  SAVDAT!CPCD = M_CPCD
  SAVDAT!ARCD = M_ARCD
  SAVDAT!TXCD = M_TXCD
  SAVDAT!TAXGRP = GetCode("TAXMST", M_TXCD & "", "CODE", "GRPCOD")
  SAVDAT!ICOD = GetCode("ITMMST", FLEX.TextMatrix(i, 1), "NAME", "CODE")
  SAVDAT!PCES = Val(TXTTPCS)
  SAVDAT!QNTY = Val(TXTTQTY)
  SAVDAT!AMNT = Val(TXTITOT)
  SAVDAT!BADJ = Val(TXTBNET) - Val(TXTITOT)
  SAVDAT!BNET = Val(TXTBNET)
  SAVDAT!TTYP = ""
  SAVDAT!RORT = TXTRTORTAX
  SAVDAT!RECSTAT = "A"
  i = 0
  For i = 0 To flexBTRM.Rows - 1
    J = 0
    For J = 0 To SAVDAT.Fields.COUNT - 1
      If Trim(SAVDAT.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
        SAVDAT.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 2))
      End If
      If Trim(SAVDAT.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
        SAVDAT.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 1))
      End If
    Next
  Next
  SAVDAT.Update

  'Call UPDATESTATUS
  '---------------------------
  'DAILYSTATUS ENTRY
  If SAVEFLAG = True Then
  Call DAILYSTATUS("SAL", M_PCOD, M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "N", Now, TXTVBDT)
  Else
  Call DAILYSTATUS("SAL", M_PCOD, M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "M", Now, TXTVBDT)
  End If

  
  CN.CommitTrans
  Exit Sub
  
LAST:
 MsgBox ERR.Description
 If SAVDAT.State = 1 Then
   SAVDAT.CancelUpdate
   SAVDAT.Close
 End If
 CN.RollbackTrans
End Sub

Private Sub DELETESAL()
  Dim SAVDAT As New ADODB.Recordset
  Dim m_rtyp As String
  Dim m_rsrn As String
  Set SAVDAT = New ADODB.Recordset
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND DVCD='" & DIVCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & _
              "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  Do While Not SAVDAT.EOF
    m_rtyp = SAVDAT!RTYP & ""
    m_rsrn = SAVDAT!RSRN & ""
    
    CN.Execute "UPDATE SPTRAN SET RTYP=NULL, RSRN=NULL WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND VTYP='" & m_rtyp & "' AND RSRN='" & m_rsrn & "'"
               
    SAVDAT.MoveNext
  Loop
  
  CN.Execute "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & _
             "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & _
             "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
             
  CN.Execute "DELETE FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & _
             "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
             
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & M_DBCD & _
             "' AND VTYP='SAL' AND VBNO='" & TXTVBNO & "'"
             
  CN.Execute "DELETE FROM TRDBOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND GRNNO='" & TXTVBNO & "' AND RECSTAT <> 'D'"
  
End Sub

Private Sub TXTVBNO_GotFocus()
  TXTVBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTVBNO_LostFocus()
  TXTVBNO.BackColor = vbWhite
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)

 If FLEX.COL = 3 And SAVEFLAG = False And KeyAscii <> 13 And UCase(UNT_MMS_INSTALL) = "Y" Then
    Exit Sub
 End If

  On Error GoTo LAST
  FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.ROW
  Dim ALLOW_KEY As Boolean
  Dim FWD_COL As Boolean
  Dim ENTER_PRESS As Boolean
  Dim MSTDAT As New ADODB.Recordset
  
  'VFLQTY = 0
  Set MSTDAT = New ADODB.Recordset
  FWD_COL = False
  ALLOW_KEY = False
  If FLEX.COL = 6 Or FLEX.COL = 7 Or FLEX.COL = 8 Or FLEX.COL = 9 Or FLEX.COL = 11 Then
    If Not FLEX.COL = 8 Then
      If InStr(1, FLEX.TextMatrix(FLEX.ROW, FLEX.COL), ".") > 0 And KeyAscii = 46 Then
        KeyAscii = 0
        Exit Sub
      End If
    End If
  End If
  
  If Emptycell = True And (Not KeyAscii = 13) Then
    If FLEX.COL <> 7 Then
       FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty
    End If
    Emptycell = False
  End If
  
  Select Case FLEX.COL
   Case 1
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    If Chr(KeyAscii) = vbKeyF2 Or FLEX.TextMatrix(FLEX.ROW, 1) = Empty Then
      FLEX.TextMatrix(FLEX.ROW, 1) = SearchList1("select  TOP 20 code,name from itmmst", 0, FLEX.TextMatrix(FLEX.ROW, 1), "SELECT ITEM FROM LIST")
    End If
    If key_PressNew = True Then
       M_DESC = ""
       Key = ""
       FLEX.TextMatrix(FLEX.ROW, 3) = ""
       frm_Item.Show
    End If
    
    'BECAUSE OF STOCK EFFECT OF MMS
    If Chr(KeyAscii) <> vbKeyF2 Or KeyAscii = 13 Then
      ALLOW_KEY = False
    Else
      ALLOW_KEY = True
    End If
    
   Case 2
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 3, 4
       
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 5
    If FLEX.COL = 5 Then
        If KeyAscii = 13 Then
            Exit Sub
        End If
        FLEX.TextMatrix(FLEX.ROW, 5) = Empty
    End If
   
    If UCase(Chr(KeyAscii)) = "Q" Or UCase(Chr(KeyAscii)) = "P" Or UCase(Chr(KeyAscii)) = "X" Then
      ALLOW_KEY = True
     Else
      ALLOW_KEY = False
    End If
   Case Else
        ALLOW_KEY = True
  End Select
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
   Else
    ENTER_PRESS = False
  End If
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))
    If lnth > 0 Then
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Mid(FLEX.TextMatrix(FLEX.ROW, FLEX.COL), 1, lnth - 1)
      Call CalculatePcsMtr
      Exit Sub
    End If
  End If
  
  If ALLOW_KEY = False Then
    If ENTER_PRESS = True Then
     Else
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
    If ALLOW_KEY = True Then
    If ENTER_PRESS = False Then
        If FLEX.COL = 5 Then
            FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = FLEX.TextMatrix(FLEX.ROW, FLEX.COL) + UCase(Chr(KeyAscii))
        Else
            FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = FLEX.TextMatrix(FLEX.ROW, FLEX.COL) + Chr(KeyAscii)
        End If
       Call CalculatePcsMtr
    End If
    End If
  
    FWD_COL = False
    If ENTER_PRESS = True Then
    Select Case FLEX.COL
     Case 1
       FWD_COL = True
     Case 2
      FWD_COL = True
     Case 3
      FWD_COL = True
     Case 4
      FWD_COL = True
     Case 5
      If FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = "Q" Or FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = "P" Or FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = "X" Then
        FWD_COL = True
      End If
     Case 6
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
      Else
        FWD_COL = False
      End If
     Case 7
           FWD_COL = True
     End Select
    
    
      If FWD_COL = True Then
      
      If FLEX.COL = 7 Then
        'Allowed to add row with msgbox
        'Check all the cell are filled
        
        Call CHKFLEX
        If Not CHK_FLX Then
          MsgBox "Invalid Data in item details "
          FLEX.ROW = FLXROW
          FLEX.COL = FLXCOL
          FLEX.SetFocus
          Exit Sub
        End If
        
        If SAVEFLAG = False And UNT_MMS_INSTALL = "Y" Then
           Exit Sub
        End If
        
        Dim AYS
        AYS = MsgBox("Want to Add More Item ", vbYesNo)
       ' If AYS = vbYes Then
       '   FLEX.Rows = FLEX.Rows + 1
       '   FLEX.ROW = FLEX.Rows - 1
       '   FLEX.COL = 1
       '  Else
       '   If txtDEPT.Enabled = True Then txtDEPT.SetFocus: Exit Sub
       '   If TXTCRDS.Enabled = True Then TXTCRDS.SetFocus: Exit Sub
       ' End If
        
       Else
        If FLEX.COL = 5 Then
         If Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "Q" Then
           FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 3)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
         ElseIf Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "P" Then
           FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 2)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
         End If
        End If
        FLEX.COL = FLEX.COL + 1
        End If
      End If
        Emptycell = True
    End If
  Exit Sub
LAST:
  MsgBox ERR.Description
  FLEX.SetFocus
  Exit Sub
End Sub

Private Sub CHKFLEX()
  CHK_FLX = True
  Dim FLXR As Double
  
  For FLXR = 1 To FLEX.Rows - 1
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 2)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 2
       Exit For
    End If
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 3)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 3
       Exit For
    End If
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 4)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 4
       Exit For
    End If
  Next
End Sub

Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "SAL"
  DLYSTA!PCOD = TXTDBAC.Text
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = Val(TXTBNET)
  DLYSTA!CUSR = cUName
  If SAVEFLAG = True Then
    DLYSTA!ACTN = "E"
   Else
    DLYSTA!ACTN = "M"
  End If
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "SAL"
  DLYSTA!PCOD = TXTDBAC.Text
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = Val(TXTBNET)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub FindSerial()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND NAME='" & cmbSaleType.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not PKTYPRS.EOF Then
 M_DBCD = Trim(PKTYPRS!CODE & "")
End If

If M_DBCD <> Empty Then
 TXTVBNO = GenVNO("SAL", M_DBCD)
End If

End Sub

Private Sub SetSaleType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND FYCD='" & FYCD & "' AND ACTIVE='Y'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbSaleType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop

If cmbSaleType.ListCount >= 1 Then cmbSaleType.ListIndex = 0

End Sub


Private Sub CalculatePcsMtr()
Dim i As Long
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  For i = 1 To FLEX.Rows - 1
     If Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "Q" Then
        FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 3)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
     ElseIf Mid(FLEX.TextMatrix(FLEX.ROW, 5), 1, 1) = "P" Then
        FLEX.TextMatrix(FLEX.ROW, 6) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 2)) * Val(FLEX.TextMatrix(FLEX.ROW, 4)), "#########.00")
     End If
    
     TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 2)), "######")
     TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 3)), "########.000")
     TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 6)), "########.00")
  Next
End Sub



'FIFO----------------------
Private Sub SetFIFOConsumption()
On Error GoTo FIFOERR

'VARIABLE DECLARATION
Dim ICOD As String, Item As String, INDEX As Long
Dim BALQNTY As Double, TMPQTY As Double
Dim FIFORS As ADODB.Recordset: Set FIFORS = New ADODB.Recordset

'-------------------------------------------------------------
'-------------------------------------------------------------
For INDEX = 1 To FLEX.Rows - 1
'-------------------
'INITIALISE
Item = FLEX.TextMatrix(INDEX, 1)
BALQNTY = Val(FLEX.TextMatrix(INDEX, 3))
'-------------------

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then ICOD = Trim(FIFORS!CODE & "")
FIFORS.Close

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then
Do While Not FIFORS.EOF
        
        TMPQTY = Val(FIFORS!BAL_QNTY)  'Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
            
        If BALQNTY > TMPQTY Then
           FIFORS!ISS_QNTY = Val(FIFORS!ISS_QNTY) + TMPQTY
           FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
           FIFORS!LAST_ISS_DT = Format(TXTVBDT, "YYYY/MM/DD")
           BALQNTY = BALQNTY - TMPQTY
           FIFORS.Update
        ElseIf BALQNTY > 0 Or BALQNTY = TMPQTY Then
           FIFORS!ISS_QNTY = Val(FIFORS!ISS_QNTY) + BALQNTY
           FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
           FIFORS!LAST_ISS_DT = Format(TXTVBDT, "YYYY/MM/DD")
           FIFORS.Update
           BALQNTY = 0
           Exit Do
        End If
                
FIFORS.MoveNext
Loop
End If
Next INDEX

Exit Sub
FIFOERR:
MsgBox ERR.Description
End Sub

Private Function IsStockSupport() As Boolean
On Error GoTo ERR
Dim ITMCOD As String
Dim TOTISSQTY As Double, TOTRECQTY As Double, TXTSTK As Double
Dim chkitm As ADODB.Recordset
Set chkitm = New ADODB.Recordset

IsStockSupport = True

Dim K As Long
For K = 1 To FLEX.Rows - 1
    
    If chkitm.State = 1 Then chkitm.Close
    chkitm.Open "SELECT CODE FROM ITMMST WHERE NAME='" & FLEX.TextMatrix(K, 1) & "'", CN, adOpenDynamic, adLockOptimistic
    If Not chkitm.EOF Then
       ITMCOD = Trim(chkitm!CODE & "")
    Else
        chkitm.Close
        Exit Function
    End If

   'Consumption department wise
    If chkitm.State = 1 Then chkitm.Close
    
    chkitm.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                "' AND OPER='+' AND RECSTAT<>'D' AND ICOD='" & ITMCOD & "'", CN, adOpenDynamic, adLockOptimistic
    
    If Not chkitm.EOF Then
        TOTRECQTY = chkitm!QNTY
    Else
        TOTRECQTY = 0
    End If
    
    If chkitm.State = 1 Then chkitm.Close
   
    chkitm.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                "' AND OPER='-' AND RECSTAT<>'D'  AND ICOD='" & ITMCOD & "'", CN, adOpenDynamic, adLockOptimistic
                
    If Not chkitm.EOF Then
        TOTISSQTY = chkitm!QNTY
    Else
        TOTISSQTY = 0
    End If
    
    TXTSTK = Format(Round(TOTRECQTY, 3) - Round(TOTISSQTY, 3), "#######.000")
    
    If Val(FLEX.TextMatrix(K, 3)) > TXTSTK Then
       IsStockSupport = False
       MsgBox "Item Stock Not Supported"
       FLEX.COL = 1
       FLEX.ROW = K
       FLEX.SetFocus
       Exit Function
    End If

Next K

Exit Function
ERR:
MsgBox ERR.Description
CN.RollbackTrans
End Function

'FIFO
Private Function FindFIFORate(Item As String, QNTY As Double) As Double
On Error GoTo FIFOERR
Dim ICOD As String
Dim Top As Double
Dim Bottom As Double
Dim BALQNTY As Double
Dim FIFORS As ADODB.Recordset
Set FIFORS = New ADODB.Recordset

FindFIFORate = 0
BALQNTY = QNTY

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then ICOD = Trim(FIFORS!CODE & "")
FIFORS.Close

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT BAL_QNTY AS QNTY,RATE,NETRATE FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then
 If QNTY <= Val(FIFORS!QNTY) Then FindFIFORate = Val(FIFORS!RATE): Exit Function
End If

Do While Not FIFORS.EOF
       
   If BALQNTY >= Val(FIFORS!QNTY) Then
      Top = Top + (Val(FIFORS!QNTY) * Val(FIFORS!RATE))
      Bottom = Bottom + Val(FIFORS!QNTY)
      BALQNTY = BALQNTY - Val(FIFORS!QNTY)
   Else
      Top = Top + (BALQNTY * Val(FIFORS!RATE))
      Bottom = Bottom + BALQNTY
      BALQNTY = 0
      Exit Do
   End If
     
   
FIFORS.MoveNext
Loop
FIFORS.Close

If Top > 0 And Bottom > 0 Then
  FindFIFORate = Top / Bottom
Else
  FindFIFORate = 0
End If

Exit Function
FIFOERR:
MsgBox ERR.Description
End Function

Public Sub FillList(ROW As Long)
On Error GoTo LAST
Dim ITNM As String
Dim SQL As String, AddSQL As String, ROLLSQL As String
Dim Item As ListItem
Dim INDEX As Integer: INDEX = 1

If SAVEFLAG = True Then

ITNM = FLEX.TextMatrix(ROW, 1)

SQL = "SELECT ITMMST.NAME AS ITEM,VBNO,ICOD,NTWGT AS BALQTY,GRSWGT,TRWGT,TRDBOXREGISTER.RATE FROM TRDBOXREGISTER INNER JOIN ITMMST ON ITMMST.CODE = TRDBOXREGISTER.ICOD " & _
      " AND ITMMST.NAME = '" & ITNM & "' WHERE TRDBOXREGISTER.COMP = '" & compPth & "' AND TRDBOXREGISTER.UNIT = '" & UNCD & "' AND TRDBOXREGISTER.DVCD = '000001'" & _
      " AND RECSTAT <> 'D' AND OPER = '+' AND RVTYP IS NULL AND RVBNO IS NULL "
  
  
lstRolls.ListItems.Clear
'If Trim(FLEX.TextMatrix(ROW, 9)) <> "" Then
   
   ROLLSQL = SQL
   
   If RS.State = 1 Then RS.Close
   RS.Open ROLLSQL, CN, adOpenDynamic, adLockOptimistic
   Do While Not RS.EOF
       Set Item = lstRolls.ListItems.ADD
       lstRolls.ListItems(INDEX).Checked = False
       Item.Text = Trim(RS!VBNO)
       Item.SubItems(1) = Trim(RS!BALQTY)
       Item.SubItems(2) = ROW
       Item.SubItems(3) = nstr(RS!GRSWGT, 10, 3)
       Item.SubItems(4) = nstr(RS!TRWGT, 10, 3)
       Item.SubItems(5) = nstr(RS!RATE, 9, 2)
       Item.SubItems(6) = Trim(RS!ICOD)
       INDEX = INDEX + 1
   RS.MoveNext
   Loop
   
 '  End If
   End If
   If SAVEFLAG = False Then
   
   SQL = "SELECT ITMMST.NAME AS ITEM,VBNO,NTWGT AS BALQTY,GRSWGT,TRWGT,TRDBOXREGISTER.RATE,TRDBOXREGISTER.ICOD FROM TRDBOXREGISTER INNER JOIN ITMMST ON ITMMST.CODE = TRDBOXREGISTER.ICOD " & _
      " AND ITMMST.NAME = '" & FLEX.TextMatrix(ROW, 1) & "' WHERE TRDBOXREGISTER.COMP = '" & compPth & "' AND TRDBOXREGISTER.UNIT = '" & UNCD & "' AND TRDBOXREGISTER.DVCD = '000001'" & _
      " AND RECSTAT <> 'D' AND VTYP = 'SAL' AND OPER = '-' AND RVTYP IS NULL AND RVBNO IS NULL "
   
   SQL = SQL & " AND VBNO IN (" & Trim(FLEX.TextMatrix(ROW, 9)) & ") "
   lstRolls.ListItems.Clear
   
If RS.State = 1 Then RS.Close
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic

Do While Not RS.EOF
   Set Item = lstRolls.ListItems.ADD
   lstRolls.ListItems(INDEX).Checked = True
   Item.Text = Trim(RS!VBNO)
   Item.SubItems(1) = Trim(RS!BALQTY)
   Item.SubItems(2) = ROW
   Item.SubItems(3) = Trim(RS!GRSWGT)
   Item.SubItems(4) = Trim(RS!TRWGT)
   Item.SubItems(5) = nstr(RS!RATE, 9, 2)
   Item.SubItems(6) = Trim(RS!ICOD)
   INDEX = INDEX + 1
RS.MoveNext
Loop
End If


Call CalculateRolls
Exit Sub
LAST:
MsgBox ERR.Description
Resume
End Sub

Public Sub CalculateRolls()
txtTotalPcs.Text = Val("0")
txtTotalQty.Text = Val("0")

Dim COUNT As Integer
Dim INDEX As Long
Dim ROW As Long
Dim ROLLNO As String

For INDEX = 1 To lstRolls.ListItems.COUNT
   If lstRolls.ListItems(INDEX).Checked = True Then
     COUNT = COUNT + 1
     ROW = Val(lstRolls.ListItems(INDEX).ListSubItems(2))
     txtTotalQty.Text = CStr(Val(txtTotalQty.Text) + Val(lstRolls.ListItems(INDEX).ListSubItems(1)))
     If ROLLNO <> Empty Then ROLLNO = ROLLNO & ","
     ROLLNO = ROLLNO & "'" & Trim(lstRolls.ListItems(INDEX)) & "'"
   End If
Next INDEX
txtTotalPcs.Text = COUNT
txtTotalQty.Text = Format(txtTotalQty, "########.000")
If ROW <> 0 Then
FLEX.TextMatrix(ROW, 2) = txtTotalPcs.Text
FLEX.TextMatrix(ROW, 3) = txtTotalQty.Text
FLEX.TextMatrix(ROW, 9) = ROLLNO
End If
End Sub

Private Sub lstRolls_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   SendKeys "{DOWN}"
   Call CalculateRolls
End Sub



Private Sub SetRollBalQty()
On Error GoTo SETERR
Dim INDEX As Long
Dim ROLDET As String
Dim ITNAME As String
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset

   
    Dim Item As ListItem
    Set Item = lstRolls.ListItems.ADD
    Dim i As Long
    
    For i = 1 To lstRolls.ListItems.COUNT
   
    CN.Execute "UPDATE TRDBOXREGISTER SET RVTYP = NULL,RVBDT = NULL,RVBNO = NULL,RDBC = NULL  WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & _
                  "' AND DVCD = '000001' AND OPER = '+' AND RVTYP = 'SAL' AND GRNNO = '" & Trim(TXTVBNO) & "' AND VBNO = '" & lstRolls.ListItems(i).Text & "' "

    
    CN.Execute "DELETE FROM TRDBOXREGISTER WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & "' AND DVCD = '000001' AND VTYP = 'SAL'" & _
               " AND VBNO = '" & Trim(lstRolls.ListItems(i).Text) & "' AND OPER = '-'  AND GRNNO = '" & Trim(TXTVBNO) & "' AND RVTYP IS NULL AND RVBNO IS NULL "
    
    Next
    
    For i = 1 To lstRolls.ListItems.COUNT
    If lstRolls.ListItems(i).Checked = True Then
    
    CN.Execute "UPDATE TRDBOXREGISTER SET RVTYP = 'SAL' ,RVBDT = '" & Format(TXTVBDT, "YYYY/MM/DD") & "' ,RVBNO = '" & Trim(lstRolls.ListItems(i).Text) & "' , RDBC = '" & M_DBCD & "' WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & _
               "' AND DVCD = '000001' AND OPER = '+' AND  RVTYP IS NULL AND RVBNO IS NULL AND VBNO  = '" & Trim(lstRolls.ListItems(i).Text) & "' "
                  
    CN.Execute "INSERT INTO TRDBOXREGISTER(COMP,UNIT,DVCD,DBCD,GRNNO,VTYP,VBNO,VBDT,PCOD,ICOD,GRSWGT,TRWGT,NTWGT,RECSTAT,OPER,RATE)VALUES('" & compPth & _
                  "','" & UNCD & "','" & DIVCOD & "','" & M_DBCD & "','" & Trim(TXTVBNO) & "','SAL' , '" & Trim(lstRolls.ListItems(i).Text) & "', '" & Format(TXTVBDT, "YYYY/MM/DD") & _
                  "','" & GetCode("ACCMST", TXTDBAC.Text, "NAME", "CODE") & "','" & Trim(lstRolls.ListItems(i).SubItems(6)) & "','" & Val(lstRolls.ListItems(i).SubItems(3)) & _
                  "','" & Val(lstRolls.ListItems(i).SubItems(4)) & "','" & Val(lstRolls.ListItems(i).SubItems(1)) & "','A','-','" & Val(lstRolls.ListItems(i).SubItems(5)) & "')"
   End If
   Next
    

Exit Sub
SETERR:
MsgBox ERR.Description
End Sub





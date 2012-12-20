VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "welchbutton.ocx"
Begin VB.Form frmPackingTransfer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing Transfer"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   11385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6824.804
   ScaleMode       =   0  'User
   ScaleWidth      =   30395.35
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   6795
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11986
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
      Begin VB.TextBox TXTMCCD 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   3375
      End
      Begin VB.ComboBox cmbToPackingType 
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
         ItemData        =   "frmPackingTransfer.frx":0000
         Left            =   480
         List            =   "frmPackingTransfer.frx":0002
         TabIndex        =   35
         Tag             =   "0"
         Text            =   "Select Type of Packing"
         Top             =   6240
         Width           =   2655
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
         ItemData        =   "frmPackingTransfer.frx":0004
         Left            =   2640
         List            =   "frmPackingTransfer.frx":0006
         TabIndex        =   1
         Tag             =   "0"
         Text            =   "Select Type of Packing"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtLTNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TXTGRAD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TXTITEM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1560
         Width           =   2775
      End
      Begin MSComctlLib.ListView lstBox 
         Height          =   3615
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6376
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
            Object.Width           =   354
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
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   12000
         TabIndex        =   20
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
         Image           =   "frmPackingTransfer.frx":0008
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   4320
         TabIndex        =   18
         Top             =   6000
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
         Image           =   "frmPackingTransfer.frx":03A2
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5400
         TabIndex        =   19
         Top             =   6000
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
         Image           =   "frmPackingTransfer.frx":112C
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6480
         TabIndex        =   21
         Top             =   6000
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
         Image           =   "frmPackingTransfer.frx":157E
         cBack           =   -2147483633
      End
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   7200
         TabIndex        =   3
         Top             =   720
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
      Begin MSMask.MaskEdBox dtTo 
         Height          =   330
         Left            =   9600
         TabIndex        =   5
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
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
      Begin WelchButton.lvButtons_H BTNSEARCH 
         Height          =   495
         Left            =   9960
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Search"
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
         Image           =   "frmPackingTransfer.frx":19D0
         cBack           =   -2147483633
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   825
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   3375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   11280
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "To Type of Packing"
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
         Left            =   720
         TabIndex        =   36
         Top             =   5955
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
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
         Left            =   8640
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
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
         Left            =   6000
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Station :"
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
         Left            =   5880
         TabIndex        =   34
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   5760
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label LBLDESC2 
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
         Left            =   7800
         TabIndex        =   33
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "From Type of Packing :"
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
         Left            =   360
         TabIndex        =   0
         Top             =   720
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   120
         Width           =   975
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
         TabIndex        =   30
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
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
         Left            =   3840
         TabIndex        =   8
         Tag             =   "S"
         Top             =   1320
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
         Left            =   9000
         TabIndex        =   14
         Tag             =   "S"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1335
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   600
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
         Left            =   8160
         TabIndex        =   12
         Tag             =   "S"
         Top             =   1320
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
         Left            =   5520
         TabIndex        =   10
         Tag             =   "S"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Machine"
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
         TabIndex        =   6
         Tag             =   "S"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
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
         Left            =   8400
         TabIndex        =   29
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Total Net Wt"
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
         Height          =   435
         Left            =   10200
         TabIndex        =   28
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Total Cops"
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
         Height          =   435
         Left            =   9240
         TabIndex        =   27
         Top             =   5880
         Width           =   780
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
         Left            =   9240
         TabIndex        =   26
         Top             =   6360
         Width           =   780
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
         Left            =   10080
         TabIndex        =   25
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label lblNTWT 
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
         Left            =   8400
         TabIndex        =   24
         Top             =   5880
         Width           =   855
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
         Left            =   8160
         TabIndex        =   23
         Top             =   6360
         Width           =   1005
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   945
         Left            =   8040
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   3255
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   825
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmPackingTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CHLNTYP As String
Dim LSPKGCOD As String
Dim DIVCODE As String
Dim TTLCOPS As Long
Dim TTLBOXES As Long
Dim TTLQTY As Double, TTLGRSWGT As Double, TTLTAREWGT As Double
Dim grad As String, SUBGRD As String, SGRD As String, SITEM As String, SMCCD As String
Dim VTCD As String
Dim M_DBCD As String

Private Sub btnSearch_Click()
  If Not IsDate(dtFrom) Or Not IsDate(dtTo) Then
     dtFrom.SetFocus
     Exit Sub
  End If
  
  If Trim(cmbPackingType.Text) = Empty Then
    cmbPackingType.SetFocus
    Exit Sub
  End If
  
  Call GenerateBoxList
End Sub

Private Sub cmbPackingType_Click()
If cmbPackingType <> Empty Then
    Call SetToPackingType
 End If
End Sub

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmdCancel_Click()
    Call ClsData(Me)
    lstBox.ListItems.Clear
    dtFrom = Format(Now, "dd/mm/yyyy")
    dtTo = Format(Now, "dd/mm/yyyy")
    If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 0
    If cmdExit.Enabled Then cmdExit.SetFocus
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST



Dim INDEX As Long
Dim FLAG As Boolean
Dim SQL As String
Dim TRF_DBCD As String

FLAG = False

For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then: FLAG = True: Exit For
Next

If FLAG = False Then Exit Sub

If RS.State = 1 Then RS.Close
RS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND FYCD='" & FYCD & "' AND NAME ='" & cmbToPackingType & "' ", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   TRF_DBCD = Trim(RS!CODE & "")
Else
   MsgBox "Unable To Transfer"
   cmbToPackingType.SetFocus
   Exit Sub
End If

CN.BeginTrans

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   
   SQL = "SELECT DBCD,CHLN FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
         "' AND PKG_STCOD='" & LSPKGCOD & "' AND DBCD='" & M_DBCD & "' AND VTYP='PPF' AND " & _
         "VBNO='" & lstBox.ListItems(INDEX).Text & "'"
   
   If RS.State = 1 Then RS.Close
   RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      RS!dbcd = TRF_DBCD
      RS.Update
      
      CN.Execute "UPDATE STORETRAN SET DBCD='" & TRF_DBCD & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
         "' AND DVCD='" & DIVCODE & "' AND DBCD='" & M_DBCD & "' AND VTYP='PPF' AND VBNO = '" & RS!chln & "' "
   
   End If
 End If
Next INDEX

CN.CommitTrans

MsgBox "Box Transfer Successfully."

lstBox.ListItems.Clear

Call CLEARDATA
Call cmdCancel_Click

Exit Sub
LAST:
MsgBox ERR.Description
Exit Sub
End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or LBLDIV = Empty Then
  Unload Me
End If

If LSPKGCOD = Empty Or LBLDESC2 = Empty Then
  Unload Me
End If

Me.BackColor = RGB(RED, GREEN, BLUE)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)

    Me.Left = 50: Me.KeyPreview = True

    
    '-------DIVISION MASTER
    M_DESC = Empty: Key = Empty: NEW_VISIBLE = False: DIVCODE = Empty
    If DIVCODE = Empty Then
      LBLDIV.Caption = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
      DIVCODE = Key
    End If
  
    '-------PACKING STATION MASTER
    M_DESC = Empty:  Key = Empty:  NEW_VISIBLE = False: LSPKGCOD = Empty
    LBLDESC2 = SearchList1("SELECT TOP 20 CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, "", "SELECT PACKING STATION FROM MASTER LIST")
    If Key = Empty Then Exit Sub
    LSPKGCOD = Key
    '---------------------------

    If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 2

    Call SetPackingType
    
    dtFrom = Format(Now, "DD/MM/YYYY")
    dtTo = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
    cmbPackingType.AddItem Trim(PKTYPRS!NAME)
    PKTYPRS.MoveNext
Loop
End Sub

Private Sub SetToPackingType()
cmbToPackingType.Clear
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "' AND NAME NOT LIKE '%" & cmbPackingType & "%' ", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
   cmbToPackingType.AddItem Trim(PKTYPRS!NAME)
   PKTYPRS.MoveNext
Loop
End Sub

Private Sub txtItem_GotFocus()
    txtitem.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTGRAD_GotFocus()
  TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
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
End Sub


Private Sub TXTMCCD_GotFocus()
  TXTMCCD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTMCCD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
   If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False:  M_DESC = Empty:   Key = Empty
        TXTMCCD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'", 0, TXTMCCD, "List of Machine Name")
   ElseIf KeyCode = vbKeyDelete Then
        TXTMCCD = Empty
   End If
Me.KeyPreview = True
End Sub

Private Sub TXTMCCD_LostFocus(): TXTMCCD.BackColor = vbWhite: End Sub

Private Sub TXTSUBGRD_GotFocus()
  TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub


Private Sub TXTLTNO_GotFocus()
  txtLTNO.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  
End Sub

Private Sub txtLTNO_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNO = Empty
ElseIf KeyCode = vbKeyF2 Then
   M_DESC = Empty:   NEW_VISIBLE = False
   txtLTNO = SearchList("SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'")
End If
   txtitem = FindItem
Me.KeyPreview = True
End Sub

Private Sub txtItem_LostFocus(): txtitem.BackColor = vbWhite: End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
Dim SQL As String
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
ElseIf KeyCode = vbKeyF2 Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT RDIFF,NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND GRAD='" & GetCode("GRDMST", TXTGRAD, "GRAD", "CODE") & "'"
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
End If
Me.KeyPreview = True
End Sub

Private Sub TXTSUBGRD_LostFocus(): TXTSUBGRD.BackColor = vbWhite: End Sub
Private Sub TXTLTNO_LostFocus(): txtLTNO.BackColor = vbWhite: End Sub

Private Sub CLEARDATA()
 txtitem = Empty: txtLTNO = Empty: TXTGRAD = Empty: TXTSUBGRD = Empty
 txtCTRN = Empty: txtCOPs = Empty: txtNTWT = Empty
End Sub
Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub lstBox_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim I As Integer, J As Integer, ctr As Integer
    If Item.Checked = True Then
        txtCTRN.Caption = Val(txtCTRN.Caption) + 1
        txtNTWT.Caption = nstr(Val(txtNTWT.Caption) + Val(Item.SubItems(2)), 10, 3)
        txtNTWT.Caption = Trim(txtNTWT.Caption)
        txtCOPs.Caption = Val(txtCOPs.Caption) + Val(Item.SubItems(1))
    Else
        txtCTRN.Caption = Val(txtCTRN.Caption) - 1
        txtNTWT.Caption = nstr(Val(txtNTWT.Caption) - Val(Item.SubItems(2)), 10, 3)
        txtNTWT.Caption = Trim(txtNTWT.Caption)
        txtCOPs.Caption = Val(txtCOPs.Caption) - Val(Item.SubItems(1))
    End If
    
    If Item.INDEX < lstBox.ListItems.COUNT Then lstBox.ListItems.Item(Item.INDEX + 1).Selected = True: lstBox.ListItems(Item.INDEX + 1).EnsureVisible
End Sub

Private Function FindItem() As String
Dim FICD As String
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset


If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNO & "'", CN, adOpenDynamic, adLockOptimistic
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

'If TXTITEM <> Empty Or TXTGRAD = Empty Or txtLTNO = Empty Then Exit Sub

Dim SQL As String
Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset

SQL = "SELECT * FROM BOXREGISTER WHERE BOXREGISTER.COMP = '" & compPth & "' AND BOXREGISTER.UNIT = '" & UNCD & _
      "' AND BOXREGISTER.DVCD = '" & DIVCODE & "' AND VTYP='PPF' AND BOXREGISTER.RECSTAT<>'D'  AND BOXREGISTER.VBDT> = '" & Format(dtFrom, "MM/DD/YYYY") & "' AND BOXREGISTER.VBDT <= '" & Format(dtTo, "MM/DD/YYYY") & "' AND RVBNO IS NULL "
      
'FOR PACKING STATION
If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "' AND NAME = '" & cmbPackingType.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   M_DBCD = Trim(RSDATA!CODE & "")
End If

SQL = SQL & " AND BOXREGISTER.DBCD = '" & M_DBCD & "' "

'FOR ITEM
If txtitem <> Empty Then
If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND NAME = '" & txtitem & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SITEM = Trim(RSDATA!CODE & "")
   SQL = SQL & " AND BOXREGISTER.ICOD = '" & SITEM & "' "
End If
RSDATA.Close
End If

'FOR GRADE
If TXTGRAD <> Empty Then
   SGRD = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")
   SQL = SQL & " AND BOXREGISTER.GRAD = '" & SGRD & "' "
End If

'FOR SUBGRADE
If TXTSUBGRD <> Empty Then
If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND GRAD = '" & SGRD & "' AND NAME = '" & TXTSUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SUBGRD = Trim(RSDATA!SUBGRD & "")
   SQL = SQL & " AND BOXREGISTER.SUBGRD = '" & SUBGRD & "' "
End If
RSDATA.Close
End If

'FOR MACHINE
If TXTMCCD <> Empty Then
If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND NAME = '" & TXTMCCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SMCCD = Trim(RSDATA!CODE & "")
   SQL = SQL & " AND BOXREGISTER.MCCD = '" & SMCCD & "' "
End If
RSDATA.Close
End If

'FOR LOT
If txtLTNO <> Empty Then
   SQL = SQL & " AND BOXREGISTER.LOTNO = '" & txtLTNO & "' "
End If

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   MsgBox "Boxes are not available for this criteria."
   TXTGRAD.Enabled = True: TXTGRAD.SetFocus
   Exit Sub
End If
  lstBox.ListItems.Clear
  Do While Not RSDATA.EOF
   Set Item = lstBox.ListItems.ADD
   Item.Text = RSDATA!VBNO
   Item.SubItems(1) = RSDATA!COPS
   Item.SubItems(2) = nstr(RSDATA!NTWGT, 9, 3)
   Item.SubItems(2) = Trim(Item.SubItems(2))
   If SUBGRD = "S" Or SUBGRD = "Z" Or SUBGRD = "0" Then
     Item.SubItems(3) = SUBGRD
     If lstBox.SelectedItem.ListSubItems.COUNT = 2 Then lstBox.ColumnHeaders(4).Text = "Twist"
   Else
     Item.SubItems(3) = SUBGRD
     If lstBox.ListItems.COUNT = 1 Then lstBox.ColumnHeaders(4).Text = "SG"
   End If
   
   Item.SubItems(4) = nstr(RSDATA!GRSWGT, 9, 3)
   Item.SubItems(4) = Trim(Item.SubItems(4) & "")
   Item.SubItems(5) = nstr(RSDATA!TRWGT, 9, 3)
   Item.SubItems(5) = Trim(Item.SubItems(5) & "")
   Item.SubItems(6) = Format(RSDATA!VBDT, "DD/MM/YYYY")
   Item.SubItems(7) = Trim(RSDATA!RMRK & "")
   Item.SubItems(8) = Trim(RSDATA!PKG_STCOD & "")
   Item.SubItems(9) = Trim(RSDATA!ISRETURNABLE & "")
      
   RSDATA.MoveNext
  Loop
  RSDATA.Close
  
End Sub


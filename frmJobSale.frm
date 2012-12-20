VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmJobSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Module"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleMode       =   0  'User
   ScaleWidth      =   11385.16
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   6840
   End
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12938
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
      Begin VB.ComboBox TXTRMRK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   1  'Simple Combo
         TabIndex        =   29
         Top             =   5760
         Width           =   5775
      End
      Begin VB.TextBox TXTCHLN 
         Height          =   285
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtDEST 
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   27
         Top             =   5400
         Width           =   3255
      End
      Begin VB.TextBox txtLRNO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   21
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtFREIGHT 
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   19
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox TXTCONSIGNEE 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox M_RTTX 
         Height          =   315
         ItemData        =   "frmJobSale.frx":0000
         Left            =   7200
         List            =   "frmJobSale.frx":000A
         TabIndex        =   12
         Text            =   "M_RTTX"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox M_BRNM 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox M_TXNM 
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
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
         TabIndex        =   49
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox TXTPARTY 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox TXTCRDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   25
         Top             =   5400
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
         Left            =   5520
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   4320
         Width           =   1695
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
         Left            =   3120
         TabIndex        =   44
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Frame FRMBTRM 
         Height          =   2295
         Left            =   7320
         TabIndex        =   39
         Top             =   4320
         Width           =   3975
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
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtBEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1440
            TabIndex        =   41
            Top             =   1320
            Visible         =   0   'False
            Width           =   345
         End
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   1800
            Width           =   1905
         End
         Begin MSFlexGridLib.MSFlexGrid flexBTRM 
            Height          =   1635
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   3735
            _ExtentX        =   6588
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
            TabIndex        =   43
            Top             =   1920
            Width           =   1305
         End
      End
      Begin VB.ComboBox CMBSELECTION 
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
         ItemData        =   "frmJobSale.frx":002B
         Left            =   1680
         List            =   "frmJobSale.frx":002D
         TabIndex        =   1
         Tag             =   "0"
         Text            =   "CMBSELECTION"
         Top             =   120
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   330
         Left            =   6480
         TabIndex        =   2
         Top             =   120
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
         Format          =   24182785
         CurrentDate     =   38429
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   720
         TabIndex        =   32
         Top             =   6720
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmJobSale.frx":002F
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2160
         TabIndex        =   33
         Top             =   6720
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmJobSale.frx":0DB9
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   3600
         TabIndex        =   34
         Top             =   6720
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmJobSale.frx":120B
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdGo 
         Height          =   375
         Left            =   9960
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Filter"
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
         Image           =   "frmJobSale.frx":165D
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView lstChallan 
         Height          =   2175
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Chln.No"
            Object.Width           =   2239
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Chln Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agent Name"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Consinee Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Pcs"
            Object.Width           =   795
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Quantity"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Item Desc."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Consinee Address"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "PCOD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "DCOD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "BRCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "ADD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "CHLNDBCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "GRNNO"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.DTPicker LRDT 
         Height          =   330
         Left            =   3840
         TabIndex        =   23
         Top             =   5040
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
         Format          =   24182785
         CurrentDate     =   38429
      End
      Begin MSMask.MaskEdBox txtPR 
         Height          =   330
         Left            =   1800
         TabIndex        =   31
         Top             =   6120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "&Preparation Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Challan &No."
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
         Left            =   9720
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Destina&tion"
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
         Left            =   2760
         TabIndex        =   26
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label11 
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
         Left            =   240
         TabIndex        =   20
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "L.R &Dt."
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
         Left            =   2760
         TabIndex        =   22
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight / KG"
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
         Left            =   240
         TabIndex        =   18
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label18 
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
         Left            =   5520
         TabIndex        =   11
         Tag             =   "S"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Category"
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
         Left            =   5880
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000080&
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1335
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   11070
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pcs"
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
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label LBLNAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label LBLPARTY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Days"
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
         Left            =   240
         TabIndex        =   24
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   240
         TabIndex        =   28
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label lblBill 
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
         Left            =   9480
         TabIndex        =   48
         Top             =   120
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   8040
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         Left            =   2280
         TabIndex        =   47
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Amt "
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
         Left            =   4560
         TabIndex        =   46
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No. :"
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
         Left            =   8160
         TabIndex        =   38
         Top             =   120
         Width           =   1335
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
         TabIndex        =   37
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of  Sale :"
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
         TabIndex        =   36
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date :"
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
         Left            =   5040
         TabIndex        =   35
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmJobSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DIVCODE As String
Dim DIVNAME As String
Dim M_DBCD As String
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
Dim itemcode As String
Dim M_BRCD As String
Dim M_TXCD As String
Dim SAVEFLAG As Boolean
'JOB WORK VARIABLE
Dim GRNNO As String
Dim GRNITMRATE As Double
Dim LESSOTHER As Double
Dim RMCOST As Double
Dim ISJOBBILL As Boolean
Dim ITMRO As String
'-----------------------

Private Sub CMBSELECTION_Click()
 SendKeys "{HOME}"
 Call FindSerial
 Call FillList
End Sub

Private Sub CMBSELECTION_GotFocus()
  SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbSelection_KeyDown(KeyCode As Integer, Shift As Integer)
  'KeyCode = 0
End Sub

Private Sub cmbSelection_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
  TXTVBDT.Enabled = True: TXTVBDT.SetFocus
End Sub

Private Sub cmdCancel_Click()

    If TXTCHLN <> Empty Then
       lstChallan.ListItems.Clear
       txtParty = Empty
       txtConsignee = Empty
       M_BRNM = Empty
       TXTCHLN = Empty
       TXTCHLN.SetFocus
       TXTFREIGHT = 0
    End If
    
    TXTCRDS = Empty
    TXTRMRK = Empty
    TXTLRNO = Empty
    TXTTPCS = "0"
    TXTTQTY = "0.000"
    TXTITOT = "0.00"
    TXTADLS = "0.00"
    TXTBNET = "0.00"
    Call FindSerial
    Call FillList
    Call FillCombo("Select Distinct BRMK from BILLMAIN where BRMK is not null or BRMK<>''", TXTRMRK)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGo_Click()
If TXTCHLN = Empty Then
    If txtParty = Empty Then txtParty.Enabled = True: txtParty.SetFocus: Exit Sub
    If M_BRNM = Empty Then M_BRNM.Enabled = True: M_BRNM.SetFocus: Exit Sub
    If M_TXNM = Empty Then M_TXNM.Enabled = True: M_TXNM.SetFocus: Exit Sub
End If
    
    Call FillList
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

'CHECKING OF DATA
 If Not IsValidSelection Then
     MsgBox "Don't Mix Job Challan With Other Challan.", vbCritical
     Exit Sub
 End If
'====================

Dim i As Long, J As Long, COUNTER As Long
Dim SQL As String, RTCD As String, DPFCODE As String

COUNTER = 0
For i = 1 To lstChallan.ListItems.COUNT
 If lstChallan.ListItems(i).Checked = True Then
  COUNTER = COUNTER + 1
  If COUNTER = 1 Then
     DPFCODE = Trim(lstChallan.ListItems(i).ListSubItems(15))
     DRAC = Trim(lstChallan.ListItems(i).ListSubItems(11))
     PCOD = Trim(lstChallan.ListItems(i).ListSubItems(11))
     DCOD = Trim(lstChallan.ListItems(i).ListSubItems(13))
     ADDRESS = Trim(lstChallan.ListItems(i).ListSubItems(14))
     GRNNO = Trim(lstChallan.ListItems(i).ListSubItems(16))
     CPCD = GetCode("ACCMST", DRAC, "CODE", "CPCD")
     ARCD = GetCode("ACCMST", DRAC, "CODE", "ARCD")
  End If
 End If
Next

BRCD = GetCode("REFMST", M_BRNM, "NAME", "CODE")
TTYP = Trim(M_RTTX)
TXCD = GetCode("TAXMST", M_TXNM, "NAME", "CODE")
RTCD = GetCode("TAXMST", M_TXNM, "NAME", "RATE_CODE")

If COUNTER = 0 Then MsgBox "First Select Then Save Data": lstChallan.SetFocus: Exit Sub
'If Val(TXTCRDS) = 0 Then MsgBox "Credit Days Empty": TXTCRDS.SetFocus: Exit Sub
   
Call CalculateQtyAmt

COUNTER = 0
Call FindSerial

Dim CHKDAT As ADODB.Recordset
Set CHKDAT = New ADODB.Recordset
If CHKDAT.State = 1 Then CHKDAT.Close
CHKDAT.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & Trim(lblBill.Caption) & "'", CN, adOpenDynamic, adLockOptimistic
If Not CHKDAT.EOF Then
   MsgBox "Bill No. Already Exist.Change Bill From Unit Configuration."
   cmdSave.SetFocus
   Exit Sub
End If
CHKDAT.Close

CN.BeginTrans

Dim SAVDAT As ADODB.Recordset
Set SAVDAT = New ADODB.Recordset

For i = 1 To lstChallan.ListItems.COUNT
 If lstChallan.ListItems(i).Checked = True Then
 COUNTER = COUNTER + 1
 Dim RECSET As ADODB.Recordset
 Set RECSET = New ADODB.Recordset
 Dim EXCO As String
 Dim CHAP As String

SQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND VTYP='DPF' AND RECSTAT='A' AND VBNO = '" & Trim(lstChallan.ListItems(i)) & "' AND "
SQL = SQL & "DBCD = '" & Trim(lstChallan.ListItems(i).ListSubItems(15)) & "'"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
Else
   
   If Not RECSET.EOF Then
      '---------------------------------------------------------------------------------------
      '------------- EXCISE INFORMATION -------------------
      'Exciseable Commodity and chapter No.
      
      If SAVDAT.State = 1 Then SAVDAT.Close
      If Trim(RECSET!ltno & "") = "" Or Trim(RECSET!ltno & "") = "WASTE" Then
         SAVDAT.Open "SELECT WEXCO AS EXCO,WCHAP AS CHAP FROM UNTCFG WHERE COMP='" & compPth & _
         "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
      Else
         SAVDAT.Open "SELECT EXCOMMODITY AS EXCO,CHAPTERNO AS CHAP FROM DIVMST WHERE COMP='" & compPth & _
         "' AND UNIT='" & UNCD & "' AND CODE='" & DIVCODE & "'", CN, adOpenDynamic, adLockOptimistic
      End If
      If Not SAVDAT.EOF Then
        EXCO = Trim(SAVDAT!EXCO & "")
        CHAP = Trim(SAVDAT!CHAP & "")
      End If
      SAVDAT.Close
      '---------------------------------------------------------------------------------------
    End If
    
End If

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "BRCD,DCOD,ADDRESS,LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,ARAT,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,VEHICALNO,TRCD,TXCD,RTCD,GRNNO)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','SAL','" & M_DBCD & "','" & Trim(lblBill.Caption) & "','" & Trim(lstChallan.ListItems(i)) & _
"','" & Format(Trim(lstChallan.ListItems(i).ListSubItems(1)), "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & RECSET!DRAC & "','" & RECSET!PCOD & _
"','" & BRCD & "','" & RECSET!DCOD & "','" & RECSET!ADDRESS & "','" & RECSET!ltno & "','" & RECSET!ICOD & "','" & RECSET!grad & _
"','" & RECSET!SUBGRD & "','" & RECSET!PCES & "','" & RECSET!QNTY & "'," & GetReverseRate(Trim(M_BRNM), Trim(M_TXNM), Val(RECSET!RATE), 0) & _
"," & RECSET!RATE & "," & RECSET!QNTY * GetReverseRate(Trim(M_BRNM), Trim(M_TXNM), Val(RECSET!RATE), 0) & _
",'Q','N','" & cUName & "','-','A','" & RECSET!COPS & "','" & Trim(RECSET!VEHICALNO) & _
"','" & RECSET!TRCD & "','" & TXCD & "','" & RTCD & "','" & GRNNO & "')"

CN.Execute SQL
  
SQL = "UPDATE SPTRAN SET RTYP='SAL',SDBC='" & M_DBCD & "',SVBN='" & Trim(lblBill.Caption) & _
"',BRCD='" & BRCD & "',TXCD='" & TXCD & "',RTCD='" & RTCD & "' WHERE COMP='" & compPth & _
"' AND UNIT = '" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND VTYP='DPF' AND RECSTAT='A' AND VBNO = '" & Trim(lstChallan.ListItems(i)) & _
"' AND DBCD = '" & Trim(lstChallan.ListItems(i).ListSubItems(15)) & "'"
   
CN.Execute SQL

SQL = "UPDATE PKGSTK SET BRCD='" & BRCD & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
"' AND VTYP='DPF' AND RECSTAT='A' AND CHLN = '" & Trim(lstChallan.ListItems(i)) & _
"' AND DBCD = '" & Trim(lstChallan.ListItems(i).ListSubItems(15)) & "'"
 
End If
Next
RECSET.Close

  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO = '" & lblBill.Caption & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "SAL"
  SAVDAT!SRNO = 0
  SAVDAT!SRCH = 1
  SAVDAT!dbcd = M_DBCD
  SAVDAT!Date = Format(TXTVBDT.Value, "YYYY/MM/DD")
  SAVDAT!VBNO = Trim(lblBill)
  SAVDAT!CVBN = Trim(lblBill)
  SAVDAT!CRAC = "XXXXXX"
  SAVDAT!DRAC = DRAC
  SAVDAT!PCOD = PCOD
  SAVDAT!DCOD = DCOD
  SAVDAT!ADDRESS = ADDRESS
  SAVDAT!BRCD = BRCD
  SAVDAT!CPCD = CPCD
  SAVDAT!ARCD = ARCD
  SAVDAT!TXCD = TXCD
  SAVDAT!TAXGRP = GetCode("TAXMST", TXCD & "", "CODE", "GRPCOD")
  SAVDAT!TPCS = Val(TXTTPCS.Text)
  SAVDAT!TQTY = Val(TXTTQTY.Text)
  SAVDAT!ITOT = Val(TXTITOT.Text)
  SAVDAT!BADJ = Val(TXTBNET.Text) - Val(TXTITOT.Text)
  SAVDAT!BNET = Val(TXTBNET.Text)
  SAVDAT!TTYP = TTYP
  SAVDAT!CDAY = Val(TXTCRDS)
  SAVDAT!PRTM = Trim(txtPR)
  
  If SAVEFLAG = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  
  SAVDAT![User] = cUName & ""
  SAVDAT!DVCD = DIVCODE
  SAVDAT!unit = UNCD
  SAVDAT!TRCD = ""
  SAVDAT!RECSTAT = "A"
  SAVDAT!BRMK = Trim(TXTRMRK.Text)
  
  SAVDAT!LRNO = TXTLRNO
  SAVDAT!LRDT = Format(LRDT.Value, "YYYY/MM/DD")
  SAVDAT!DSTN = GetCode("CITYMASTER", txtDEST, "NAME", "CODE")
  SAVDAT!EXCO = EXCO
  SAVDAT!CHAP = CHAP
  SAVDAT!EXCSRNO = GetExciseSerial(DIVCODE)
  
  If ISJOBBILL Then
     Call FindGrnDetail
     'FOR JOBWORK
      SAVDAT!ITOT = Round(SAVDAT!ITOT, 0) - Round(TXTTQTY * GRNITMRATE, 0)
      SAVDAT!GRNNO = GRNNO
      SAVDAT!GRNRATE = GRNITMRATE
      SAVDAT!LESSOTHER = Round(LESSOTHER, 0)
      SAVDAT!RMCOST = RMCOST
      SAVDAT!BNET = SAVDAT!BNET - SAVDAT!LESSOTHER - Round(TXTTQTY * GRNITMRATE, 0)
      TXTBNET.Text = Val(SAVDAT!BNET)
      
     '----------------------------
  End If
      
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
    
  Dim EXCISE As ADODB.Recordset
  Set EXCISE = New ADODB.Recordset
  If EXCISE.State = 1 Then EXCISE.Close
  EXCISE.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  EXCISE.AddNew
  EXCISE!COMP = compPth
  EXCISE!VTYP = "SAL"
  EXCISE!SRNO = 0
  EXCISE!SRCH = 1
  EXCISE!Date = Format(TXTVBDT, "YYYY/MM/DD")
  EXCISE!dbcd = M_DBCD
  EXCISE!CRAC = "XXXXXX"
  EXCISE!DRAC = DRAC & ""
  EXCISE!DCOD = DCOD & ""
  EXCISE!ADDRESS = ADDRESS & ""
  EXCISE!BRCD = BRCD
  EXCISE!CPCD = CPCD
  EXCISE!ARCD = ARCD
  EXCISE!TXCD = TXCD
  EXCISE!TAXGRP = GetCode("TAXMST", TXCD, "CODE", "GRPCOD")
  EXCISE!VBNO = lblBill
  EXCISE!chln = lblBill
  EXCISE!CHDT = Format(TXTVBDT, "YYYY/MM/DD")
  EXCISE!TRCD = "" 'TRCD
  EXCISE!LRNO = "" 'Trim(TXTLRNO.Text)
 'EXCISE!LRDT = "" 'Format(TXTLRDT.Value, "YYYY/MM/DD")
  EXCISE!ICOD = itemcode
  EXCISE!PCES = Val(TXTTPCS)
  EXCISE!QNTY = Val(TXTTQTY)
  EXCISE!AMNT = Val(TXTITOT)
  EXCISE!ITOT = Val(TXTITOT)
  EXCISE!BADJ = Val(TXTBNET.Text) - Val(TXTITOT.Text)
  EXCISE!BNET = Val(TXTBNET)
  EXCISE!RORT = TTYP
  EXCISE!TTYP = TXCD
  EXCISE!RECSTAT = "A"
  EXCISE!unit = UNCD
  EXCISE!EXCO = EXCO
  EXCISE!CHAP = CHAP
  
  i = 0
  For i = 0 To flexBTRM.Rows - 1
    J = 0
    For J = 0 To EXCISE.Fields.COUNT - 1
      If Trim(EXCISE.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
        EXCISE.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 2))
      End If
      If Trim(EXCISE.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
        EXCISE.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 1))
      End If
    Next
  Next
  EXCISE.Update
  
  
SQL = "UPDATE SERIALMASTER SET [SRNO]='" & lblBill & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
"' AND VTYP='SAL' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "'"
   
CN.Execute SQL

SQL = "UPDATE DIVMST SET [SRNO]='" & GetExciseSerial(DIVCODE) & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
      "' AND CODE ='" & DIVCODE & "' "
   
CN.Execute SQL

'---------------------------------
'DAILYSTATUS ENTRY
  Call DAILYSTATUS("SAL", DRAC, M_DBCD, Val(TXTTQTY), lblBill, Val(TXTITOT), cUName, "N", Now, TXTVBDT)
  
'---------------------------------

CN.CommitTrans

MsgBox "Your Invoice No. is : " & lblBill.Caption

 If IsOnlineBillPrintReq Then 'IS PRINTING ONLINE CHALLAN REQUIRED ???
     
    OnlineBillNum = lblBill
    
    LOAD frmRPT_InvPrinting
    frmRPT_InvPrinting.Hide
        
    frmRPT_InvPrinting.cboStatus.ListIndex = 0
    frmRPT_InvPrinting.txtUNIT = UntNm
    frmRPT_InvPrinting.txtUNIT.Tag = UNCD
    frmRPT_InvPrinting.txtDVCD = DIVNAME
    frmRPT_InvPrinting.txtDVCD.Tag = DIVCODE
        
    frmRPT_InvPrinting.cmbSaleType.AddItem cmbSelection.Text
    frmRPT_InvPrinting.cmbSaleType.Text = cmbSelection
        
    frmRPT_InvPrinting.lstInvoice_GotFocus
    frmRPT_InvPrinting.opPlain.Value = True
    frmRPT_InvPrinting.cmdpreview_Click
    
 End If

Call cmdCancel_Click
Call FillList
Call FillCombo("Select Distinct BRMK from BILLMAIN where BRMK is not null or BRMK<>''", TXTRMRK)

Exit Sub

LAST:
MsgBox ERR.Description
Exit Sub
End Sub

Private Sub flexBTRM_KeyDown(KeyCode As Integer, Shift As Integer)
With flexBTRM
If KeyCode = vbKeyReturn And .Rows > 0 And (.COL = 1 Or .COL = 2) Then
 If .COL = 1 Then
    .COL = .COL + 1
    Exit Sub
 End If
 
 If .COL = 2 And .ROW <> .Rows - 1 Then
    .ROW = .ROW + 1
    .COL = 1
 Else
    TXTFREIGHT.Enabled = True
   TXTFREIGHT.SetFocus
 End If
End If
End With
End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or DIVNAME = Empty Then
  Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)

  flexBTRM.ColWidth(0) = 1500
  flexBTRM.ColWidth(1) = 800
  flexBTRM.ColWidth(2) = 1200
  
  M_DESC = Empty
  Key = Empty
  NEW_VISIBLE = False
  DIVCODE = Empty
  DIVNAME = Empty
  
  If DIVCODE = Empty Then
    DIVNAME = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
  Me.Caption = "SALE MODULE ( DIVISION : " + DIVNAME + " )"
      
  Call SetSaleType
  
  TXTVBDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  LRDT = Now
  If M_RTTX.ListCount > 0 Then M_RTTX.ListIndex = 0
  
  ITMRO = "N"
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT ITEMRO FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    ITMRO = RS!ITEMRO & ""
  End If
  RS.Close
  
  Call FillCombo("Select Distinct BRMK from BILLMAIN where BRMK is not null or BRMK<>''", TXTRMRK)
  txtPR = Format(CStr(Now), "HH:MM")
End Sub

Private Sub SetSaleType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND FYCD='" & FYCD & "'  AND ACTIVE='Y'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbSelection.AddItem Trim(PKTYPRS!NAME)
PKTYPRS.MoveNext
Loop

If cmbSelection.ListCount >= 1 Then cmbSelection.ListIndex = 0

End Sub

Private Sub lstChallan_Click()
  
 Dim i As Long, COUNTER As Long, POS As Long
 Dim SQL As String
 
 If Not IsValidSelection Then
     MsgBox "Don't Mix Job Challan With Other Challan.", vbCritical
     Exit Sub
 Else
     COUNTER = 0
     For i = 1 To lstChallan.ListItems.COUNT
         If lstChallan.ListItems(i).Checked = True Then
            COUNTER = COUNTER + 1
            If COUNTER = 1 Then POS = i
         End If
     Next
     
     If Not ISJOBBILL Then
        Call CalculateQtyAmt
        If COUNTER = 1 Then
            SQL = SQL & " AND SPTRAN.DBCD ='" & Trim(lstChallan.ListItems(POS).ListSubItems(15)) & "' "
            Call FillList(SQL, Trim(lstChallan.ListItems(POS)))
        ElseIf COUNTER = 0 Then
            Call FillList
        End If
        Exit Sub
     End If
 End If
 
 If COUNTER = 0 Then
     TXCD = Empty
     Call FIL_Billingterm
     Call FillList
     Exit Sub
 End If

 If COUNTER = 1 Then

     If Trim(lstChallan.ListItems(POS).ListSubItems(15)) = "000003" Then
        SQL = SQL & " AND SPTRAN.GRNNO ='" & Trim(lstChallan.ListItems(POS).ListSubItems(16)) & "' "
     End If
    
     Call FillList(SQL, Trim(lstChallan.ListItems(POS)))
     Call FIL_Billingterm
     calBTRM 0
     Call calADLS
     
 End If
 
Call CalculateQtyAmt
    
End Sub

Private Sub lstChallan_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 Call lstChallan_Click
End Sub

Private Sub lstChallan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  TXTFREIGHT.Enabled = True
  TXTFREIGHT.SetFocus
End If
End Sub

Private Sub lstChallan_LostFocus()
Call CalculateQtyAmt
End Sub

Private Sub lstChallan_Validate(Cancel As Boolean)
Call CalculateQtyAmt
End Sub

Private Sub M_BRNM_Change()
If M_BRNM <> Empty Then
   Call cmdGo_Click
End If
End Sub

Private Sub M_BRNM_GotFocus()
M_BRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_BRNM_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
 If KeyCode = vbKeyF2 Or (Trim(M_BRNM) = Empty And KeyCode = vbKeyReturn) Then
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    M_BRNM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM REFMST WHERE CATA='B'", 0, M_BRNM.Text, "SELECT AGENT FROM LIST")
    If key_PressNew = True Then
       M_DESC = ""
       Key = ""
       Ref_Cat = "B"
       M_BRNM.Text = ""
       Frm_Ref_FAS.Show
    Else
       M_BRCD = Key
       End If
    End If
Me.KeyPreview = True
End Sub

Private Sub M_BRNM_LostFocus()
M_BRNM.BackColor = vbWhite
End Sub

Private Sub M_RTTX_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  If flexBTRM.Rows > 0 Then
     flexBTRM.ROW = 0
     flexBTRM.COL = 1
     flexBTRM.SetFocus
     Exit Sub
  End If
End If

TXTCRDS.Enabled = True
TXTCRDS.SetFocus
End Sub

Private Sub M_TXNM_Change()
  If M_TXNM <> Empty Then
     TXCD = GetCode("TAXMST", M_TXNM, "NAME", "CODE")
     Call cmdGo_Click
     Call FIL_Billingterm
  Else
     TXCD = Empty
     Call FIL_Billingterm
  End If
 calBTRM 0
 Call calADLS
End Sub

Private Sub M_TXNM_GotFocus()
M_TXNM.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TXNM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or (Trim(M_TXNM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_TXNM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM TAXMST WHERE RECSTAT='A'", 0, M_TXNM.Text, "SELECT TAX FROM LIST")
        If key_PressNew = True Then
            M_DESC = "": Key = "":  M_TXNM.Text = ""
            FrmSaleTaxMaster.Show
        Else
            M_TXCD = Key
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub M_TXNM_LostFocus()
M_TXNM.BackColor = vbWhite
End Sub

Private Sub txtFREIGHT_GotFocus()
  TXTFREIGHT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTFREIGHT_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTFREIGHT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtFREIGHT_LostFocus()
  TXTFREIGHT.BackColor = vbWhite
  Dim i As Double
  i = 0
  For i = 0 To flexBTRM.Rows - 1
   If UCase(flexBTRM.TextMatrix(i, 0)) = "FREIGHT" Then
     flexBTRM.TextMatrix(i, 1) = Val(TXTFREIGHT)
     Call calBTRM(0)
   End If
  Next
End Sub

Private Sub txtCHLN_Change()
If Len(TXTCHLN) = 10 Then
   Call cmdGo_Click
   lstChallan.SetFocus
End If
End Sub

Private Sub TXTCHLN_GotFocus()
  TXTCHLN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCHLN_LostFocus()
  TXTCHLN.BackColor = vbWhite
End Sub

Private Sub txtConsignee_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtConsignee = Empty
  ElseIf KeyCode = vbKeyF2 Or txtConsignee = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtConsignee = SearchList1("Select DISTINCT CODE,NAME From PADDMST WHERE RECSTAT='A'", 0, Empty, "Select Consinee Name ")
     txtConsignee.Tag = Key
  End If
 Me.KeyPreview = True
End Sub

Private Sub TXTCRDS_GotFocus()
  TXTCRDS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCRDS_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TXTCRDS_LostFocus()
  TXTCRDS.BackColor = vbWhite
End Sub

Private Sub txtParty_GotFocus()
txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
 
 If KeyCode = vbKeyF2 Or (Trim(txtParty) = Empty And KeyCode = vbKeyReturn) Then
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    txtParty.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST WHERE DRCR='D'", 0, txtParty.Text, "SELECT PARTY FROM LIST")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            txtParty.Text = ""
            frm_Acc.Show
        Else
            txtParty.Tag = Key
        End If
 End If
 
Me.KeyPreview = True
End Sub

Private Sub txtParty_LostFocus()
txtParty.BackColor = vbWhite

    Dim GETRS As ADODB.Recordset
     Set GETRS = New ADODB.Recordset
  
     If GETRS.State = 1 Then GETRS.Close
     GETRS.Open "SELECT BRCD,RCOD,TXCD,TTYP,CDAY FROM ACCMST WHERE NAME='" & txtParty & "' ", CN, adOpenDynamic, adLockOptimistic
     If Not GETRS.EOF Then
        M_BRNM = GetCode("REFMST", GETRS!BRCD & "", "CODE", "NAME")
        M_BRCD = Trim(GETRS!BRCD & "")
        M_TXNM = GetCode("TAXMST", GETRS!TXCD & "", "CODE", "NAME")
        M_TXCD = Trim(GETRS!TXCD & "")
        txtConsignee = GetCode("PADDMST", GETRS!RCOD & "", "CODE", "NAME")
        txtConsignee.Tag = Trim(GETRS!RCOD & "")
        TXTCRDS = Val(GETRS!CDAY)
        M_RTTX = Trim(GETRS!TTYP & "")
     End If
     
End Sub

Private Sub txtConsignee_GotFocus()
  txtConsignee.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtConsignee_LostFocus()
  txtConsignee.BackColor = vbWhite
End Sub

Private Sub TXTRMRK_GotFocus()
    TXTRMRK.Height = 1155
    TXTRMRK.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRMRK_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case 39
            ' This is a (') Symbol
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRMRK_LostFocus()
  TXTRMRK.BackColor = vbWhite
  TXTRMRK.Height = 325
End Sub

Public Sub FillList(Optional FILTER As String, Optional chln As String)

If txtParty = Empty And txtConsignee = Empty And TXTCHLN = Empty Then Exit Sub

Dim SQL As String
Dim M_ROW As Integer

lstChallan.ListItems.Clear

Screen.MousePointer = vbHourglass
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT DISTINCT SPTRAN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITEM,PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS CADDRESS,"
SQL = SQL & "PADDMST.ADDR AS CADDRESS FROM SPTRAN INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.DRAC "
SQL = SQL & "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS "
SQL = SQL & "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
"' AND SPTRAN.DVCD='" & DIVCODE & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.RECSTAT='A' AND SVBN IS NULL "

If txtParty <> Empty Then
   SQL = SQL & "AND SPTRAN.DRAC ='" & Trim(txtParty.Tag) & "' "
End If

If TXTCHLN <> Empty Then
   SQL = SQL & "AND SPTRAN.VBNO = '" & Trim(TXTCHLN) & "' "
End If

If txtConsignee <> Empty Then
   SQL = SQL & "AND SPTRAN.DCOD ='" & Trim(txtConsignee.Tag) & "' "
End If

If InStr(1, UCase(cmbSelection.Text), "EXPORT") <> 0 Then
   SQL = SQL & " AND SPTRAN.DBCD ='000002' "
End If

If FILTER <> Empty Then SQL = SQL & FILTER

SQL = SQL & " ORDER BY SPTRAN.VBNO"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
Else
   If chln <> Empty Then
      TXTLRNO = Trim(RECSET!LRNO & "")
      If Not IsNull(RECSET!LRDT) Then
         LRDT = Format(RECSET!LRDT & "", "DD/MM/YYYY")
      End If
   Else
      TXTLRNO = Empty
      LRDT = Format(Now, "DD/MM/YYYY")
   End If
End If
   
    Do While RECSET.EOF = False
        Set Item = lstChallan.ListItems.ADD
        Item.Text = RECSET!VBNO
        Item.SubItems(1) = RECSET!Date
        Item.SubItems(2) = RECSET!ACNM
                
        Item.SubItems(4) = RECSET!CONSINEE
                        
        Item.SubItems(5) = RECSET!PCES
        Item.SubItems(6) = nstr(RECSET!QNTY, 12, 3)
        Item.SubItems(7) = GetReverseRate(Trim(M_BRNM), Trim(M_TXNM), Val(RECSET!RATE), Val(TXTFREIGHT))
        'Item.SubItems(8) = RECSET!AMNT
        
        Dim NEWAMT As Double
        NEWAMT = Val(Item.SubItems(6)) * Val(Item.SubItems(7))
        
        If ITMRO = "Y" Then
          Item.SubItems(8) = Round(NEWAMT, 0)
         Else
          Item.SubItems(8) = NEWAMT
        End If
        
        Item.SubItems(9) = RECSET!Item & ""
        Item.SubItems(10) = RECSET!CADDRESS
        Item.SubItems(11) = RECSET!DRAC
        Item.SubItems(12) = RECSET!PCOD
        Item.SubItems(13) = RECSET!DCOD
        Item.SubItems(14) = RECSET!ADDRESS
        Item.SubItems(15) = Trim(RECSET!dbcd & "")
        Item.SubItems(16) = Trim(RECSET!GRNNO & "")
        
        If TXTCHLN <> Empty Then
            txtParty = RECSET!ACNM
            txtParty.Tag = RECSET!DRAC
            txtConsignee = RECSET!CONSINEE
            txtConsignee.Tag = RECSET!DCOD
            M_BRNM = GetCode("REFMST", GetCode("ACCMST", RECSET!PCOD, "CODE", "BRCD"), "CODE", "NAME")
            M_TXNM = GetCode("TAXMST", GetCode("ACCMST", RECSET!PCOD, "CODE", "TXCD"), "CODE", "NAME")
        End If
                
        RECSET.MoveNext
    Loop
    RECSET.Close
    
    If chln <> Empty Then
        For i = 1 To lstChallan.ListItems.COUNT
         If lstChallan.ListItems(i) = chln Then lstChallan.ListItems(i).Checked = True
        Next
    End If
    
    If TXTCHLN <> Empty And txtParty <> Empty And M_BRNM = Empty Then
       If M_BRNM.Enabled Then M_BRNM.SetFocus
    End If
    
Call CalculateQtyAmt
Screen.MousePointer = vbNormal
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then KeyCode = 0: Exit Sub
  cmbSelection.Enabled = True: cmbSelection.SetFocus
End Sub

Private Sub FindSerial()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND NAME='" & cmbSelection.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not PKTYPRS.EOF Then
 M_DBCD = Trim(PKTYPRS!CODE & "")
End If

If M_DBCD <> Empty Then
 lblBill.Caption = GenVNO("SAL", M_DBCD)
End If

End Sub

Private Sub CalculateQtyAmt()
Dim i As Long
Dim QTY As Double
Dim PCS As Double
Dim AMT As Double
PCS = 0
QTY = 0
AMT = 0

For i = 1 To lstChallan.ListItems.COUNT
   If lstChallan.ListItems(i).Checked = True Then
        PCS = PCS + Val(lstChallan.ListItems(i).ListSubItems(5))
        QTY = QTY + Val(lstChallan.ListItems(i).ListSubItems(6))
        AMT = AMT + Val(lstChallan.ListItems(i).ListSubItems(8))
   End If
Next

TXTTPCS = nstr(PCS, 5, 0)
TXTTQTY = nstr(QTY, 12, 3)
TXTITOT = nstr(AMT, 12, 2)

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
  RS.Open "select * from config where comp='" & compPth & "' and vtyp='SAL' AND DBCD='" & TXCD & "'  AND UNIT='" & UNCD & "' order by srch", CN, adOpenKeyset, adLockPessimistic
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
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(1), " +AMT_02")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(1), " -AMT_02")
        End If
    End If
    If M_NICK(2) <> "" Then
        If M_OPER(2) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(2), " +AMT_03")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(2), " -AMT_03")
        End If
    End If
    
    If M_NICK(3) <> "" Then
        If M_OPER(3) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(3), " +AMT_04")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(3), " -AMT_04")
        End If
    End If
    
    If M_NICK(4) <> "" Then
        If M_OPER(4) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(4), " +AMT_05")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(4), " -AMT_05")
        End If
    End If
    
    If M_NICK(5) <> "" Then
        If M_OPER(5) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(5), " +AMT_06")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(5), " -AMT_06")
        End If
    End If
    
    If M_NICK(6) <> "" Then
        If M_OPER(6) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(6), " +AMT_07")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(6), " -AMT_07")
        End If
    End If
    
    If M_NICK(7) <> "" Then
        If M_OPER(7) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(7), " +AMT_08")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(7), " -AMT_08")
        End If
    End If
    
    If M_NICK(8) <> "" Then
        If M_OPER(8) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(8), " +AMT_09")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(8), " -AMT_09")
        End If
    End If
  Next
  If flexBTRM.Rows > 0 Then
    'O.k
   Else
    flexBTRM.Enabled = False
  End If
End Sub

Private Sub FillChargesPercent()
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset
If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT * FROM TAXMST WHERE NAME ='" & M_TXNM & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
  i = 0
  For i = 1 To flexBTRM.Rows
    J = 0
    For J = 0 To FINDRS.Fields.COUNT - 1
      If Trim(FINDRS.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i - 1, 0)) Then
         flexBTRM.TextMatrix(i - 1, 1) = FINDRS.Fields(J).Value
      End If
    Next
  Next
End If
FINDRS.Close
calBTRM 0
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
    If M_BILRDOF = "Y" Then
        TXTBNET.Text = Format(FormatNumber(Val(TXTITOT.Text) + Val(TXTADLS.Text), 0), "##########.00")
    Else
        TXTBNET.Text = Format(Val(TXTITOT.Text) + Val(TXTADLS.Text), "##########.00")
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
        TXTBNET.Text = Val(TXTITOT.Text) + subTot
    Next J

    
End Sub

Private Sub TXTITOT_Change()
  If flexBTRM.Rows > 0 Then
    flexBTRM.COL = 0
    flexBTRM.ROW = 0
  End If
  calBTRM 0
  Call calADLS
End Sub

Private Sub flexBTRM_GotFocus()
    Me.KeyPreview = False
    
    Msg "Billing Terms"
    If flexBTRM.Rows > 0 Then
      flexBTRM.COL = 1
      flexBTRM.TopRow = 0
      flexBTRM.LeftCol = 1
     Else
      TXTBNET = TXTITOT
    End If
End Sub


Private Sub TXTLRNO_GotFocus()
  TXTLRNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTLRNO_LostFocus()
 TXTLRNO.BackColor = vbWhite
End Sub

Private Sub txtDEST_GotFocus()
  txtDEST.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDEST_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or (Trim(txtDEST) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtDEST.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM CITYMASTER", 0, txtDEST.Text, "SELECT DESTINATION/CITY FROM LIST")
        If key_PressNew = True Then
          M_DESC = ""
          txtDEST = Empty
          frm_citymaster.Show
        End If
    End If
    
    Me.KeyPreview = True
End Sub

Private Sub txtDEST_LostFocus()
txtDEST.BackColor = vbWhite
End Sub

Private Sub LRDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub FindGrnDetail()

'RAW MATERIAL COST
RMCOST = 0
GRNITMRATE = 0

Dim GRNRS As ADODB.Recordset
Set GRNRS = New ADODB.Recordset


    'STEP:1
    If GRNRS.State = 1 Then GRNRS.Close
    GRNRS.Open "SELECT RATE FROM JOBIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND VTYP='IVR' AND DBCD='000002' AND VBNO='" & GRNNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
    If Not GRNRS.EOF Then
       GRNITMRATE = Val(GRNRS!RATE)
    End If
    GRNRS.Close
    '------------------------------------------------
            
    RMCOST = Val(TXTTQTY) * GRNITMRATE
    
    'STEP:2
    LESSOTHER = (12.36) * RMCOST * 0.01
        

End Sub

Private Function IsValidSelection() As Boolean
  IsValidSelection = True
  ISJOBBILL = False
  
  Dim JOB As Long, OTHER As Long
  Dim i As Long
    
  'Don't Mix Job Challan With Other Challan.
  JOB = 0: OTHER = 0
  
  For i = 1 To lstChallan.ListItems.COUNT
    If lstChallan.ListItems(i).Checked = True Then
       If Trim(lstChallan.ListItems(i).ListSubItems(15)) = "000003" Then
          JOB = JOB + 1
       Else
          OTHER = OTHER + 1
          Exit For
       End If
    End If
  Next
  
  If JOB > 0 And OTHER > 0 Then
     IsValidSelection = False
     ISJOBBILL = False
     Exit Function
  ElseIf JOB = 0 Then
     ISJOBBILL = False
  ElseIf JOB > 0 And OTHER = 0 Then
     ISJOBBILL = True
  End If
  '---------------------------------------------------------------
  
End Function

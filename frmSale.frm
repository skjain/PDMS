VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Module"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleMode       =   0  'User
   ScaleWidth      =   11415.32
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   7440
   End
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13361
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
         Left            =   960
         TabIndex        =   46
         Top             =   4440
         Width           =   1215
      End
      Begin VB.ComboBox cmbSelection 
         BackColor       =   &H00C0C0FF&
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
         Left            =   1680
         TabIndex        =   3
         Tag             =   "0"
         Text            =   "cmbSelection"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox TXTSEARCH 
         Height          =   285
         Left            =   6360
         MaxLength       =   30
         TabIndex        =   4
         Top             =   720
         Width           =   3495
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
         Left            =   5640
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   4440
         Width           =   1575
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
         TabIndex        =   36
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Frame FRMBTRM 
         Height          =   2535
         Left            =   7440
         TabIndex        =   30
         Top             =   4320
         Width           =   3855
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
            Text            =   "0.00"
            Top             =   2040
            Width           =   1905
         End
         Begin MSFlexGridLib.MSFlexGrid flexBTRM 
            Height          =   1875
            Left            =   60
            TabIndex        =   34
            Top             =   120
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   3307
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
            TabIndex        =   35
            Top             =   2160
            Width           =   1305
         End
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
         ItemData        =   "frmSale.frx":0000
         Left            =   1680
         List            =   "frmSale.frx":0002
         TabIndex        =   1
         Tag             =   "0"
         Text            =   "cmbPackingType"
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
         Format          =   56688641
         CurrentDate     =   38429
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1320
         TabIndex        =   40
         Top             =   6960
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
         Image           =   "frmSale.frx":0004
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2640
         TabIndex        =   41
         Top             =   6960
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
         Image           =   "frmSale.frx":0D8E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   3960
         TabIndex        =   42
         Top             =   6960
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
         Image           =   "frmSale.frx":11E0
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdGo 
         Height          =   375
         Left            =   9960
         TabIndex        =   5
         Top             =   720
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
         Image           =   "frmSale.frx":1632
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView lstChallan 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   5318
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
         NumItems        =   23
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
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Consinee Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Total Pcs"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ICOD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Item Desc."
            Object.Width           =   3528
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
            Text            =   "OrderNo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "DONO"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "ORDBCD"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "CHLNDBCD"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "TXCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "RTCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Retail/Tax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "JOB GRN"
            Object.Width           =   0
         EndProperty
      End
      Begin TabDlg.SSTab FRMLRDTL 
         Height          =   2055
         Left            =   120
         TabIndex        =   48
         Top             =   4800
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3625
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         ForeColor       =   128
         TabCaption(0)   =   "Transport Details"
         TabPicture(0)   =   "frmSale.frx":19CC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label8"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label9(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label10"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label11"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtPR"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "LRDT"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "TXTCRDS"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtLRNO"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtDEST"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtRMRK"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "  Export Details"
         TabPicture(1)   =   "frmSale.frx":19E8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CHKCIFRAT"
         Tab(1).Control(1)=   "TXTPORT"
         Tab(1).Control(2)=   "TXTMODE"
         Tab(1).Control(3)=   "TXTEXPTYP"
         Tab(1).Control(4)=   "LBLPORT"
         Tab(1).Control(5)=   "LBLMODE"
         Tab(1).Control(6)=   "LBLEXPTYP"
         Tab(1).ControlCount=   7
         Begin VB.ComboBox txtRMRK 
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
            Left            =   1200
            Style           =   1  'Simple Combo
            TabIndex        =   8
            Top             =   480
            Width           =   5895
         End
         Begin VB.TextBox txtDEST 
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   15
            Top             =   1200
            Width           =   3255
         End
         Begin VB.TextBox txtLRNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXTCRDS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   13
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox CHKCIFRAT 
            Caption         =   "CIF Value Required"
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
            Left            =   -70080
            TabIndex        =   21
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox TXTPORT 
            Height          =   285
            Left            =   -73080
            MaxLength       =   25
            TabIndex        =   25
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox TXTMODE 
            Height          =   285
            Left            =   -73080
            MaxLength       =   150
            TabIndex        =   23
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox TXTEXPTYP 
            Height          =   285
            Left            =   -73080
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   480
            Width           =   2055
         End
         Begin MSFlexGridLib.MSFlexGrid FLEXPLY 
            Height          =   765
            Left            =   -74880
            TabIndex        =   49
            Top             =   960
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   1349
            _Version        =   393216
            Cols            =   5
            BackColor       =   -2147483628
            BackColorBkg    =   -2147483633
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
         Begin MSComCtl2.DTPicker LRDT 
            Height          =   330
            Left            =   3840
            TabIndex        =   12
            Top             =   840
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
            Format          =   56688641
            CurrentDate     =   38429
         End
         Begin MSMask.MaskEdBox txtPR 
            Height          =   330
            Left            =   1680
            TabIndex        =   16
            Top             =   1560
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
         Begin VB.Label Label11 
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "No. of Ply"
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
            Index           =   2
            Left            =   -70320
            TabIndex        =   52
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "No. of Cops"
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
            Index           =   1
            Left            =   -72600
            TabIndex        =   51
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "No. of Pallets"
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
            Index           =   3
            Left            =   -74760
            TabIndex        =   50
            Top             =   480
            Width           =   1215
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00000080&
            Height          =   2655
            Left            =   -72840
            Shape           =   4  'Rounded Rectangle
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label10 
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
            TabIndex        =   14
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label9 
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
            Index           =   0
            Left            =   2760
            TabIndex        =   10
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label8 
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
            TabIndex        =   7
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Da&ys"
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
            TabIndex        =   11
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Remar&ks"
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
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
         Begin VB.Label LBLPORT 
            Caption         =   "Port/City of Loading :"
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
            Left            =   -74880
            TabIndex        =   24
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label LBLMODE 
            Caption         =   "Payment Terms"
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
            Left            =   -74880
            TabIndex        =   22
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label LBLEXPTYP 
            Caption         =   "ExportType"
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
            Left            =   -74880
            TabIndex        =   19
            Top             =   480
            Width           =   975
         End
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
         TabIndex        =   47
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label LBLNAME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3840
         TabIndex        =   45
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label LBLSEARCH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Criteria :"
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
         TabIndex        =   44
         Top             =   720
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000080&
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   615
         Left            =   105
         Top             =   585
         Width           =   11070
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
         TabIndex        =   43
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
         TabIndex        =   39
         Top             =   4440
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
         Left            =   4680
         TabIndex        =   38
         Top             =   4440
         Width           =   1215
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Dim TXRT As String
Dim TTYP As String
Dim SAVEFLAG As Boolean
Dim itemcode As String
Dim EXP_TYP_COD As String
Dim SAVDAT As ADODB.Recordset

Private Sub CHKCIFRAT_Click()
  If CHKCIFRAT.Value = 1 Then
    FRM_RATECIF.Show 1
  End If
End Sub

Private Sub CHKCIFRAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub cmbPackingType_Click()
    SendKeys "{HOME}"
    Call FindSerial
    Call FIL_Billingterm
End Sub

Private Sub cmbPackingType_GotFocus()
  SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
  TXTVBDT.Enabled = True: TXTVBDT.SetFocus
End Sub


Private Sub CMBSELECTION_Click()
SendKeys "{HOME}"
TXTSEARCH = Empty
Select Case cmbSelection.Text
Case "AgentWise"
      lblName.Caption = "Select Agent Name : "
Case "A/c PartyWise"
      lblName.Caption = "Select A/C Party Name : "
Case "ConsineeWise"
      lblName.Caption = "Select Consinee Name : "
End Select

End Sub

Private Sub cmbSelection_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
  TXTSEARCH.Enabled = True: TXTSEARCH.SetFocus
End Sub

Private Sub cmdCancel_Click()
    TXTCRDS = Empty
    TXTRMRK = Empty
    TXTSEARCH = Empty
    txtDEST = Empty: TXTLRNO = Empty
    TXTMODE = Empty: TXTEXPTYP = Empty: TXTPORT = Empty
    TXTTPCS = "0"
    TXTTQTY = "0.000"
    TXTITOT = "0.00"
    TXTADLS = "0.00"
    TXTBNET = "0.00"
    Call FindSerial
    TXCD = Empty
    TXRT = Empty
    Call FIL_Billingterm
    Call FillList
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub cmdGo_GotFocus()

If TXTSEARCH = Empty Then
   Call FillList
   Exit Sub
End If

Select Case cmbSelection.Text
Case "AgentWise"
      Call FillList(" AND SPTRAN.BRCD ='" & Trim(TXTSEARCH.Tag) & "'")
Case "A/c PartyWise"
      Call FillList(" AND SPTRAN.DRAC ='" & Trim(TXTSEARCH.Tag) & "'")
Case "ConsineeWise"
      Call FillList(" AND SPTRAN.DCOD ='" & Trim(TXTSEARCH.Tag) & "'")
End Select
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim i As Long, COUNTER As Long, J As Long
Dim SQL As String, SRCH As Long, DPFCODE As String
SRCH = 0

COUNTER = 0
For i = 1 To lstChallan.ListItems.COUNT
 If lstChallan.ListItems(i).Checked = True Then
  COUNTER = COUNTER + 1
  If COUNTER = 1 Then
  DPFCODE = Trim(lstChallan.ListItems(i).ListSubItems(18))
  DRAC = Trim(lstChallan.ListItems(i).ListSubItems(11))
  PCOD = Trim(lstChallan.ListItems(i).ListSubItems(11))
  DCOD = Trim(lstChallan.ListItems(i).ListSubItems(13))
  ADDRESS = Trim(lstChallan.ListItems(i).ListSubItems(14))
  BRCD = Trim(lstChallan.ListItems(i).ListSubItems(12))
  CPCD = GetCode("ACCMST", DRAC, "CODE", "CPCD")
  ARCD = GetCode("ACCMST", DRAC, "CODE", "ARCD")
  End If
 End If
Next

If COUNTER = 0 Then MsgBox "First Select Then Save Data": lstChallan.SetFocus: Exit Sub
TXTCRDS = Val(TXTCRDS)
COUNTER = 0

Dim SAVDAT As ADODB.Recordset
Set SAVDAT = New ADODB.Recordset
If SAVDAT.State = 1 Then SAVDAT.Close
SAVDAT.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DBCD='" & M_DBCD & "' AND VTYP='SAL' AND VBNO='" & Trim(lblBill.Caption) & "'", CN, adOpenDynamic, adLockOptimistic
If Not SAVDAT.EOF Then
   MsgBox "Bill No. Already Exist.Change Bill From Unit Configuration."
   cmdSave.SetFocus
   Exit Sub
End If

'Export Type Code
 EXP_TYP_COD = Empty
 If TXTEXPTYP.Enabled = True And Trim(TXTEXPTYP) <> Empty Then
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM EXPTYPMST WHERE NAME='" & TXTEXPTYP.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then EXP_TYP_COD = RS!CODE & "" '
 End If
'-----------------------------------------

Call FindSerial

CN.BeginTrans
For i = 1 To lstChallan.ListItems.COUNT
 If lstChallan.ListItems(i).Checked = True Then
 COUNTER = COUNTER + 1
 Dim RECSET As ADODB.Recordset
 Set RECSET = New ADODB.Recordset

SQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND RECSTAT='A' AND "
SQL = SQL & "VBNO = '" & Trim(lstChallan.ListItems(i)) & _
"' AND DBCD = '" & Trim(lstChallan.ListItems(i).ListSubItems(18)) & _
"' AND ICOD = '" & Trim(lstChallan.ListItems(i).ListSubItems(7)) & _
"' AND EXTRA1='" & Trim(lstChallan.ListItems(i).ListSubItems(15)) & "'"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

'REPORT PURPOSE TO SHOW BILL NO
SQL = "UPDATE ORDTRN SET VBNO='" & Trim(lblBill.Caption) & _
"',PKGS=PKGS + " & Val(RECSET!PCES) & ",COPS=COPS + " & Val(RECSET!COPS) & " WHERE COMP='" & compPth & _
"' AND UNIT = '" & UNCD & "' AND DVCD='" & DIVCODE & "' AND ORDN='" & RECSET!EXTRA1 & "' AND DONO='" & RECSET!EXTRA2 & _
"' AND DBCD='" & Trim(RECSET!EXTRA3) & "' AND RECSTAT='A'"
 
CN.Execute SQL

If COUNTER = 1 Then
  Call FindTax(Trim(lstChallan.ListItems(i).ListSubItems(15)))
End If

Set SAVDAT = New ADODB.Recordset

If Not RECSET.EOF Then
  '---------------------------------------------------------------------------------------
  '------------- EXCISE INFORMATION -------------------
  'Exciseable Commodity and chapter No.
  Dim EXCO As String
  Dim CHAP As String
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  
  If Trim(RECSET!ltno & "") = "" Or Trim(RECSET!ltno & "") = "WASTE" Then
     SAVDAT.Open "SELECT WEXCO AS EXCO,WCHAP AS CHAP FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
  Else
     SAVDAT.Open "SELECT EXCOMMODITY AS EXCO,CHAPTERNO AS CHAP FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & DIVCODE & "'", CN, adOpenDynamic, adLockOptimistic
  End If
     
  If Not SAVDAT.EOF Then
    EXCO = Trim(SAVDAT!EXCO & "")
    CHAP = Trim(SAVDAT!CHAP & "")
  Else
    EXCO = Empty
    CHAP = Empty
  End If
  SAVDAT.Close
  '---------------------------------------------------------------------------------------

End If
 
Do While Not RECSET.EOF
SRCH = SRCH + 1

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,SRCH,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,BRCD,"
SQL = SQL & "DCOD,ADDRESS,LTNO,TXRT,ICOD,GRAD,SUBGRD,PCES,GWGT,TWGT,QNTY,EXRATE,ARAT,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,VEHICALNO,TRCD,EXTRA1,EXTRA2,EXTRA3,TXCD,RTCD,RTYP,SDBC,SVBN,ORDN,GATEPASSNO)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','SAL','" & M_DBCD & "','" & SRCH & "','" & Trim(lblBill.Caption) & "','" & Trim(Trim(lstChallan.ListItems(i))) & _
"','" & Format(Trim(lstChallan.ListItems(i).ListSubItems(1)), "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & RECSET!DRAC & "','" & RECSET!PCOD & _
"','" & BRCD & "','" & RECSET!DCOD & "','" & ADDRESS & "','" & RECSET!ltno & "','" & RECSET!TXRT & "','" & RECSET!ICOD & "','" & RECSET!grad & _
"','" & RECSET!SUBGRD & "','" & RECSET!PCES & "','" & Val(RECSET!GWGT) & "','" & Val(RECSET!TWGT) & "','" & Val(RECSET!QNTY) & _
"'," & RECSET!EXRATE & "," & RECSET!ARAT & "," & RECSET!RATE & "," & RECSET!AMNT & _
",'Q','N','" & cUName & "','-','A','" & RECSET!COPS & "','" & Trim(RECSET!VEHICALNO) & "','" & RECSET!TRCD & "','" & RECSET!EXTRA1 & _
"','" & RECSET!EXTRA2 & "','" & RECSET!EXTRA3 & "','" & Trim(RECSET!TXCD & "") & "','" & Trim(RECSET!RTCD & "") & _
"','DPF','" & Trim(lstChallan.ListItems(i).ListSubItems(18)) & "','" & Trim(lstChallan.ListItems(i)) & _
"','" & Trim(RECSET!ORDN & "") & "','" & Trim(RECSET!GATEPASSNO) & "')"

CN.Execute SQL

RECSET.MoveNext
Loop
 
SQL = "UPDATE SPTRAN SET RTYP='SAL',SDBC='" & M_DBCD & "',SVBN='" & Trim(lblBill.Caption) & _
"' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND VTYP='DPF' AND RECSTAT='A' AND VBNO = '" & Trim(lstChallan.ListItems(i)) & _
"' AND DBCD = '" & Trim(lstChallan.ListItems(i).ListSubItems(18)) & "'"
   
CN.Execute SQL
 
TXCD = Trim(lstChallan.ListItems(i).ListSubItems(19))
TXRT = Trim(lstChallan.ListItems(i).ListSubItems(21))

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
  SAVDAT!TTYP = TXRT
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
  SAVDAT!LRNO = TXTLRNO
  SAVDAT!LRDT = Format(LRDT.Value, "YYYY/MM/DD")
  SAVDAT!DSTN = GetCode("CITYMASTER", txtDEST, "NAME", "CODE")
  SAVDAT!TRCD = ""
  SAVDAT!RECSTAT = "A"
  SAVDAT!BRMK = Trim(TXTRMRK.Text)
  
  SAVDAT!EXCO = EXCO
  SAVDAT!CHAP = CHAP
  SAVDAT!EXCSRNO = GetExciseSerial(DIVCODE)
  
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
  
    'EXPORT
        SAVDAT!exptypcod = EXP_TYP_COD
        SAVDAT!PAYMODE = Trim(TXTMODE.Text)
        SAVDAT!Port = Trim(TXTPORT.Text)
    'For CIF RATE
        SAVDAT!FOB_RATE = Val(FRM_RATECIF.FOB_RAT)
        SAVDAT!INS_RATE = Val(FRM_RATECIF.INS_RAT)
        SAVDAT!FRT_RATE = Val(FRM_RATECIF.FRT_RAT)
        SAVDAT!CIF_RATE = Val(FRM_RATECIF.CIF_RAT)
        SAVDAT!FOB_VALU = Val(FRM_RATECIF.FOB_VAL)
        SAVDAT!INS_VALU = Val(FRM_RATECIF.INS_VAL)
        SAVDAT!FRT_VALU = Val(FRM_RATECIF.FRT_VAL)
        SAVDAT!CIF_VALU = Val(FRM_RATECIF.CIF_VAL)
        SAVDAT!ADVANCE = Val(FRM_RATECIF.ADVN)
        SAVDAT!TOTPKG = GetTotalPallet
     '------------------------------------------------
  
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
  EXCISE!DCOD = DCOD
  EXCISE!ADDRESS = ADDRESS
  EXCISE!BRCD = BRCD
  EXCISE!CPCD = CPCD
  EXCISE!ARCD = ARCD
  EXCISE!TXCD = TXCD
  EXCISE!TAXGRP = GetCode("TAXMST", TXCD, "CODE", "GRPCOD")
  EXCISE!VBNO = lblBill
  EXCISE!chln = lblBill
  EXCISE!CHDT = Format(TXTVBDT, "YYYY/MM/DD")
  EXCISE!TRCD = ""
  EXCISE!LRNO = "" 'Trim(TXTLRNO.Text)
  EXCISE!ICOD = itemcode
  EXCISE!PCES = TXTTPCS
  EXCISE!QNTY = TXTTQTY
  EXCISE!AMNT = TXTITOT
  EXCISE!ITOT = TXTITOT
  EXCISE!BADJ = Val(TXTBNET.Text) - Val(TXTITOT.Text)
  EXCISE!BNET = TXTBNET
  EXCISE!RORT = TXRT
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
  
  '---------------------------------------------------------------------------------------
  '------------- BILLING TERM CHARGES -------------------
  '---------------------------------------------------------------------------------------
  'Add Records In Tranman
  Dim BOK_TOT As Double
  BOK_TOT = 0
  BOK_TOT = Val(TXTBNET)
  i = 0
  For i = 0 To flexBTRM.Rows - 1
    If M_POSTYESNO(i) = "Y" Then
      If M_OPER(i) = "+" Then
        BOK_TOT = BOK_TOT - Val(flexBTRM.TextMatrix(i, 2))
       Else
        BOK_TOT = BOK_TOT + Val(flexBTRM.TextMatrix(i, 2))
      End If
    End If
  Next


   
SQL = "UPDATE SERIALMASTER SET [SRNO]='" & lblBill & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
"' AND VTYP='SAL' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "'"
   
CN.Execute SQL

SQL = "UPDATE DIVMST SET [SRNO]='" & GetExciseSerial(DIVCODE) & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
      "' AND CODE='" & DIVCODE & "' "
   
CN.Execute SQL

'-----------------
'DAILYSTATUS ENTRY
  Call DAILYSTATUS("SAL", DRAC, M_DBCD, Val(TXTTQTY), lblBill, Val(TXTITOT), cUName, "N", Now, TXTVBDT)
'-------------------------------------------------------------------------------------------------------

CN.CommitTrans

MsgBox "Your Invoice No. is : " & lblBill.Caption
Call cmdCancel_Click
Call FillList

Call FillCombo("SELECT DISTINCT TOP 10 BRMK FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND DVCD='" & DIVCODE & "' AND VTYP='SAL' AND BRMK<>'' AND BRMK<>'0' AND BRMK<>'.' AND BRMK is not null ", TXTRMRK)
                 
Exit Sub

LAST:
MsgBox ERR.Description
CN.RollbackTrans
Exit Sub
End Sub

Private Sub flexBTRM_KeyDown(KeyCode As Integer, Shift As Integer)
If flexBTRM.Rows = 0 Then Exit Sub

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
   TXTLRNO.Enabled = True
   TXTLRNO.SetFocus
 End If
End If
End With
End Sub

Private Sub flexBTRM_KeyPress(KeyAscii As Integer)
If flexBTRM.Rows = 0 Then Exit Sub

With flexBTRM
If .COL = 1 Then '--------------------------------------

If InStr(1, flexBTRM.TextMatrix(flexBTRM.ROW, flexBTRM.COL), ".") > 0 And KeyAscii = 46 Then
   KeyAscii = 0
   Exit Sub
End If

If KeyAscii = 8 Then  'BACK SPACE
   Dim lnth As Double
   lnth = Len(flexBTRM.TextMatrix(flexBTRM.ROW, flexBTRM.COL))
    If lnth > 0 Then
      flexBTRM.TextMatrix(flexBTRM.ROW, flexBTRM.COL) = Mid(flexBTRM.TextMatrix(flexBTRM.ROW, flexBTRM.COL), 1, lnth - 1)
      calBTRM 0
      Call calADLS
      Exit Sub
    End If
End If

If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 46) Then             ' 0- 9
   flexBTRM.TextMatrix(flexBTRM.ROW, flexBTRM.COL) = (flexBTRM.TextMatrix(flexBTRM.ROW, flexBTRM.COL)) + Chr(KeyAscii)
   calBTRM 0
   Call calADLS
Else
   KeyAscii = 0
End If

End If '------------------------------------------------
End With
End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or DIVNAME = Empty Then
  Unload Me
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
  
  cmbSelection.Clear
  cmbSelection.AddItem ("A/c PartyWise")
  cmbSelection.AddItem ("AgentWise")
  cmbSelection.AddItem ("ConsineeWise")
  cmbSelection.ListIndex = 0
  
  Call SetSaleType
  TXTVBDT = Now
  LRDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  
  Call FillCombo("SELECT DISTINCT TOP 10 BRMK FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND DVCD='" & DIVCODE & "' AND VTYP='SAL' AND BRMK<>'' AND BRMK<>'0' AND BRMK<>'.' AND BRMK is not null ", TXTRMRK)
  
End Sub

Private Sub SetSaleType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND FYCD='" & FYCD & "' AND ACTIVE='Y'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop

If cmbPackingType.ListCount >= 1 Then cmbPackingType.ListIndex = 0

End Sub

Private Sub InDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub


Private Sub LRDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub lstChallan_Click()
Dim i As Long, COUNTER As Long, POS As Long
Dim SQL As String

COUNTER = 0
For i = 1 To lstChallan.ListItems.COUNT
 If lstChallan.ListItems(i).Checked = True Then
    COUNTER = COUNTER + 1
    If COUNTER = 1 Then POS = i
 End If
Next

If COUNTER = 0 Then
   TXCD = Empty
   Call FIL_Billingterm
   Call FillList
   Exit Sub
End If

If COUNTER = 1 Then
 
 SQL = SQL & " AND SPTRAN.DBCD ='" & Trim(lstChallan.ListItems(POS).ListSubItems(18)) & "' "
 SQL = SQL & " AND SPTRAN.DRAC ='" & Trim(lstChallan.ListItems(POS).ListSubItems(11)) & "' "
 SQL = SQL & " AND SPTRAN.DCOD ='" & Trim(lstChallan.ListItems(POS).ListSubItems(13)) & "' "
 SQL = SQL & " AND SPTRAN.ADDRESS ='" & Trim(lstChallan.ListItems(POS).ListSubItems(14)) & "' "
 SQL = SQL & " AND SPTRAN.BRCD ='" & GetCode("REFMST", Trim(lstChallan.ListItems(POS).ListSubItems(3)), "NAME", "CODE") & "' "
 SQL = SQL & " AND SPTRAN.TXCD ='" & Trim(lstChallan.ListItems(POS).ListSubItems(19)) & "' "
 TXCD = Trim(lstChallan.ListItems(POS).ListSubItems(19))
 SQL = SQL & " AND SPTRAN.RTCD ='" & Trim(lstChallan.ListItems(POS).ListSubItems(20)) & "' "
 If Trim(lstChallan.ListItems(POS).ListSubItems(21)) <> Empty Then 'EXEMPT IN CASE OF EXPORT
    SQL = SQL & " AND SPTRAN.TXRT ='" & Trim(lstChallan.ListItems(POS).ListSubItems(21)) & "' "
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
  TXTRMRK.Enabled = True
  TXTRMRK.SetFocus
End If
End Sub

Private Sub lstChallan_LostFocus()
  Call CalculateQtyAmt
End Sub

Private Sub lstChallan_Validate(Cancel As Boolean)
  Call CalculateQtyAmt
End Sub

Private Sub TimerBillNo1_Timer()
    Static ctr As Integer
    
    If ctr Mod 45 = 0 And ctr <= 45 Then
        lblAlert.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        Shape1.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        lblBill.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
    ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
        lblAlert.ForeColor = vbRed
        Shape1.BorderColor = vbRed
        lblBill.ForeColor = vbRed
    ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
        lblAlert.ForeColor = vbBlue
        Shape1.BorderColor = vbBlue
        lblBill.ForeColor = vbBlue
        ctr = 0
    End If
    
    ctr = ctr + 15
End Sub

Private Sub TXTCRDS_GotFocus()
  TXTCRDS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCRDS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TXTCRDS_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTCRDS, Me, False) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTCRDS_LostFocus()
  TXTCRDS.BackColor = vbWhite
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
    ElseIf KeyCode = vbKeyDelete Then
       txtDEST = Empty
    End If
    
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
    End If
    Me.KeyPreview = True
End Sub

Private Sub txtDEST_LostFocus()
txtDEST.BackColor = vbWhite
End Sub

Private Sub TXTLRNO_GotFocus()
  TXTLRNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtLRNO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TXTLRNO_LostFocus()
 TXTLRNO.BackColor = vbWhite
End Sub

Private Sub TXTMODE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TXTPORT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtPR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   FRMLRDTL.Tab = 1
   If TXTEXPTYP.Enabled Then TXTEXPTYP.SetFocus
End If
End Sub

Private Sub TXTRMRK_GotFocus()
   TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
   TXTRMRK.Height = 1155
   TXTRMRK.ZOrder
End Sub

Private Sub TXTRMRK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub TXTRMRK_LostFocus()
  TXTRMRK.BackColor = vbWhite
  TXTRMRK.Height = 325
End Sub

Public Sub FillList(Optional FILTER As String, Optional chln As String)
Dim SQL As String
Dim M_ROW As Integer
Dim Item
lstChallan.ListItems.Clear

Screen.MousePointer = vbHourglass
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT SPTRAN.VBNO,SPTRAN.ICOD,SPTRAN.DATE,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITEM,REFMST.NAME AS AGENT," & _
      "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS CADDRESS,SPTRAN.ADDRESS,SPTRAN.EXTRA1,SPTRAN.EXTRA2," & _
      "SPTRAN.EXTRA3,SPTRAN.EXTRA4,SPTRAN.DBCD,SPTRAN.TXCD,SPTRAN.RTCD,SPTRAN.DRAC,SPTRAN.BRCD,SPTRAN.TXRT," & _
      "SPTRAN.DCOD,SPTRAN.LRNO,SPTRAN.LRDT,SUM(ISNULL(QNTY,0)) AS QNTY," & _
      "SUM(ISNULL(PCES,0)) AS PCES,SUM(ISNULL(AMNT,0)) AS AMNT FROM SPTRAN " & _
      "INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.DRAC " & _
      "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND " & _
      "FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD " & _
      "INNER JOIN REFMST ON REFMST.CODE=SPTRAN.BRCD " & _
      "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS " & _
      "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & "' AND SPTRAN.DVCD='" & DIVCODE & _
      "' AND SPTRAN.VTYP='DPF' AND SPTRAN.DBCD NOT IN ('000004','000005') AND " & _
      "SPTRAN.RECSTAT='A' AND SVBN IS NULL AND LTrim(RTrim(IsNull(SPTRAN.EXTRA1,0)))<>'' "

'SALE DATE ARE GREATER THEN OR EQUAL TO CHALLAN DATE
SQL = SQL & " AND SPTRAN.DATE <= '" & Format(TXTVBDT.Value, "MM/DD/YYYY") & "' "
'---------------------------------------------------

If FILTER <> Empty Then SQL = SQL & FILTER

If InStr(1, UCase(cmbPackingType.Text), "COMMERCIAL TAX") <> 0 Then
   SQL = SQL & " AND SPTRAN.TXRT = 'TAX INVOICE' "
ElseIf InStr(1, UCase(cmbPackingType.Text), "COMMERCIAL RETAIL") <> 0 Then
   SQL = SQL & " AND SPTRAN.TXRT = 'RETAIL INVOICE' "
End If

If InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then
   SQL = SQL & " AND SPTRAN.DBCD ='000002' "
End If

SQL = SQL & " GROUP BY SPTRAN.VBNO,SPTRAN.ICOD,SPTRAN.DATE,ACCMST.NAME,FINITMMST.NAME,REFMST.NAME,PADDMST.NAME," & _
      "PADDMST.ADDR,SPTRAN.ADDRESS,SPTRAN.EXTRA1,SPTRAN.EXTRA2,SPTRAN.EXTRA3,SPTRAN.EXTRA4,SPTRAN.DBCD,SPTRAN.TXCD," & _
      "SPTRAN.RTCD,SPTRAN.DRAC,SPTRAN.BRCD,SPTRAN.TXRT,SPTRAN.DCOD,SPTRAN.LRNO,SPTRAN.LRDT"

TXTMODE = Empty
If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
Else
   If FILTER <> Empty Then TXTMODE = Trim(RECSET!EXTRA4 & "")
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
        Item.SubItems(3) = RECSET!AGENT
        Item.SubItems(4) = RECSET!CONSINEE
        Item.SubItems(5) = RECSET!PCES
        Item.SubItems(6) = nstr(RECSET!QNTY, 12, 3)
        Item.SubItems(7) = RECSET!ICOD & ""
        Item.SubItems(8) = RECSET!AMNT
        Item.SubItems(9) = RECSET!Item
        Item.SubItems(10) = RECSET!CADDRESS
        Item.SubItems(11) = RECSET!DRAC
        Item.SubItems(12) = RECSET!BRCD
        Item.SubItems(13) = RECSET!DCOD
        Item.SubItems(14) = RECSET!ADDRESS
        Item.SubItems(15) = Trim(RECSET!EXTRA1 & "")
        Item.SubItems(16) = Trim(RECSET!EXTRA2 & "")
        Item.SubItems(17) = Trim(RECSET!EXTRA3 & "")
        Item.SubItems(18) = Trim(RECSET!dbcd & "")
        Item.SubItems(19) = Trim(RECSET!TXCD & "")
        Item.SubItems(20) = Trim(RECSET!RTCD & "")
        Item.SubItems(21) = Trim(RECSET!TXRT & "")
        RECSET.MoveNext
    Loop
    RECSET.Close
    
    Dim i As Long
    If chln <> Empty Then
        For i = 1 To lstChallan.ListItems.COUNT
         If lstChallan.ListItems(i) = chln Then lstChallan.ListItems(i).Checked = True
        Next
    End If
    
Call CalculateQtyAmt
Screen.MousePointer = vbNormal
End Sub

Private Sub TXTSEARCH_GotFocus()
TXTSEARCH.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then TXTSEARCH = Empty: FillList: Exit Sub

If KeyCode = vbKeyF2 Then
Select Case cmbSelection.Text
Case "AgentWise"
      lblName.Caption = "Select Agent Name : "
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTSEARCH = SearchList1("SELECT TOP 20 CODE, NAME FROM REFMST WHERE CATA='B'", 0, TXTSEARCH.Text, "SELECT AGENT FROM LIST")
      TXTSEARCH.Tag = Key
Case "A/c PartyWise"
      lblName.Caption = "Select A/C Party Name : "
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTSEARCH = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST", 0, TXTSEARCH.Text, "SELECT A/C PARTY FROM LIST")
      TXTSEARCH.Tag = Key
Case "ConsineeWise"
      lblName.Caption = "Select Consinee Name : "
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTSEARCH = SearchList1("SELECT TOP 20 CODE, NAME FROM PADDMST", 0, TXTSEARCH.Text, "SELECT CONSINEE NAME FROM LIST")
      TXTSEARCH.Tag = Key
End Select
End If

If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TXTSEARCH_LostFocus()
  TXTSEARCH.BackColor = vbWhite
End Sub

Private Sub TXTVBDT_Change()
   Call FillList
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  cmbSelection.Enabled = True
  cmbSelection.SetFocus
End If
End Sub

Private Sub FindSerial()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND NAME='" & cmbPackingType.Text & "'", CN, adOpenDynamic, adLockOptimistic
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
    
    TXTBNET.Text = Format(FormatNumber(Val(TXTITOT.Text) + Val(TXTADLS.Text), 0), "##########.00")
    
End Sub


Private Sub calBTRM(ByVal ICTR As Integer)
Dim SAVEFALG As Boolean
SAVEFLAG = True
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
        If Val(flexBTRM.TextMatrix(J, 1)) = 0 Then flexBTRM.TextMatrix(J, 2) = 0
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
        
        TXTBNET.Text = Format(FormatNumber(Val(TXTITOT.Text) + subTot, 0), "##########.00")
        
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

Private Sub FindTax(ORDNO As String)
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset

If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT TXCD,RTCD,ICOD FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND RECSTAT='A' AND ORDN = '" & ORDNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
  TXCD = Trim(FINDRS!TXCD & "")
  TTYP = Trim(FINDRS!RTCD & "")
  itemcode = Trim(FINDRS!ICOD & "")
End If
FINDRS.Close
End Sub

Private Sub TXTEXPTYP_GotFocus()
 TXTEXPTYP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTEXPTYP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTEXPTYP.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTEXPTYP.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM EXPTYPMST WHERE RECSTAT='A'", 0, TXTEXPTYP, "List of Export Type Master")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTEXPTYP.Text = ""
            frmExportTypeMaster.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTEXPTYP = Empty
    ElseIf KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TXTEXPTYP_LostFocus()
 TXTEXPTYP.BackColor = vbWhite
End Sub

Private Sub TXTMODE_GotFocus()
 TXTMODE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTMODE_LostFocus()
 TXTMODE.BackColor = vbWhite
End Sub

Private Sub TXTPORT_GotFocus()
 TXTPORT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPORT_LostFocus()
 TXTPORT.BackColor = vbWhite
End Sub

Private Function GetTotalPallet() As Long
GetTotalPallet = 0
Dim TMPRS As ADODB.Recordset
Set TMPRS = New ADODB.Recordset
Dim SQL As String

SQL = "SELECT DISTINCT PLTNO FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND RECSTAT<>'D' AND " & _
      "RVBNO IN (SELECT DISTINCT VBNO FROM SPTRAN " & _
      "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND DVCD='" & DIVCODE & "' AND RTYP='SAL' AND SDBC='" & M_DBCD & _
      "' AND SVBN='" & Trim(lblBill.Caption) & "' AND RECSTAT<>'D') AND " & _
      "RDBC IN (SELECT DISTINCT DBCD FROM SPTRAN " & _
      "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND DVCD='" & DIVCODE & "' AND RTYP='SAL' AND SDBC='" & M_DBCD & _
      "' AND SVBN='" & Trim(lblBill.Caption) & "' AND RECSTAT<>'D')"

If TMPRS.State = 1 Then TMPRS.Close
TMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not TMPRS.EOF Then
   GetTotalPallet = Val(TMPRS.RecordCount)
End If
TMPRS.Close
End Function


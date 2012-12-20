VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmCaptiveBoxChallan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captive Challan ( Box Dispatch )"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11415
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   45
      Top             =   6960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrTool 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1920
         Top             =   0
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
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   6840
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   7035
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12409
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
      Begin VB.OptionButton optParty 
         BackColor       =   &H0080C0FF&
         Caption         =   "For Jobwork"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton optInternal 
         BackColor       =   &H0080C0FF&
         Caption         =   "For Internal"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox TXTMACHINE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox TXTSTKQTY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   20
         Text            =   ".000"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox TXTITM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox TXTGRAD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TXTRATE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         MaxLength       =   200
         TabIndex        =   19
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox BRMK 
         Height          =   285
         Left            =   8160
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox txtLTNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox TXTIGRP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox TXTINAM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox TXTFROMDIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox TXTTODIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   9120
         TabIndex        =   11
         Top             =   840
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
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
         Format          =   57081857
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   3480
         TabIndex        =   0
         Top             =   6240
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
         Image           =   "frmCaptiveBoxChallan.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   6720
         TabIndex        =   3
         Top             =   6240
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
         Image           =   "frmCaptiveBoxChallan.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   7800
         TabIndex        =   4
         Top             =   6240
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
         Image           =   "frmCaptiveBoxChallan.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   4560
         TabIndex        =   1
         Top             =   6240
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
         Image           =   "frmCaptiveBoxChallan.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5640
         TabIndex        =   2
         Top             =   6240
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
         Image           =   "frmCaptiveBoxChallan.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   8880
         TabIndex        =   5
         Top             =   6240
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
         Image           =   "frmCaptiveBoxChallan.frx":1CAA
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView lstBox 
         Height          =   2535
         Left            =   240
         TabIndex        =   22
         Top             =   3480
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4471
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
         NumItems        =   10
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
            Text            =   "Grade"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Pallet No."
            Object.Width           =   0
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   9840
         TabIndex        =   49
         Top             =   6240
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
         Image           =   "frmCaptiveBoxChallan.frx":20FC
         cBack           =   -2147483633
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   11280
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label LBLMAC 
         BackStyle       =   0  'Transparent
         Caption         =   "To Machine"
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
         Left            =   1080
         TabIndex        =   48
         Tag             =   "S"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Qnty"
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
         Left            =   6600
         TabIndex        =   47
         Tag             =   "S"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   8040
         X2              =   8040
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Captive Challan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   0
         Width           =   4455
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   705
         Left            =   3360
         Shape           =   4  'Rounded Rectangle
         Top             =   6120
         Width           =   7935
      End
      Begin VB.Label Label19 
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
         Left            =   240
         TabIndex        =   43
         Top             =   6240
         Width           =   885
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Net Wt"
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
         TabIndex        =   42
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   41
         Top             =   6240
         Width           =   900
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
         Left            =   1200
         TabIndex        =   40
         Top             =   6480
         Width           =   900
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
         Left            =   2160
         TabIndex        =   39
         Top             =   6480
         Width           =   1095
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
         Left            =   240
         TabIndex        =   38
         Top             =   6480
         Width           =   885
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   705
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6120
         Width           =   3255
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   9120
         TabIndex        =   37
         Tag             =   "S"
         Top             =   2640
         Width           =   975
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000080&
         X1              =   6600
         X2              =   6600
         Y1              =   1800
         Y2              =   2280
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   5160
         X2              =   5160
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         X1              =   3840
         X2              =   3840
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   6360
         X2              =   6360
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Item Description"
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
         Left            =   3360
         TabIndex        =   36
         Tag             =   "S"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label3 
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
         Left            =   2880
         TabIndex        =   35
         Tag             =   "S"
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label15 
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
         Left            =   5400
         TabIndex        =   34
         Tag             =   "S"
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label14 
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
         Left            =   4080
         TabIndex        =   33
         Tag             =   "S"
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label13 
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
         Left            =   720
         TabIndex        =   32
         Tag             =   "S"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Captive Challan No. :"
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
         Left            =   6840
         TabIndex        =   31
         Top             =   480
         Width           =   2295
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   6720
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   4095
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
         Left            =   9120
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   3255
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   11175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Captive Challan Date"
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
         Left            =   6840
         TabIndex        =   10
         Top             =   840
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   6120
         X2              =   6600
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Group"
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
         TabIndex        =   29
         Tag             =   "S"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
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
         TabIndex        =   28
         Tag             =   "S"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "--->>"
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
         Left            =   6600
         TabIndex        =   27
         Tag             =   "S"
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "----->"
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
         Left            =   6600
         TabIndex        =   26
         Tag             =   "S"
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "From Division :"
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
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LBLTODIV 
         BackStyle       =   0  'Transparent
         Caption         =   "To Division     :"
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
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCaptiveBoxChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ITM_CONVERSION As Double
Dim FDVCD As String
Dim TDVCD As String
Dim SQL As String
Dim SAVEFLAG As Boolean
Dim M_DBCD As String
Dim MCOD As String
Public RAWITMGRP As String
Dim RAWITM As String
Dim SITEM As String
Dim SGRD As String
Dim SUBGRD As String
Dim ALLOWEDITDEL As Boolean
Public CHALLAN As String
Dim TTLCOPS As Long
Dim TTLBOXES As Long
Dim TTLQTY As Double

Private Sub cmdCancel_Click()
  
  TXTFROMDIV.Tag = TXTFROMDIV
  TXTTODIV.Tag = TXTTODIV
  Call ClsData(Me)
  TXTFROMDIV = TXTFROMDIV.Tag
  TXTTODIV = TXTTODIV.Tag
  
  txtCTRN = Empty: txtCops = Empty: txtNTWT = Empty
  TXTSTKQTY = ".000"
  Call btn_sts(True)
  lstBox.ListItems.Clear
    
  If zoomflag = True Then
     Call CMDEXIT_Click
     Exit Sub
  End If
  TXTVBDT = Now
  lblBill.Caption = GenDPFVNO("DPF", M_DBCD, FDVCD)
  If cmdExit.Enabled Then cmdExit.SetFocus
End Sub

Private Sub cmdDelete_Click()
  ALLOWEDITDEL = True
  SAVEFLAG = False
  CHALLAN = Empty
  btn_sts (True)
  
  TXTFROMDIV.Enabled = True
  TXTTODIV.Enabled = True
        
    If FDVCD = TDVCD Then
       MsgBox "Invalid Selection of Division"
       TXTFROMDIV.Enabled = True
       TXTFROMDIV.SetFocus
       Exit Sub
    End If

    If TXTFROMDIV = Empty Then
       If TXTFROMDIV.Enabled Then TXTFROMDIV.SetFocus
       Exit Sub
    End If
    
    If TXTTODIV = Empty Then
       If TXTTODIV.Enabled Then TXTTODIV.SetFocus
       Exit Sub
    End If
    
    If FDVCD = Empty Or TDVCD = Empty Then
       If TXTFROMDIV.Enabled Then TXTFROMDIV.SetFocus
       Exit Sub
    End If
       
    frmCaptiveBoxChlnList.FDVCD = FDVCD
    frmCaptiveBoxChlnList.TDVCD = TDVCD
    frmCaptiveBoxChlnList.M_DBCD = M_DBCD
    
    CHALLAN = Empty
    
    frmCaptiveBoxChlnList.Show 1
  
  If ALLOWEDITDEL = False Then
    MsgBox "Sale Bill has been made can not edit/delete ", vbInformation
   Else
    If Not CHALLAN = Empty Then
      Dim AYS
      AYS = MsgBox("Are you sure to delete this Captive Challan ", vbYesNo)
      If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE SPTRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO='" & CHALLAN & "'"
        CN.Execute "UPDATE PKGMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND SLIPNO='" & CHALLAN & "'"
        CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO='" & CHALLAN & "'"
        
        'IN CASE OF AUTO RGP
        CN.Execute "DELETE FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND  VTYP='ANX' AND TRCD='000004' AND CHLN='" & CHALLAN & "' AND RECSTAT<>'D'"
        '===============================
                
        SQL = "UPDATE BOXREGISTER SET VTYP='PPF',RVBNO=NULL,RVBDT= NULL,RDBC = NULL,RVTYP = NULL WHERE COMP='" & compPth & _
        "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & "' AND RVBNO='" & CHALLAN & "' AND RDBC = '" & M_DBCD & _
        "' AND RVTYP='DPF' AND VTYP='DPF'"

        CN.Execute SQL

        'Call DAILYSTATUS("DPF", MCOD, M_DBCD, Val(txtNTWT), lblBill, 0, cUName, "D", Now, TXTVBDT)
        CN.CommitTrans
        MsgBox "Data Successfully Deleted."
      End If
    End If
  End If
  Call cmdCancel_Click
  If cmdAdd.Enabled Then cmdAdd.SetFocus
End Sub


Private Sub cmdEdit_Click()
    SAVEFLAG = False
    ALLOWEDITDEL = False
    TXTFROMDIV.Enabled = True
    TXTTODIV.Enabled = True
    
    
    If FDVCD = TDVCD Then
       MsgBox "Invalid Selection of Division"
       TXTFROMDIV.Enabled = True
       TXTFROMDIV.SetFocus
       Exit Sub
    End If

    If TXTFROMDIV = Empty Then
       If TXTFROMDIV.Enabled Then TXTFROMDIV.SetFocus
       Exit Sub
    End If
    
    If TXTTODIV = Empty Then
       If TXTTODIV.Enabled Then TXTTODIV.SetFocus
       Exit Sub
    End If
    
    If FDVCD = Empty Or TDVCD = Empty Then
       If TXTFROMDIV.Enabled Then TXTFROMDIV.SetFocus
       Exit Sub
    End If
       
    frmCaptiveBoxChlnList.FDVCD = FDVCD
    frmCaptiveBoxChlnList.TDVCD = TDVCD
    frmCaptiveBoxChlnList.M_DBCD = M_DBCD
    
    CHALLAN = Empty
    
    frmCaptiveBoxChlnList.Show 1
    
    If CHALLAN = Empty Or CHALLAN = "" Then
        btn_sts (True)
        cmdAdd.Enabled = True
        cmdAdd.SetFocus
    Else
        btn_sts (False)
        TXTFROMDIV.Enabled = True
        TXTFROMDIV.SetFocus
        Call FindStock
    End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim INDEX As Long
Dim FLAG As Boolean
Dim SLIP As String
Dim COPS As Double
Dim PCS As Double


FLAG = False
For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then: FLAG = True: Exit For
Next

If FLAG = False Then Exit Sub

If INVALIDDATA Then Exit Sub

Call SetInternal

'FOR YARN EMBROIDERY
Dim ISYARNEMB As Boolean
ISYARNEMB = False

If ISYARNEMB Then
   Call SETDIVISION
   Exit Sub
End If

'-----------------------

If SAVEFLAG Then
   Dim NSQL As String
   Dim MSGS As String: MSGS = "Unit"
   SLIP = GenDPFVNO("DPF", M_DBCD, FDVCD)
   
   NSQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & _
           "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & SLIP & "' "
   
   If UNT_DIVSERIES_REQ = "Y" Then
      NSQL = NSQL & " AND DVCD='" & FDVCD & "' "
      MSGS = "Division"
   End If
   
   If RS.State Then RS.Close
   RS.Open NSQL, CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      MsgBox "Captive Challan No. " & SLIP & " Already Exist. Check Last No. In " & MSGS & " Configuration", vbCritical
      Exit Sub
   End If
   RS.Close
End If


CN.BeginTrans

If SAVEFLAG Then

SLIP = GenDPFVNO("DPF", M_DBCD, FDVCD)

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA1,EXTRA2,EXTRA3)VALUES('" & compPth & "','" & UNCD & "','" & FDVCD & _
"','DPF','" & M_DBCD & "','" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','XXXXXX','" & MCOD & "','" & txtLTNo & _
"','" & SITEM & "','" & SGRD & "','" & SUBGRD & "','" & TTLBOXES & "','" & TTLQTY & _
"'," & TXTRATE & "," & (TTLQTY * Val(TXTRATE)) & ",'Q','N','" & cUName & "','*','A','" & TTLCOPS & _
"','" & BRMK & "','" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "','" & TDVCD & "')"

CN.Execute SQL

SQL = "INSERT INTO PKGMAN (COMP,UNIT,DVCD,DBCD,VTYP,DATE,SLIPNO,PKG_STCOD,"
SQL = SQL & "LOTNO,FINITMCOD,GRAD,SUBGRAD,NOB,QNTY,SYSR,[USER],OPER,RECSTAT) VALUES "
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & FDVCD & "','" & M_DBCD & "','DPF'"
SQL = SQL & ",'" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & SLIP & "','000000',"
SQL = SQL & "'" & txtLTNo & "','" & SITEM & "','" & SGRD & "','" & SUBGRD & "','" & TTLBOXES & "','" & TTLQTY & _
"','N','" & cUName & "','-','A')"

CN.Execute SQL

If optInternal.Value = True Then
    SQL = "INSERT INTO STORETRAN(COMP,UNIT,SRNO,DVCD,DBCD,VTYP,VBNO,CHLN,CHDT,[DATE],PCOD,ICOD,PCES,"
    SQL = SQL & "QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,LTNO,GRAD,SUBGRD,COPS,RECSTAT)VALUES"
    SQL = SQL & "('" & compPth & "','" & UNCD & "','" & FDVCD & "','" & TDVCD & "','" & M_DBCD & "','DPF'"
    SQL = SQL & ",'" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
    "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & MCOD & "','" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "','" & TTLBOXES & "','" & TTLQTY & _
    "'," & TXTRATE & "," & (TTLQTY * Val(TXTRATE)) & ",'Q','N','" & cUName & "','+','" & txtLTNo & _
    "','" & SGRD & "','" & SUBGRD & "','" & TTLCOPS & "','A')"
    
    CN.Execute SQL
    
End If

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & SLIP & "',RVBDT= '" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "',RDBC = '" & M_DBCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & _
   "' AND PKG_STCOD='" & Trim(lstBox.ListItems(INDEX).SubItems(8)) & "' AND (VTYP='PPF' OR VTYP='OPN') AND " & _
   "VBNO='" & lstBox.ListItems(INDEX).Text & "'"
   
   CN.Execute SQL
   
 End If
Next INDEX

Dim UPSQL As String
UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & SLIP & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
         "' AND VTYP='DPF' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "' "

If UNT_DIVSERIES_REQ = "Y" Then
   UPSQL = UPSQL & " AND DVCD='" & FDVCD & "' "
End If
 
CN.Execute UPSQL

'RGP
If optInternal.Value = False Then
   UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & SLIP & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
         "' AND VTYP='ANX' AND CODE='000001' AND FYCD='" & FYCD & "' "
   CN.Execute UPSQL
End If

Else

SLIP = lblBill.Caption

SQL = "UPDATE SPTRAN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',PCOD='" & MCOD & _
"',ICOD='" & SITEM & "',PCES='" & TTLBOXES & "',QNTY='" & TTLQTY & "',RATE=" & TXTRATE & _
",AMNT=" & (TTLQTY * Val(TXTRATE)) & ",GRAD='" & SGRD & "',LTNO='" & txtLTNo & "',COPS='" & TTLCOPS & _
"',EXTRA1='" & BRMK & "',EXTRA2='" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "',EXTRA3='" & TDVCD & _
"' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & _
"' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & SLIP & "'"
   
CN.Execute SQL

SQL = "UPDATE PKGMAN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & _
"',LOTNO='" & txtLTNo & "',FINITMCOD='" & SITEM & "',GRAD='" & SGRD & "',SUBGRAD='" & SUBGRD & _
"',NOB='" & TTLBOXES & "',QNTY='" & TTLQTY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & FDVCD & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND  SLIPNO = '" & SLIP & "'"

CN.Execute SQL

If optInternal.Value = True Then
    SQL = "UPDATE STORETRAN SET DATE='" & Format(TXTVBDT, "YYYY/MM/DD") & "',PCOD='" & MCOD & _
    "',ICOD ='" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "',PCES='" & TTLBOXES & _
    "',QNTY='" & TTLQTY & "',RATE=" & TXTRATE & _
    ",AMNT=" & (TTLQTY * Val(TXTRATE)) & ",LTNO='" & txtLTNo & "',GRAD='" & SGRD & "',SUBGRD='" & SUBGRD & _
    "',COPS ='" & TTLCOPS & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & TDVCD & _
    "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND  VBNO = '" & SLIP & "' AND RECSTAT='A'"
    
    CN.Execute SQL
End If

Dim L As Long
SQL = "UPDATE BOXREGISTER SET VTYP='PPF',RVBNO=NULL,RVBDT= NULL,RDBC = NULL,RVTYP = NULL WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & "' AND RVBNO='" & SLIP & "' AND RDBC = '" & M_DBCD & _
"' AND RVTYP='DPF' AND VTYP='DPF'"

CN.Execute SQL, L

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & SLIP & "',RVBDT= '" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "',RDBC = '" & M_DBCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & _
   "' AND PKG_STCOD='" & Trim(lstBox.ListItems(INDEX).SubItems(8)) & "' AND (VTYP='PPF' OR VTYP='OPN') " & _
   "  AND VBNO='" & lstBox.ListItems(INDEX).Text & "'"
   CN.Execute SQL, L
 End If
Next INDEX

End If
 
 If optInternal.Value = False Then
    'RGP SETTING=====================
    
     CN.Execute "DELETE FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                "' AND  VTYP='ANX' AND TRCD='000004' AND CHLN='" & SLIP & "' AND RECSTAT<>'D'"
     
     CN.Execute "INSERT INTO JOBOUT([COMP], [UNIT], [VTYP],[TRCD], [DBCD], [CHLN],[VBNO], [SRCH], [DATE], [PCOD], " & _
                "[ICOD], [IDNO], [PCES], [QNTY], [COPS], [LTNO],[RATE], [AMNT], [OPER]," & _
                " [GRNNO], [XDAYS],[RMRK], [USER], [SYSR],[RECSTAT]) VALUES('" & compPth & _
                "','" & UNCD & "','ANX','000004','000001','" & SLIP & "','" & SLIP & _
                "','1','" & Format(TXTVBDT, "YYYY/MM/DD") & _
                "','" & GetCode("ACCMST", TXTTODIV, "NAME", "CODE") & "','" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & _
                "','" & SLIP & "','" & TTLBOXES & "','" & (TTLQTY * ITM_CONVERSION) & "','" & TTLCOPS & _
                "','" & txtLTNo & "','" & Val(TXTRATE) & "','" & (Val(TXTRATE) * TTLQTY * ITM_CONVERSION) & _
                "','+','AUTO',0,'" & BRMK & "','" & cUName & "','N','A')"
                
     '=============================
 End If
 
'---------------------------
'DAILYSTATUS ENTRY
 If SAVEFLAG = True Then
  Call DAILYSTATUS("DPF", SITEM, M_DBCD, Val(txtNTWT), SLIP, 0, cUName, "N", Now, TXTVBDT)
  Else
  Call DAILYSTATUS("DPF", SITEM, M_DBCD, Val(txtNTWT), SLIP, 0, cUName, "M", Now, TXTVBDT)
 End If
'---------------------------
CN.CommitTrans

If SAVEFLAG Then
  MsgBox "Your Captive Challan No. is : " & SLIP
Else
  MsgBox "Your Captive Challan No. : " & SLIP & " is Successfully Edited."
End If

Call cmdCancel_Click

Exit Sub
LAST:
MsgBox ERR.Description
Exit Sub
End Sub

Private Sub cmdSavePrint_Click()
Call cmdSave_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If UCase(ActiveControl.NAME) = "TXTTODIV" And Not SAVEFLAG And TXTFROMDIV <> Empty And TXTTODIV <> Empty And txtitm = Empty And ALLOWEDITDEL = False Then
  Call cmdEdit_Click
  Exit Sub
ElseIf UCase(ActiveControl.NAME) = "TXTTODIV" And Not SAVEFLAG And TXTFROMDIV <> Empty And TXTTODIV <> Empty And txtitm = Empty And ALLOWEDITDEL = True Then
  Call cmdDelete_Click
  Exit Sub
End If

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
  ALLOWEDITDEL = False
  M_DBCD = "000004"
   
  TXTVBDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
    
    If zoomflag = True Then
        btn_sts (False)
        SAVEFLAG = False
    Else
        btn_sts (True)
    End If
End Sub

Private Sub cmdAdd_Click()
    zoomflag = False
    btn_sts (False)
    TXTFROMDIV.SetFocus
    SAVEFLAG = True
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub lstBox_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Msg ("Press F8 For Select All F9 For De-Select All")
Dim i As Integer, J As Integer, ctr As Integer
    If Item.Checked = True Then
        txtCTRN.Caption = Val(txtCTRN.Caption) + 1
        txtNTWT.Caption = nstr(Val(txtNTWT.Caption) + Val(Item.SubItems(2)), 10, 3)
        txtNTWT.Caption = Trim(txtNTWT.Caption)
        txtCops.Caption = Val(txtCops.Caption) + Val(Item.SubItems(1))
    Else
        txtCTRN.Caption = Val(txtCTRN.Caption) - 1
        txtNTWT.Caption = nstr(Val(txtNTWT.Caption) - Val(Item.SubItems(2)), 10, 3)
        txtNTWT.Caption = Trim(txtNTWT.Caption)
        txtCops.Caption = Val(txtCops.Caption) - Val(Item.SubItems(1))
    End If
    
    If Item.INDEX < lstBox.ListItems.COUNT Then lstBox.ListItems.Item(Item.INDEX + 1).Selected = True: lstBox.ListItems(Item.INDEX + 1).EnsureVisible
End Sub


Private Sub optInternal_Click()
Call setStatus
End Sub

Private Sub optParty_Click()
  Call setStatus
End Sub

Private Sub TXTFROMDIV_GotFocus()
   TXTFROMDIV.BackColor = RGB(BRED, BGREEN, BBLUE)
   SendKeys "{HOME}+{END}"
If TXTFROMDIV = Empty Then
   ToolTip Me, "Press {F2} / {Enter} For Division Master Help", "", TXTFROMDIV.Left, TXTFROMDIV.Top + TXTFROMDIV.Height + 100
Else
   ToolTip Me, "Press {F2} For Division Master Help", "", TXTFROMDIV.Left, TXTFROMDIV.Top + TXTFROMDIV.Height + 100
End If
End Sub

Private Sub TXTFROMDIV_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
          
   If KeyCode = vbKeyF2 Or (Trim(TXTFROMDIV) = Empty And KeyCode = vbKeyReturn) Then
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTFROMDIV.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, TXTFROMDIV.Text, "SELECT DIVISION FROM LIST")
      TXTFROMDIV.Tag = Key
      FDVCD = Key
      lblBill.Caption = GenDPFVNO("DPF", M_DBCD, FDVCD)
   End If
    
  Me.KeyPreview = True
End Sub

Private Sub TXTFROMDIV_LostFocus()
  TXTFROMDIV.BackColor = vbWhite
  picToolTip.Visible = False
End Sub

Private Sub TXTGRAD_Change()
If SAVEFLAG Then lstBox.ListItems.Clear: Call GenerateBoxList
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

Private Sub txtIGRP_GotFocus()
txtIGRP.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
  
  If txtIGRP = Empty Then
      ToolTip Me, "Press {F2}/{Enter} For Item Group Master Help", "", txtIGRP.Left, txtIGRP.Top + txtIGRP.Height + 100
  Else
      ToolTip Me, "Press {F2} For Item Group Master Help", "", txtIGRP.Left, txtIGRP.Top + txtIGRP.Height + 100
  End If
End Sub

Private Sub TXTIGRP_KeyDown(KeyCode As Integer, Shift As Integer)
    M_DESC = Empty
    Key = Empty
    sTxt = ""
    If KeyCode = vbKeyF2 Or (txtIGRP = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        txtIGRP = SearchList1("select TOP 20 code, name from IGMMST", 0, "", "List Of Item Group")
        RAWITMGRP = Key
    End If
    If key_PressNew = True Then
        M_DESC = ""
        FRM_IGRP.Show
    End If
End Sub

Private Sub txtIGRP_LostFocus()
txtIGRP.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub txtINAM_GotFocus()
TXTINAM.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
  
  If TXTINAM = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Item Master Help(RAW)", "", TXTINAM.Left, TXTINAM.Top + TXTINAM.Height + 100
  Else
      ToolTip Me, "Press {F2} For Item Master Help(RAW)", "", TXTINAM.Left, TXTINAM.Top + TXTINAM.Height + 100
  End If
End Sub

Private Sub TXTINAM_KeyDown(KeyCode As Integer, Shift As Integer)
If RAWITMGRP = Empty Or RAWITMGRP = "" Then MsgBox "Select Item Group First then Item.", txtIGRP.Enabled = True: txtIGRP.SetFocus: Exit Sub
   If KeyCode = vbKeyF2 Or (TXTINAM = Empty And vbKeyReturn) Then
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    TXTINAM = SearchList1("select TOP 20 code,name from itmmst WHERE IGCD='" & RAWITMGRP & "'", 0, TXTINAM, "SELECT ITEM FROM LIST")
    If key_PressNew = True Then
       M_DESC = ""
       Key = ""
       TXTINAM = ""
       frm_Item.Show
    End If
    End If
End Sub

Private Sub txtINAM_LostFocus()
TXTINAM.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub TXTITM_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If TXTFROMDIV = Empty Then TXTFROMDIV.Enabled = True: TXTFROMDIV.SetFocus: Exit Sub
    
    If KeyCode = vbKeyF2 Or (Trim(txtitm) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtitm.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & "'", 0, txtitm.Text, "SELECT FINISH ITEM FROM LIST")
        
        If key_PressNew = True Then
          M_DESC = ""
          frm_FinItmMst.ONLINEITEM = True
          txtitm = Empty
          frm_FinItmMst.Show
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub txtLTNO_Change()
If SAVEFLAG Then lstBox.ListItems.Clear: Call GenerateBoxList
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNo = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtLTNo = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & "'"
   txtLTNo = SearchList(SQL)
End If
   txtitm = FindItem
Me.KeyPreview = True
End Sub

Private Sub txtMachine_GotFocus()
  TXTMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
If TXTTODIV = Empty Then TXTTODIV.Enabled = True: TXTTODIV.SetFocus: Exit Sub
Me.KeyPreview = False
          
   If KeyCode = vbKeyF2 Or (Trim(TXTMACHINE) = Empty And KeyCode = vbKeyReturn) Then
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTMACHINE.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & TDVCD & "'", 0, TXTMACHINE.Text, "SELECT MACHINE FROM LIST")
      TXTMACHINE.Tag = Key
      MCOD = Key
   End If
    
  Me.KeyPreview = True

End Sub

Private Sub txtMACHINE_LostFocus()
TXTMACHINE.BackColor = vbWhite
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTRATE, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTSUBGRD_Change()
If SAVEFLAG Then lstBox.ListItems.Clear: Call GenerateBoxList
End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If TXTGRAD = Empty Then TXTGRAD.Enabled = True: TXTGRAD.SetFocus: Exit Sub

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
ElseIf KeyCode = vbKeyF2 Or (TXTSUBGRD = Empty And KeyCode = vbKeyReturn) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT RDIFF,NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & "' AND GRAD='" & GetCode("GRDMST", TXTGRAD, "GRAD", "CODE") & "'"
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
End If
Me.KeyPreview = True
End Sub

Private Sub txtRate_LostFocus()
TXTRATE.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_LostFocus()
TXTSUBGRD.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub BRMK_GotFocus()
BRMK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTGRAD_GotFocus()
 TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
  
  If TXTGRAD = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Grade Master Help", "", TXTGRAD.Left, TXTGRAD.Top + TXTGRAD.Height + 100
  Else
      ToolTip Me, "Press {F2} For Grade Master Help", "", TXTGRAD.Left, TXTGRAD.Top + TXTGRAD.Height + 100
  End If
End Sub

Private Sub TXTGRAD_LostFocus()
TXTGRAD.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub TXTITM_GotFocus()
txtitm.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
If txtitm = Empty Then
   ToolTip Me, "Press {F2} / {Enter} For Finish Item Master Help", "", txtitm.Left, txtitm.Top + txtitm.Height + 100
Else
   ToolTip Me, "Press {F2} For Finish Item Master Help", "", txtitm.Left, txtitm.Top + txtitm.Height + 100
End If
End Sub

Private Sub TXTITM_LostFocus()
txtitm.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub txtltno_GotFocus()
  txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  
  If txtLTNo = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Lot Master Help", "", txtLTNo.Left, txtLTNo.Top + txtLTNo.Height + 100
  Else
      ToolTip Me, "Press {F2} For Lot Master Help", "", txtLTNo.Left, txtLTNo.Top + txtLTNo.Height + 100
  End If
End Sub

Private Sub txtltno_LostFocus()
  txtLTNo.BackColor = vbWhite
   picToolTip.Visible = False
End Sub

Private Sub txtRate_GotFocus()
  TXTRATE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub BRMK_LostFocus()
  BRMK.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_GotFocus()
  TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
  
  If TXTSUBGRD = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For SubGrade Master Help", "", TXTSUBGRD.Left, TXTSUBGRD.Top + TXTSUBGRD.Height + 100
  Else
      ToolTip Me, "Press {F2} For SubGrade Master Help", "", TXTSUBGRD.Left, TXTSUBGRD.Top + TXTSUBGRD.Height + 100
  End If
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    TXTVBDT.Enabled = Not Yes
    TXTFROMDIV.Enabled = Not Yes
    TXTTODIV.Enabled = Not Yes
    txtLTNo.Enabled = Not Yes
    txtitm.Enabled = Not Yes
    txtIGRP.Enabled = Not Yes
    TXTINAM.Enabled = Not Yes
    TXTGRAD.Enabled = Not Yes
    TXTSUBGRD.Enabled = Not Yes
    TXTRATE.Enabled = Not Yes
    BRMK.Enabled = Not Yes
End Sub

Private Sub TimerBillNo1_Timer()
    Static ctr As Integer
    
    If ctr Mod 45 = 0 And ctr <= 45 Then
        lblAlert.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        BORDER.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        lblBill.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
    ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
        lblAlert.ForeColor = vbRed
        BORDER.BorderColor = vbRed
        lblBill.ForeColor = vbRed
    ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
        lblAlert.ForeColor = vbBlue
        BORDER.BorderColor = vbBlue
        lblBill.ForeColor = vbBlue
        ctr = 0
    End If
    
    ctr = ctr + 15
End Sub

Private Function FindItem() As String
Dim FICD As String
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset


If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & FDVCD & "' AND LTNO='" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   FICD = Trim(FINDRS!FICD & "")
Else
   FICD = Empty
   Exit Function
End If
FINDRS.Close

If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & FDVCD & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   FindItem = Trim(FINDRS!NAME & "")
Else
   FindItem = Empty
   Exit Function
End If
FINDRS.Close

End Function

Private Sub SetInternal()
Dim INDEX As Long
Dim CONVERSION As Double
TTLQTY = 0: TTLBOXES = 0: TTLCOPS = 0

SGRD = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")

Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & FDVCD & "' AND GRAD = '" & SGRD & "' AND NAME = '" & TXTSUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   SUBGRD = Trim(GRRS!SUBGRD & "")
End If
GRRS.Close

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CODE,CONVERSION FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & FDVCD & "' AND NAME = '" & txtitm & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   SITEM = Trim(GRRS!CODE & "")
   ITM_CONVERSION = Val(GRRS!CONVERSION)
   If ITM_CONVERSION < 1 Then ITM_CONVERSION = 1
End If
GRRS.Close

For INDEX = 1 To lstBox.ListItems.COUNT
  If lstBox.ListItems(INDEX).Checked = True Then
     TTLCOPS = TTLCOPS + Val(lstBox.ListItems(INDEX).SubItems(1))
     TTLBOXES = TTLBOXES + 1
     TTLQTY = TTLQTY + Val(lstBox.ListItems(INDEX).SubItems(2))
  End If
Next

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & TDVCD & "' AND NAME = '" & TXTMACHINE & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   MCOD = Trim(GRRS!CODE & "")
End If
GRRS.Close

End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "DPF"
  DLYSTA!PCOD = "CAPTIVE"
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = Val(txtNTWT)
  DLYSTA!VBNO = lblBill & ""
  DLYSTA!AMNT = Val(TXTRATE) & TTLQTY
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub TXTTODIV_GotFocus()
  TXTTODIV.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTTODIV_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
Dim SQL As String, TITLE As String

If optInternal.Value Then
   SQL = "SELECT CODE, NAME FROM DIVMST WHERE COMP='" & compPth & _
         "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'"
         
   TITLE = "SELECT DIVISION FROM LIST"
Else
   SQL = "SELECT DISTINCT CODE,NAME FROM ACCMST "
         
   TITLE = "SELECT JOB WORKER FROM LIST"
End If
   
      If KeyCode = vbKeyF2 Or (Trim(TXTTODIV) = Empty And KeyCode = vbKeyReturn) Then
         NEW_VISIBLE = False
         M_DESC = Empty
         Key = Empty
         TXTTODIV.Text = SearchList1(SQL, 0, TXTTODIV.Text, TITLE)
         TXTTODIV.Tag = Key
         TDVCD = Key
      End If
          
  Me.KeyPreview = True
End Sub

Private Sub TXTTODIV_LostFocus()
 TXTTODIV.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Function INVALIDDATA() As Boolean

If optInternal.Value = True Then
    If FDVCD = TDVCD Then
      MsgBox "Invalid Selection of Division"
      TXTFROMDIV.Enabled = True
      TXTFROMDIV.SetFocus
      INVALIDDATA = True
      Exit Function
    End If
End If

If TXTFROMDIV = Empty Then
  If TXTFROMDIV.Enabled Then TXTFROMDIV.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTTODIV = Empty Then
  If TXTTODIV.Enabled Then TXTTODIV.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If optInternal.Value = True Then
    If TXTMACHINE = Empty Then
      If TXTMACHINE.Enabled Then TXTMACHINE.SetFocus
      INVALIDDATA = True
      Exit Function
    End If
End If

If txtIGRP = Empty Then
  If txtIGRP.Enabled Then txtIGRP.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtitm = Empty Then
  If txtitm.Enabled Then txtitm.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTINAM = Empty Then
  If TXTINAM.Enabled Then TXTINAM.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtLTNo = Empty Then
  If txtLTNo.Enabled Then txtLTNo.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTGRAD = Empty Then
  If TXTGRAD.Enabled Then TXTGRAD.SetFocus
  INVALIDDATA = True
  Exit Function
End If

'If TXTSUBGRD = Empty Then
'  If TXTSUBGRD.Enabled Then TXTSUBGRD.SetFocus
'  INVALIDDATA = True
'  Exit Function
'End If

If TXTRATE = Empty Then
  If TXTRATE.Enabled Then TXTRATE.SetFocus
  INVALIDDATA = True
  Exit Function
End If
End Function

Private Sub GenerateBoxList()
If txtitm = Empty Or TXTGRAD = Empty Or txtLTNo = Empty Then Exit Sub

SGRD = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")

Dim SQL As String
Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & FDVCD & "' AND GRAD = '" & SGRD & "' AND NAME = '" & TXTSUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SUBGRD = Trim(RSDATA!SUBGRD & "")
End If
RSDATA.Close

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & FDVCD & "' AND NAME = '" & txtitm & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
   SITEM = Trim(RSDATA!CODE & "")
End If
RSDATA.Close

SQL = "SELECT BOXREGISTER.*,SUBGRDMST.NAME AS SUBGRADE FROM BOXREGISTER LEFT JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND BOXREGISTER.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD " & _
" WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & FDVCD & _
"'AND BOXREGISTER.LOTNO ='" & txtLTNo & "' AND BOXREGISTER.DBCD <> '000006' AND BOXREGISTER.ICOD = '" & SITEM & _
"' AND BOXREGISTER.GRAD ='" & SGRD & "' AND (VTYP='PPF' OR VTYP='OPN') AND BOXREGISTER.RECSTAT<>'D' AND RVBNO IS NULL "

If SUBGRD <> Empty Then
   SQL = SQL & "  AND BOXREGISTER.SUBGRD='" & SUBGRD & "' "
End If

SQL = SQL & " ORDER BY VBDT"

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   MsgBox "Boxes are not available for this criteria."
   TXTSUBGRD.Enabled = True: TXTSUBGRD.SetFocus
   Exit Sub
End If
  Dim Item
  lstBox.ListItems.Clear
  Do While Not RSDATA.EOF
   Set Item = lstBox.ListItems.ADD
   Item.Text = RSDATA!VBNO
   Item.SubItems(1) = RSDATA!COPS
   Item.SubItems(2) = nstr(RSDATA!NTWGT, 9, 3)
   Item.SubItems(2) = Trim(Item.SubItems(2))
     
   
   If Trim(RSDATA!SUBGRD) = "S" Or Trim(RSDATA!SUBGRD) = "Z" Or Trim(RSDATA!SUBGRD) = "0" Then
     Item.SubItems(3) = Trim(RSDATA!SUBGRD)
     lstBox.ColumnHeaders(4).Text = "Twist"
   Else
     Item.SubItems(3) = Trim(RSDATA!SUBGRADE & "")
     If lstBox.ListItems.COUNT = 1 Then lstBox.ColumnHeaders(4).Text = "SubGrade"
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
  
  Call FindStock
End Sub


Private Sub FindStock()
Dim INDEX As Long
Dim STOCKQTY As Double: STOCKQTY = 0
For INDEX = 1 To lstBox.ListItems.COUNT
    STOCKQTY = STOCKQTY + Val(lstBox.ListItems(INDEX).SubItems(2))
Next INDEX
TXTSTKQTY = STOCKQTY
TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
TXTSTKQTY = Trim(TXTSTKQTY)
End Sub

Private Sub SETDIVISION()
On Error GoTo SETERR

Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset

Dim INRS As ADODB.Recordset
Set INRS = New ADODB.Recordset

Dim SLIP As String

CN.BeginTrans

SLIP = GenDPFVNO("DPF", M_DBCD, FDVCD)

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA1,EXTRA2,EXTRA3)VALUES('" & compPth & "','" & UNCD & "','" & FDVCD & _
"','DPF','" & M_DBCD & "','" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','XXXXXX','" & MCOD & "','" & txtLTNo & _
"','" & SITEM & "','" & SGRD & "','" & SUBGRD & "','" & TTLBOXES & "','" & TTLQTY & _
"'," & TXTRATE & "," & (TTLQTY * Val(TXTRATE)) & ",'Q','N','" & cUName & "','*','A','" & TTLCOPS & _
"','" & BRMK & "','" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "','" & TDVCD & "')"

CN.Execute SQL

'FIN ITEM
If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & FDVCD & "' AND NAME = '" & txtitm & "'", CN, adOpenDynamic, adLockOptimistic
If Not TEMPRS.EOF Then
   
   If INRS.State = 1 Then INRS.Close
   INRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
             "' AND DVCD = '" & TDVCD & "' AND NAME = '" & txtitm & "'", CN, adOpenDynamic, adLockOptimistic
   If INRS.EOF Then
   
      CN.Execute "INSERT INTO FINITMMST(COMP,UNIT,DVCD,CODE,NAME,DENI,UOM,QORP,ISRETURNABLE) " & _
                 "VALUES('" & compPth & "','" & UNCD & "','" & TDVCD & "','" & Trim(TEMPRS!CODE & "") & _
                 "','" & Trim(TEMPRS!NAME & "") & "','" & Trim(TEMPRS!DENI & "") & "','" & Trim(TEMPRS!UOM & "") & _
                 "','" & Trim(TEMPRS!QORP & "") & "','" & Trim(TEMPRS!ISRETURNABLE & "") & "')"
   End If
   INRS.Close
End If
TEMPRS.Close

'LOT
If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT * FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & FDVCD & "' AND LTNO = '" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
If Not TEMPRS.EOF Then
   
   If INRS.State = 1 Then INRS.Close
   INRS.Open "SELECT * FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
             "' AND DVCD = '" & TDVCD & "' AND LTNO = '" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
   If INRS.EOF Then
   
      CN.Execute "INSERT INTO TXULOT(COMP,UNIT,DVCD,LTNO,SRCH,FICD,RICD,PERC,ACTIVE,SHCD,SUBPKGCODE) VALUES ('" & compPth & _
      "','" & UNCD & "','" & TDVCD & "','" & txtLTNo & "','" & Trim(TEMPRS!SRCH & "") & _
      "','" & Trim(TEMPRS!FICD & "") & "','" & Trim(TEMPRS!RICD & "") & "','" & Val(TEMPRS!PERC & "") & _
      "','" & Trim(TEMPRS!ACTIVE & "") & "','" & Trim(TEMPRS!SHCD & "") & "','" & Trim(TEMPRS!SUBPKGCODE & "") & "')"

   End If
   INRS.Close
End If
TEMPRS.Close

Dim INDEX As Long

For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
 
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
       "' AND DVCD='" & FDVCD & "' AND PKG_STCOD='" & Trim(lstBox.ListItems(INDEX).SubItems(8)) & _
       "' AND (VTYP='PPF' OR VTYP='OPN') AND " & _
       "VBNO='" & lstBox.ListItems(INDEX).Text & "'", CN, adOpenDynamic, adLockOptimistic
   
   If Not TEMPRS.EOF Then
      
      SQL = "INSERT INTO BOXREGISTER(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,PLTNO,VBDT,CHLN,PKG_STCOD,PKGNG_COD,"
      SQL = SQL & "LOCCOD,PCOD,ISRETURNABLE,LOTNO,ICOD,GRAD,SUBGRD,MCCD,COPS,BOXWGT,COPSWGT,GRSWGT,TRWGT,"
      SQL = SQL & "NTWGT,PACKER,RMRK,RECSTAT)VALUES('" & compPth & _
      "','" & UNCD & "','" & TDVCD & "','" & Trim(TEMPRS!dbcd & "") & "','PPF','" & Trim(TEMPRS!VBNO & "") & _
      "','" & Trim(TEMPRS!PLTNO & "") & "','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & Trim(TEMPRS!chln & "") & _
      "','" & Trim(TEMPRS!PKG_STCOD & "") & "','" & Trim(TEMPRS!PKGNG_COD & "") & "','" & Trim(TEMPRS!LOCCOD & "") & "','" & Trim(TEMPRS!PCOD & "") & _
      "','" & Trim(TEMPRS!ISRETURNABLE & "") & "','" & Trim(TEMPRS!LOTNO & "") & "','" & Trim(TEMPRS!ICOD & "") & _
      "','" & Trim(TEMPRS!grad & "") & "','" & Trim(TEMPRS!SUBGRD & "") & "','" & MCOD & _
      "','" & Val(TEMPRS!COPS) & "','" & Val(TEMPRS!BOXWGT) & _
      "','" & Val(TEMPRS!COPSWGT) & "','" & Val(TEMPRS!GRSWGT) & "','" & Val(TEMPRS!TRWGT) & "','" & Val(TEMPRS!NTWGT) & _
      "','" & Trim(TEMPRS!PACKER & "") & "','" & Trim(TEMPRS!RMRK & "") & "','A')"
      
      CN.Execute SQL
   
   End If
   
 End If
Next INDEX


For INDEX = 1 To lstBox.ListItems.COUNT
 If lstBox.ListItems(INDEX).Checked = True Then
   SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & SLIP & "',RVBDT= '" & Format(TXTVBDT, "YYYY/MM/DD") & _
   "',RDBC = '" & M_DBCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & _
   "' AND PKG_STCOD='" & Trim(lstBox.ListItems(INDEX).SubItems(8)) & "' AND (VTYP='PPF' OR VTYP='OPN') AND " & _
   "VBNO='" & lstBox.ListItems(INDEX).Text & "'"
   
   CN.Execute SQL
   
 End If
Next INDEX

SQL = "UPDATE SERIALMASTER SET [SRNO]='" & SLIP & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
"' AND VTYP='DPF' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "'"
   
CN.Execute SQL


'DAILYSTATUS ENTRY
 If SAVEFLAG = True Then
  Call DAILYSTATUS("DPF", SITEM, M_DBCD, Val(txtNTWT), SLIP, 0, cUName, "N", Now, TXTVBDT)
 Else
  Call DAILYSTATUS("DPF", SITEM, M_DBCD, Val(txtNTWT), SLIP, 0, cUName, "M", Now, TXTVBDT)
 End If
'---------------------------
CN.CommitTrans

If SAVEFLAG Then
  MsgBox "Your Captive Challan No. is : " & SLIP
Else
  MsgBox "Your Captive Challan No. : " & SLIP & " is Successfully Edited."
End If

Call cmdCancel_Click

Exit Sub
SETERR:
MsgBox ERR.Description
End Sub


Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub setStatus()
If optInternal.Value = True Then
   LBLTODIV.Caption = "To Division        :"
   TXTMACHINE.Enabled = True
   LBLMAC.Enabled = True
Else
   LBLTODIV.Caption = "To Party        :"
   TXTMACHINE.Enabled = False
   LBLMAC.Enabled = False
   TXTTODIV = Empty
End If

End Sub

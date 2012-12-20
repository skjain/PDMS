VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmCaptiveChallan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captive Challan Module"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   11400
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   5640
   End
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   42
      Top             =   5760
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
         TabIndex        =   43
         Top             =   0
         Width           =   120
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   5355
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9446
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
      Begin VB.TextBox TXTMACHINE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1420
         Width           =   4335
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
         Left            =   4680
         TabIndex        =   17
         Text            =   ".00"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox TXTTODIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox TXTFROMDIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox TXTINAM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox TXTIGRP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtLTNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox BRMK 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3840
         Width           =   5415
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox TXTRATE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         MaxLength       =   200
         TabIndex        =   19
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6480
         TabIndex        =   18
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox TXTGRAD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox TXTITM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox TXTAMNT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9600
         TabIndex        =   20
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox TXTPCS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   16
         Top             =   3360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   9120
         TabIndex        =   6
         Top             =   960
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
         Format          =   16580609
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   960
         TabIndex        =   0
         Top             =   4560
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmCaptiveChallan.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   5280
         TabIndex        =   3
         Top             =   4560
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmCaptiveChallan.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6720
         TabIndex        =   4
         Top             =   4560
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmCaptiveChallan.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   4560
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
         Image           =   "frmCaptiveChallan.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   4560
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
         Image           =   "frmCaptiveChallan.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   8160
         TabIndex        =   5
         Top             =   4560
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
         Image           =   "frmCaptiveChallan.frx":1CAA
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   9600
         TabIndex        =   46
         Top             =   4560
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
         Image           =   "frmCaptiveChallan.frx":20FC
         cBack           =   -2147483633
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "To Machine     :"
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
         TabIndex        =   45
         Top             =   1410
         Width           =   1575
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         X1              =   3120
         X2              =   3120
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Qnty."
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
         Left            =   4800
         TabIndex        =   44
         Tag             =   "S"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         TabIndex        =   41
         Top             =   960
         Width           =   1575
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
         TabIndex        =   40
         Top             =   480
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   735
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4440
         Width           =   11175
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   9480
         X2              =   9480
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   7800
         X2              =   7800
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   4560
         X2              =   4560
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "------------->"
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
         TabIndex        =   39
         Tag             =   "S"
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "----------->>"
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
         TabIndex        =   38
         Tag             =   "S"
         Top             =   2040
         Width           =   975
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
         Left            =   6960
         TabIndex        =   37
         Tag             =   "S"
         Top             =   2400
         Width           =   1095
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
         Left            =   6960
         TabIndex        =   36
         Tag             =   "S"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   5520
         X2              =   6000
         Y1              =   2760
         Y2              =   2760
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
         Left            =   3240
         TabIndex        =   35
         Top             =   0
         Width           =   4455
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
         TabIndex        =   34
         Top             =   960
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1695
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   11055
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
         TabIndex        =   33
         Top             =   480
         Width           =   1575
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   6720
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   4095
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
         TabIndex        =   32
         Top             =   480
         Width           =   2295
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2415
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   11175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2880
         Y2              =   2880
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
         TabIndex        =   31
         Tag             =   "S"
         Top             =   2160
         Width           =   1095
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
         Left            =   2040
         TabIndex        =   30
         Tag             =   "S"
         Top             =   3000
         Width           =   855
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
         Left            =   8160
         TabIndex        =   29
         Tag             =   "S"
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Qnty."
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
         TabIndex        =   28
         Tag             =   "S"
         Top             =   3000
         Width           =   1095
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
         Left            =   600
         TabIndex        =   27
         Tag             =   "S"
         Top             =   3000
         Width           =   615
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
         Left            =   2760
         TabIndex        =   26
         Tag             =   "S"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   6240
         X2              =   6240
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   1680
         X2              =   1680
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000080&
         X1              =   6000
         X2              =   6000
         Y1              =   2160
         Y2              =   2760
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
         Left            =   1200
         TabIndex        =   25
         Tag             =   "S"
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   9720
         TabIndex        =   24
         Tag             =   "S"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pcs"
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
         Left            =   3240
         TabIndex        =   23
         Tag             =   "S"
         Top             =   3000
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCaptiveChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FDVCD As String
Dim TDVCD As String
Dim MCOD As String
Dim SQL As String
Dim SAVEFLAG As Boolean
Dim M_DBCD As String
Public RAWITMGRP As String
Dim RAWITM As String
Dim SITEM As String
Dim MACCOD As String
Dim SGRD As String
Dim SUBGRD As String
Dim ALLOWEDITDEL As Boolean
Public CHALLAN As String

Private Sub cmdCancel_Click()
  Call ClsData(Me)
   
    Call btn_sts(True)
    If zoomflag = True Then
       Call cmdExit_Click
       Exit Sub
    End If
    
    lblBill.Caption = GenDPFVNO("DPF", M_DBCD, FDVCD)
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
       
    frmCaptiveChlnList.FDVCD = FDVCD
    frmCaptiveChlnList.TDVCD = TDVCD
    frmCaptiveChlnList.M_DBCD = M_DBCD
    
    CHALLAN = Empty
    
    frmCaptiveChlnList.Show 1
  
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
        Call DAILYSTATUS("DPF", GetCode("MACMST", txtMACHINE, "NAME", "CODE"), M_DBCD, Val(txtQty), lblBill, Val(TXTAMNT), cUName, "N", Now, TXTVBDT)
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
       If TXTFROMDIV.Enabled = True Then TXTFROMDIV.SetFocus
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
       
    frmCaptiveChlnList.FDVCD = FDVCD
    frmCaptiveChlnList.TDVCD = TDVCD
    frmCaptiveChlnList.M_DBCD = M_DBCD
    
    CHALLAN = Empty
    
    frmCaptiveChlnList.Show 1
    
    If CHALLAN = Empty Or CHALLAN = "" Then
        btn_sts (True)
        cmdAdd.Enabled = True
        cmdAdd.SetFocus
    Else
        btn_sts (False)
        TXTFROMDIV.Enabled = True
        TXTFROMDIV.SetFocus
    End If
    TXTSTKQTY = FindStock
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim Index As Long
Dim FLAG As Boolean
Dim SLIP As String
Dim COPS As Double
Dim PCS As Double

If INVALIDDATA Then Exit Sub

Call SetInternal

If Val(txtQty) > FindStock Then
   MsgBox "Challan Quantity Exceed From Stock Quantity."
   txtQty.Enabled = True: txtQty.SetFocus: Exit Sub
End If

'DUPLICACY CHECKING OF CHALLAN
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

If SAVEFLAG = True Then

SLIP = GenDPFVNO("DPF", M_DBCD, FDVCD)

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA1,EXTRA2,EXTRA3)VALUES('" & compPth & "','" & UNCD & "','" & FDVCD & _
"','DPF','" & M_DBCD & "','" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','XXXXXX','" & MACCOD & "','" & txtLTNo & _
"','" & SITEM & "','" & SGRD & "','" & SUBGRD & "','" & TXTPCS & "','" & txtQty & _
"'," & TXTRATE & "," & TXTAMNT & ",'Q','N','" & cUName & "','*','A','" & TXTPCS & _
"','" & BRMK & "','" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "','" & TDVCD & "')"

CN.Execute SQL

SQL = "INSERT INTO PKGMAN (COMP,UNIT,DVCD,DBCD,VTYP,DATE,SLIPNO,PKG_STCOD,PCOD,"
SQL = SQL & "LOTNO,FINITMCOD,GRAD,SUBGRAD,QNTY,SYSR,[USER],OPER,RECSTAT) VALUES "
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & FDVCD & "','" & M_DBCD & "','DPF'"
SQL = SQL & ",'" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & SLIP & "','000000','" & MACCOD & "',"
SQL = SQL & "'" & txtLTNo & "','" & SITEM & "','" & SGRD & "','" & SUBGRD & "','" & txtQty & _
"','N','" & cUName & "','-','A')"

CN.Execute SQL

SQL = "INSERT INTO STORETRAN(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,CHLN,CHDT,[DATE],PCOD,ICOD,PCES,"
SQL = SQL & "QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,LTNO,GRAD,SUBGRD,COPS,RECSTAT)VALUES"
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & TDVCD & "','" & M_DBCD & "','DPF'"
SQL = SQL & ",'" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & MACCOD & "','" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "','" & TXTPCS & "','" & txtQty & _
"'," & TXTRATE & "," & TXTAMNT & ",'Q','N','" & cUName & "','+','" & txtLTNo & _
"','" & SGRD & "','" & SUBGRD & "','" & TXTPCS & "','A')"

CN.Execute SQL

Dim UPSQL As String
UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & SLIP & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
         "' AND VTYP='DPF' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "' "

If UNT_DIVSERIES_REQ = "Y" Then
   UPSQL = UPSQL & " AND DVCD='" & FDVCD & "' "
End If
 
CN.Execute UPSQL

Else

SLIP = lblBill.Caption

SQL = "UPDATE SPTRAN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',PCOD='" & MACCOD & _
"',ICOD='" & SITEM & "',PCES='" & TXTPCS & "',QNTY='" & txtQty & "',RATE=" & TXTRATE & _
",AMNT=" & TXTAMNT & ",GRAD='" & SGRD & "',LTNO='" & txtLTNo & "',COPS='" & TXTPCS & _
"',EXTRA1='" & BRMK & "',EXTRA2='" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "',EXTRA3='" & TDVCD & "' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & _
"' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & SLIP & "'"
   
CN.Execute SQL

SQL = "UPDATE PKGMAN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',PCOD='" & MACCOD & _
"',LOTNO='" & txtLTNo & "',FINITMCOD='" & SITEM & "',GRAD='" & SGRD & "',SUBGRAD='" & SUBGRD & _
"',NOB='" & TXTPCS & "',QNTY='" & txtQty & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & _
"' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND  SLIPNO = '" & SLIP & "'"

CN.Execute SQL

SQL = "UPDATE STORETRAN SET DATE='" & Format(TXTVBDT, "YYYY/MM/DD") & "',PCOD='" & MACCOD & _
"',ICOD ='" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "',PCES='" & TXTPCS & "',QNTY='" & txtQty & "',RATE=" & TXTRATE & _
",AMNT=" & TXTAMNT & ",LTNO='" & txtLTNo & "',GRAD='" & SGRD & "',SUBGRD='" & SUBGRD & _
"',COPS ='" & TXTPCS & "'"

CN.Execute SQL

End If
'--------------------------
'DAILYSTATUS ENTRY
If SAVEFLAG = True Then
  Call DAILYSTATUS("DPF", MACCOD, M_DBCD, Val(txtQty), SLIP, Val(TXTAMNT), cUName, "N", Now, TXTVBDT)
   Else
  Call DAILYSTATUS("DPF", MACCOD, M_DBCD, Val(txtQty), SLIP, Val(TXTAMNT), cUName, "M", Now, TXTVBDT)
End If

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
If UCase(ActiveControl.NAME) = "TXTTODIV" And Not SAVEFLAG And TXTFROMDIV <> Empty And TXTTODIV <> Empty And TXTITM = Empty And ALLOWEDITDEL = False Then
  Call cmdEdit_Click
  Exit Sub
ElseIf UCase(ActiveControl.NAME) = "TXTTODIV" And Not SAVEFLAG And TXTFROMDIV <> Empty And TXTTODIV <> Empty And TXTITM = Empty And ALLOWEDITDEL = True Then
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
  TXTVBDT = GetMinDate
  TXTVBDT = GetMaxDate
    
    If zoomflag = True Then
        btn_sts (False)
        SAVEFLAG = False
    Else
        btn_sts (True)
    End If
End Sub

Private Sub cmdadd_Click()
    zoomflag = False
    btn_sts (False)
    'lblBill.Caption = GenDPFVNO("DPF", M_DBCD, FDVCD)
    TXTFROMDIV.SetFocus
    SAVEFLAG = True
End Sub

Private Sub cmdExit_Click()
  Unload Me
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
   TXTSTKQTY = FindStock
   TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
   TXTSTKQTY = Trim(TXTSTKQTY)
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

Private Sub TXTITM_Change()
   TXTSTKQTY = FindStock
   TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
   TXTSTKQTY = Trim(TXTSTKQTY)
End Sub

Private Sub TXTITM_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If TXTFROMDIV = Empty Then TXTFROMDIV.Enabled = True: TXTFROMDIV.SetFocus: Exit Sub
    
    If KeyCode = vbKeyF2 Or (Trim(TXTITM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTITM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & FDVCD & "'", 0, TXTITM.Text, "SELECT FINISH ITEM FROM LIST")
        
        If key_PressNew = True Then
          M_DESC = ""
          frm_FinItmMst.ONLINEITEM = True
          TXTITM = Empty
          frm_FinItmMst.Show
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub txtLTNO_Change()
   TXTSTKQTY = FindStock
   TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
   TXTSTKQTY = Trim(TXTSTKQTY)
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
   TXTITM = FindItem
Me.KeyPreview = True
End Sub

Private Sub txtMachine_GotFocus()
  txtMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
If txtMACHINE = Empty Then
   ToolTip Me, "Press {F2} / {Enter} For Machine Master Help", "", txtMACHINE.Left, txtMACHINE.Top + txtMACHINE.Height + 100
Else
   ToolTip Me, "Press {F2} For Machine Master Help", "", txtMACHINE.Left, txtMACHINE.Top + txtMACHINE.Height + 100
End If

End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
If TXTTODIV = Empty Then TXTTODIV.Enabled = True: TXTTODIV.SetFocus: Exit Sub
Me.KeyPreview = False
          
   If KeyCode = vbKeyF2 Or (Trim(txtMACHINE) = Empty And KeyCode = vbKeyReturn) Then
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      txtMACHINE.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & TXTTODIV.Tag & "'", 0, txtMACHINE.Text, "SELECT MACHINE FROM LIST")
      txtMACHINE.Tag = Key
      MCOD = Key
   End If
    
  Me.KeyPreview = True
End Sub

Private Sub txtMACHINE_LostFocus()
 txtMACHINE.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub TXTPCS_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTPCS, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtQty, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTRATE, Me) = 0 Then KeyAscii = 0
 TXTAMNT = Val(txtQty) * Val(TXTRATE)
 TXTAMNT = nstr(TXTAMNT, 12, 2)
End Sub

Private Sub TXTSUBGRD_Change()
   TXTSTKQTY = FindStock
   TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
   TXTSTKQTY = Trim(TXTSTKQTY)
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

Private Sub TXTAMNT_LostFocus()
TXTAMNT.BackColor = vbWhite
End Sub

Private Sub TXTPCS_LostFocus()
TXTPCS.BackColor = vbWhite
End Sub

Private Sub TXTQTY_LostFocus()
txtQty.BackColor = vbWhite
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

Private Sub txtAmnt_GotFocus()
   TXTAMNT.BackColor = RGB(BRED, BGREEN, BBLUE)
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
TXTITM.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
If TXTITM = Empty Then
   ToolTip Me, "Press {F2} / {Enter} For Finish Item Master Help", "", TXTITM.Left, TXTITM.Top + TXTITM.Height + 100
Else
   ToolTip Me, "Press {F2} For Finish Item Master Help", "", TXTITM.Left, TXTITM.Top + TXTITM.Height + 100
End If
End Sub

Private Sub TXTITM_LostFocus()
TXTITM.BackColor = vbWhite
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

Private Sub TXTPCS_GotFocus()
TXTPCS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTQTY_GotFocus()
 txtQty.BackColor = RGB(BRED, BGREEN, BBLUE)
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
    TXTITM.Enabled = Not Yes
    txtIGRP.Enabled = Not Yes
    TXTINAM.Enabled = Not Yes
    TXTGRAD.Enabled = Not Yes
    TXTSUBGRD.Enabled = Not Yes
    TXTPCS.Enabled = Not Yes
    txtQty.Enabled = Not Yes
    TXTRATE.Enabled = Not Yes
    TXTAMNT.Enabled = Not Yes
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
GRRS.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & FDVCD & "' AND NAME = '" & TXTITM & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   SITEM = Trim(GRRS!CODE & "")
End If
GRRS.Close

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & TDVCD & "' AND NAME = '" & txtMACHINE & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   MACCOD = Trim(GRRS!CODE & "")
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
  DLYSTA!QNTY = Val(txtQty)
  DLYSTA!VBNO = lblBill & ""
  DLYSTA!AMNT = Val(TXTAMNT)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub TXTTODIV_GotFocus()
  TXTTODIV.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
If TXTTODIV = Empty Then
   ToolTip Me, "Press {F2} / {Enter} For Division Master Help", "", TXTTODIV.Left, TXTTODIV.Top + TXTTODIV.Height + 100
Else
   ToolTip Me, "Press {F2} For Division Master Help", "", TXTTODIV.Left, TXTTODIV.Top + TXTTODIV.Height + 100
End If
End Sub

Private Sub TXTTODIV_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
          
   If KeyCode = vbKeyF2 Or (Trim(TXTTODIV) = Empty And KeyCode = vbKeyReturn) Then
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTTODIV.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, TXTTODIV.Text, "SELECT DIVISION FROM LIST")
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

If FDVCD = TDVCD Then
  MsgBox "Invalid Selection of Division"
  TXTFROMDIV.Enabled = True
  TXTFROMDIV.SetFocus
  INVALIDDATA = True
  Exit Function
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

If txtMACHINE = Empty Then
  If txtMACHINE.Enabled Then txtMACHINE.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtIGRP = Empty Then
  If txtIGRP.Enabled Then txtIGRP.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTITM = Empty Then
  If TXTITM.Enabled Then TXTITM.SetFocus
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

If TXTSUBGRD = Empty Then
  If TXTSUBGRD.Enabled Then TXTSUBGRD.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTPCS = Empty Then
  If TXTPCS.Enabled Then TXTPCS.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtQty = Empty Or Val(txtQty) = 0 Then
  If txtQty.Enabled Then txtQty.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTRATE = Empty Then
  If TXTRATE.Enabled Then TXTRATE.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTAMNT = Empty Then
  If TXTAMNT.Enabled Then TXTAMNT.SetFocus
  INVALIDDATA = True
  Exit Function
End If
End Function

Private Function FindStock() As Double

If txtLTNo = Empty Or TXTINAM = Empty Or TXTGRAD = Empty Or TXTSUBGRD = Empty Then FindStock = 0: Exit Function

Call SetInternal

Dim PACKEDQTY As Double: PACKEDQTY = 0
Dim DISPATCHEDQTY As Double: DISPATCHEDQTY = 0

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT SUM(ISNULL(QNTY,0)) AS PACKED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & FDVCD & "' AND VTYP='PPF' AND LOTNO='" & txtLTNo & _
"' AND FINITMCOD='" & SITEM & "' AND GRAD='" & SGRD & "' AND SUBGRAD='" & SUBGRD & _
"' AND DBCD = '000001' AND OPER='+' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
 PACKEDQTY = Val(Trim(CHKRS!PACKED & ""))
End If
CHKRS.Close

If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT SUM(ISNULL(QNTY,0)) AS DISPACHED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & FDVCD & "' AND VTYP='DPF' AND DBCD='000004' AND LOTNO='" & txtLTNo & _
"' AND FINITMCOD='" & SITEM & "' AND GRAD='" & SGRD & _
"' AND SUBGRAD='" & SUBGRD & "' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic

If Not CHKRS.EOF Then
 DISPATCHEDQTY = Val(Trim(CHKRS!DISPACHED & ""))
End If
CHKRS.Close

Dim STOCK As Double
STOCK = PACKEDQTY - DISPATCHEDQTY

If Trim(TXTITM) = Trim(TXTITM.Tag) And Trim(TXTGRAD) = Trim(TXTGRAD.Tag) And Trim(TXTSUBGRD) = Trim(TXTSUBGRD.Tag) And Trim(txtLTNo.Tag) = Trim(txtLTNo) And Not SAVEFLAG Then
   STOCK = STOCK + Val(txtQty.Tag)
End If
   
   FindStock = STOCK
   
End Function


VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRatFactMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate Factor Master"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10470
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   7320
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   7395
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13044
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
      Begin VB.TextBox TXTTCESS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   16
         ToolTipText     =   "Enter the Description of Item."
         Top             =   4575
         Width           =   675
      End
      Begin VB.TextBox TXTTCESSFACTOR 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   41
         ToolTipText     =   "Enter the Description of Item."
         Top             =   4935
         Width           =   1155
      End
      Begin VB.TextBox TXTAVAT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   14
         ToolTipText     =   "Enter the Description of Item."
         Top             =   3000
         Width           =   675
      End
      Begin VB.TextBox TXTVAT2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   720
         MaxLength       =   6
         TabIndex        =   13
         ToolTipText     =   "Enter the Description of Item."
         Top             =   3000
         Width           =   675
      End
      Begin VB.TextBox TXTHEDCESS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3480
         MaxLength       =   6
         TabIndex        =   12
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox TXTEDCESS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   11
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox TXTCENVAT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   720
         MaxLength       =   6
         TabIndex        =   10
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox TXTCSTFACTOR 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   33
         ToolTipText     =   "Enter the Description of Item."
         Top             =   3960
         Width           =   1155
      End
      Begin VB.TextBox TXTVATFACTOR 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   32
         ToolTipText     =   "Enter the Description of Item."
         Top             =   3000
         Width           =   1155
      End
      Begin VB.TextBox TXTEXCISEFACTOR 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   30
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2040
         Width           =   1155
      End
      Begin VB.TextBox TXTCST 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   15
         ToolTipText     =   "Enter the Description of Item."
         Top             =   3600
         Width           =   675
      End
      Begin VB.TextBox TXTVAT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   20
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2640
         Width           =   675
      End
      Begin VB.TextBox TXTEXCISE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   19
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1680
         Width           =   675
      End
      Begin VB.CheckBox ChkBrokerage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Brokerage Required"
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
         Left            =   4560
         TabIndex        =   18
         Top             =   6000
         Width           =   2600
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4560
         MaxLength       =   49
         TabIndex        =   22
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1080
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   360
         MaxLength       =   49
         TabIndex        =   9
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1080
         Width           =   5235
      End
      Begin VB.ListBox lstRef 
         Height          =   4935
         Left            =   7560
         Sorted          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox TXTCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   17
         ToolTipText     =   "Enter the Description of Item."
         Top             =   6000
         Width           =   1155
      End
      Begin ButtonPlusCtl.ButtonPlus cmdFind 
         Height          =   375
         Left            =   9000
         TabIndex        =   8
         Top             =   6000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Find"
      End
      Begin ButtonPlusCtl.ButtonPlus cmdClear 
         Height          =   375
         Left            =   7800
         TabIndex        =   7
         Top             =   6000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C&lear"
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   1200
         TabIndex        =   0
         Top             =   6600
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
         Image           =   "frmRatFactMst.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   5160
         TabIndex        =   3
         Top             =   6600
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
         Image           =   "frmRatFactMst.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6480
         TabIndex        =   4
         Top             =   6600
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
         Image           =   "frmRatFactMst.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   6600
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
         Image           =   "frmRatFactMst.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   6600
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
         Image           =   "frmRatFactMst.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7800
         TabIndex        =   5
         Top             =   6600
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
         Image           =   "frmRatFactMst.frx":1CAA
         cBack           =   -2147483633
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TCESS                  (%)"
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
         Left            =   4800
         TabIndex        =   43
         Top             =   4575
         Width           =   2160
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H0080FFFF&
         BorderColor     =   &H00000080&
         Height          =   990
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   4440
         Width           =   6855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTOR"
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
         Left            =   4800
         TabIndex        =   42
         Top             =   4935
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTOR"
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
         Left            =   4800
         TabIndex        =   40
         Top             =   3960
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTOR"
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
         Left            =   4800
         TabIndex        =   39
         Top             =   3000
         Width           =   960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   4680
         X2              =   4680
         Y1              =   1560
         Y2              =   3480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AVAT"
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
         Left            =   2040
         TabIndex        =   38
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
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
         Left            =   840
         TabIndex        =   37
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H_ED_CESS"
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
         Left            =   3240
         TabIndex        =   36
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ED_CESS"
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
         Left            =   1920
         TabIndex        =   35
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CENVAT"
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
         Left            =   600
         TabIndex        =   34
         Top             =   1680
         Width           =   945
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H0080FFFF&
         BorderColor     =   &H00000080&
         Height          =   990
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   3465
         Width           =   6855
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0080FFFF&
         BorderColor     =   &H00000080&
         Height          =   1050
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   2430
         Width           =   6855
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080FFFF&
         BorderColor     =   &H00000080&
         Height          =   885
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   6855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTOR"
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
         Left            =   4800
         TabIndex        =   31
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST                      (%)"
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
         Left            =   4800
         TabIndex        =   29
         Top             =   3600
         Width           =   2130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT                      (%)"
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
         Left            =   4800
         TabIndex        =   28
         Top             =   2640
         Width           =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXCISE                (%)"
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
         Left            =   4800
         TabIndex        =   27
         Top             =   1680
         Width           =   2160
      End
      Begin VB.Label lblBill 
         BackStyle       =   0  'Transparent
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
         Left            =   5640
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Calculation Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   240
         Width           =   3615
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0080FFFF&
         BorderColor     =   &H00000080&
         Height          =   450
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   3015
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Description"
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
         TabIndex        =   24
         Top             =   720
         Width           =   1665
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   7440
         X2              =   7440
         Y1              =   600
         Y2              =   6480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   10320
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   150
         X2              =   10200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   7095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   10215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Discount (In Rs.)  "
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
         Left            =   600
         TabIndex        =   23
         Top             =   6000
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmRatFactMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim COND As String
Dim INDEX As Long

Private Sub ChkBrokerage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   cmdSave.SetFocus
End If
End Sub

Private Sub cmdAdd_Click()
    Call ClsData(Me)
    Call btn_sts(False)
    Call SetCharges
    
    txtName.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdAdd_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
End Sub

Private Sub cmdCancel_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdCLEAR_Click()
    Call ClsData(Me)
    lstRef.ListIndex = -1
End Sub

Private Sub cmdClear_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdDelete_Click()

  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000020", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    Dim ANS As String, TEMPRS As New ADODB.Recordset
    
    If isFurtherEntryExist("RATE", txtCode) Then
         MsgBox "Further Entry Exist"
         Call ClsData(Me)
         lstRef.ListIndex = -1
         Call btn_sts(True)
         Exit Sub
    End If
    
    
    If txtCode.Text = "" Then Exit Sub
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.Title)
    If ANS = vbYes Then
       CN.Execute "Delete from RATEMST where CODE ='" & Trim(txtCode.Text) & "'"
       lstRef.RemoveItem lstRef.ListIndex
    End If
                
    Call ClsData(Me)
    lstRef.ListIndex = -1
    Call btn_sts(True)
End Sub

Private Sub cmdDelete_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdEdit_Click()

  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000020", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    cmdCancel.Cancel = True
    Call btn_sts(False)
    
    If lstRef.ListIndex = -1 Then lstRef.SetFocus Else txtName.SetFocus
    SAVEFLAG = False
End Sub

Private Sub cmdEdit_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdExit_Click()
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdExit_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub CMDFIND_Click()
    NEW_VISIBLE = False
    If Me.Tag <> Empty Then Ref_Cat = Me.Tag
    M_DESC = Empty
    Key = Empty
    txtName.Text = SearchList1("Select TOP 20 CODE, NAME FROM RATEMST", 0, "", "List Of " & Me.Caption)
    txtCode.Text = Key
    
    lstRef.Text = txtName.Text
    'If cmdEdit.Enabled = True Then
    '    cmdEdit.SetFocus
    'End If
    
    If txtName <> Empty Then
       txtName.Enabled = True
       txtName.SetFocus
    End If
End Sub

Private Sub cmdFind_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdSave_Click()
On Error GoTo errPRIMARYKEY
    Dim SQL As String
    Dim TEMPRS As New ADODB.Recordset
    Dim Ctrl As Control
    
    txtName.Text = Trim(txtName.Text)
    
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
           
    If Trim(txtName.Text) = "" Then
        MsgBox "Please Enter Sale Tax Name.", vbInformation, App.Title
        txtName = Trim(txtName)
        txtName.SetFocus
        Exit Sub
    End If
    
    If Trim(TXTCD.Text) = "" Then
        MsgBox "Please Enter Rate Factor.", vbInformation, App.Title
        TXTCD = Trim(TXTCD)
        TXTCD.SetFocus
        Exit Sub
    End If
    
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from RATEMST where Upper([name])='" & UCase(Trim(txtName.Text)) & "' ", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = False And SAVEFLAG Then
       MsgBox "This Name Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.Title
       TEMPRS.Close
       Exit Sub
    End If
    
   If SAVEFLAG = True Then
      On Error GoTo errPRIMARYKEY
        
      txtCode.Text = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From RATEMST ")
          
      SQL = "INSERT INTO RATEMST(CODE,[NAME],BROKERAGE,CD,CENVAT,EDUCESS,H_ED_CESS,VAT,AVAT,PERVAT,PERCST,PERTCESS,PEREXCISE,"
      SQL = SQL & "FACTORVAT,FACTORCST,FACTOREXCISE,RATE_FACTOR,RECSTAT)VALUES('" & txtCode & _
      "','" & txtName & "','" & ChkBrokerage.Value & "','" & TXTCD & "','" & Val(TXTCENVAT) & _
      "','" & Val(TXTEDCESS) & "','" & Val(TXTHEDCESS) & "','" & Val(TXTVAT2) & "','" & Val(TXTAVAT) & _
      "','" & Val(TXTVAT) & "','" & Val(TXTCST) & "','" & Val(TXTTCESS) & "','" & Val(TXTEXCISE) & "','" & Val(TXTVATFACTOR) & "','" & Val(TXTCSTFACTOR) & _
      "','" & Val(TXTEXCISEFACTOR) & "','" & Val(TXTVATFACTOR) * Val(TXTCSTFACTOR) * Val(TXTEXCISEFACTOR) & "','A')"
        
      CN.BeginTrans
      CN.Execute SQL
      CN.CommitTrans
      lstRef.AddItem UCase(txtName.Text)
    Else
    CN.BeginTrans
    SQL = "Update RATEMST set NAME = '" & UCase(Trim(txtName.Text)) & _
    "',CENVAT = '" & Val(TXTCENVAT) & "',EDUCESS ='" & Val(TXTEDCESS) & "',H_ED_CESS ='" & Val(TXTHEDCESS) & _
    "',VAT = '" & Val(TXTVAT2) & "',AVAT ='" & Val(TXTAVAT) & _
    "',CD = '" & Val(TXTCD) & "',BROKERAGE ='" & ChkBrokerage.Value & "',PERVAT='" & Val(TXTVAT) & _
    "',PERTCESS='" & Val(TXTTCESS) & "',PERCST='" & Val(TXTCST) & "',PEREXCISE='" & Val(TXTEXCISE) & "',FACTORVAT='" & Val(TXTVATFACTOR) & _
    "',FACTORCST='" & Val(TXTCSTFACTOR) & "',FACTOREXCISE='" & Val(TXTEXCISEFACTOR) & _
    "',RATE_FACTOR=" & Val(TXTVATFACTOR) * Val(TXTCSTFACTOR) * Val(TXTEXCISEFACTOR) & ""
    SQL = SQL & " WHERE CODE ='" & Trim(txtCode.Text) & "'"
    
    CN.Execute SQL
    CN.CommitTrans
    lstRef.Clear
    Call FillList("Select [NAME] from RATEMST where RECSTAT='A' ORDER BY [NAME]", lstRef)
     
    lstRef.ListIndex = -1
    End If
    
    Call btn_sts(True)
    sTxt = txtName.Text
 
    Call ClsData(Me)
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub

errPRIMARYKEY:
    CN.RollbackTrans
    
    If ERR.Number = -2147217873 Or -2147217900 Then
        txtName.SetFocus
        MsgBox "This Name Already Registered With Other Category!!!", vbInformation, "Already Registered"
    Else
        ErrNumber = ERR.Number
        ErrMessage = ERR.Description
        frm_ErrorHandler.Show vbModal
    End If
    ERR.Clear
End Sub

Private Sub cmdSave_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub Form_Activate()
    Call ColorComponent(Me)
    Me.BackColor = RGB(RED, GREEN, BLUE)
    If key_PressNew Then cmdAdd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ActiveControl.NAME = "lstRef" Then Exit Sub
    If UCase(ActiveControl.NAME) = "TXTNAME" And txtName = Empty Then Exit Sub
    If UCase(ActiveControl.NAME) = "CHKBROKERAGE" Then Exit Sub
    If UCase(ActiveControl.NAME) = "CMDSAVE" Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
On Error GoTo errLoad
  Call btn_sts(True)
  Call FillList("Select [NAME] from RATEMST WHERE RECSTAT='A' ORDER BY [NAME]", lstRef)
  cmdExit.Cancel = True
  Me.Show
  Exit Sub
  
errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    cmdFind.Enabled = Not bool
    cmdClear.Enabled = Not bool
    lstRef.Enabled = Not bool
    txtName.Enabled = Not bool
    
    TXTVAT.Enabled = Not bool
    TXTCST.Enabled = Not bool
    TXTEXCISE.Enabled = Not bool
    TXTVATFACTOR.Enabled = Not bool
    TXTCSTFACTOR.Enabled = Not bool
    TXTEXCISEFACTOR.Enabled = Not bool
    TXTCD.Enabled = Not bool
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Msg Empty
End Sub

Private Sub lstRef_Click()
Dim I As Long, J As Long
    SAVEFLAG = False
    Dim SAVDAT As New ADODB.Recordset
    Set SAVDAT = New ADODB.Recordset
    
    If lstRef.ListIndex = -1 Then Exit Sub
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "Select * from RATEMST where [NAME] = '" & (lstRef.List(lstRef.ListIndex)) & "'", CN, adOpenDynamic, adLockOptimistic
    
    With SAVDAT
        txtCode.Text = !CODE & ""
        txtName.Text = ![NAME] & ""
        TXTCD.Text = Trim(![CD] & "")
        ChkBrokerage.Value = Trim(![BROKERAGE] & "")
        
        TXTVAT.Text = !PERVAT & ""
        TXTCST.Text = !PERCST & ""
        TXTTCESS.Text = !PERTCESS & ""
        TXTEXCISE = !PEREXCISE & ""
        TXTVATFACTOR = !FACTORVAT & ""
        TXTCSTFACTOR = !FACTORCST & ""
        TXTEXCISEFACTOR = !FACTOREXCISE & ""
        
        TXTCENVAT.Text = !CENVAT & ""
        TXTEDCESS.Text = !EDUCESS & ""
        TXTHEDCESS.Text = !H_ED_CESS & ""
        
        TXTVAT2.Text = !VAT & ""
        TXTAVAT.Text = !AVAT & ""
    End With
    SAVDAT.Close
End Sub

Private Sub lstRef_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtName.Enabled = True
 txtName.SetFocus
End If
End Sub

Private Sub lstRef_GotFocus()
    lstRef.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Address"
End Sub

Private Sub lstRef_LostFocus()
lstRef.BackColor = vbWhite
End Sub

Private Sub TXTAVAT_Change()
  TXTVAT = Val(TXTVAT2) + Val(TXTAVAT)
End Sub

Private Sub TXTAVAT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTAVAT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtcd_GotFocus()
TXTCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCD_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTCD, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtcd_LostFocus()
TXTCD.BackColor = vbWhite
End Sub

Private Sub TXTCENVAT_Change()
  Call CalculateExcisePercentage
End Sub

Private Sub TXTCENVAT_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTCENVAT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTCST_Change()
Call SetFactor
End Sub

Private Sub TXTCST_GotFocus()
 TXTCST.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCST_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTCST, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTCST_LostFocus()
TXTCST.BackColor = vbWhite
End Sub

Private Sub TXTTCESS_Change()
 Call SetFactor
End Sub

Private Sub TXTTCESS_GotFocus()
 TXTTCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTTCESS_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTTCESS, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTTCESS_LostFocus()
 TXTTCESS.BackColor = vbWhite
End Sub

Private Sub TXTEDCESS_Change()
Call CalculateExcisePercentage
End Sub

Private Sub TXTEDCESS_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTEDCESS, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTEXCISE_Change()
Call SetFactor
End Sub

Private Sub TXTEXCISE_GotFocus()
 TXTEXCISE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTEXCISE_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTEXCISE, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTEXCISE_LostFocus()
TXTEXCISE.BackColor = vbWhite
End Sub

Private Sub TXTHEDCESS_Change()
Call CalculateExcisePercentage
End Sub

Private Sub TXTHEDCESS_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTHEDCESS, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Packing Station Name"
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
End Sub

Private Sub TXTNAME_LostFocus()
txtName.BackColor = vbWhite
End Sub

Public Sub FillList(SQL As String, lst As ListBox)
    Dim TEMPRS As New ADODB.Recordset
    TEMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = True Then Exit Sub
    TEMPRS.MoveFirst
    Do While Not TEMPRS.EOF
        lst.AddItem Trim(TEMPRS.Fields(0).Value)
        TEMPRS.MoveNext
    Loop
    TEMPRS.Close
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

Private Sub TXTVAT_Change()
Call SetFactor
End Sub

Private Sub TXTVAT_GotFocus()
TXTVAT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTVAT_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTVAT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTVAT_LostFocus()
TXTVAT.BackColor = vbWhite
End Sub

Private Sub SetFactor()

   If Val(TXTEXCISE) > 0 Then
      TXTEXCISEFACTOR = 1 / (1 + (Val(TXTEXCISE) / 100))
   Else
      TXTEXCISEFACTOR = "1"
   End If

   If Val(TXTVAT) > 0 Then
      TXTVATFACTOR = 1 / (1 + (Val(TXTVAT) / 100))
   Else
      TXTVATFACTOR = "1"
   End If

   If Val(TXTCST) > 0 Then
      TXTCSTFACTOR = 1 / (1 + (Val(TXTCST) / 100))
   Else
      TXTCSTFACTOR = "1"
   End If
   
   If Val(TXTTCESS) > 0 Then
      TXTTCESSFACTOR = 1 / (1 + (Val(TXTTCESS) / 100))
   Else
      TXTTCESSFACTOR = "1"
   End If

End Sub

Private Sub TXTCENVAT_GotFocus()
  TXTCENVAT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCENVAT_LostFocus()
  TXTCENVAT.BackColor = vbWhite
End Sub

Private Sub TXTEDCESS_GotFocus()
  TXTEDCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTEDCESS_LostFocus()
  TXTEDCESS.BackColor = vbWhite
End Sub

Private Sub TXTHEDCESS_GotFocus()
  TXTHEDCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTHEDCESS_LostFocus()
  TXTHEDCESS.BackColor = vbWhite
End Sub

Private Sub TXTVAT2_Change()
  TXTVAT = Val(TXTVAT2) + Val(TXTAVAT)
End Sub

Private Sub TXTVAT2_GotFocus()
  TXTVAT2.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTVAT2_LostFocus()
  TXTVAT2.BackColor = vbWhite
End Sub

Private Sub TXTVAT2_KeyPress(KeyAscii As Integer)
   If CheckNumericKey(KeyAscii, TXTVAT2, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTAVAT_GotFocus()
  TXTAVAT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTAVAT_LostFocus()
  TXTAVAT.BackColor = vbWhite
End Sub

Private Sub CalculateExcisePercentage()
Dim ED_CESS As Double, H_ED_CESS As Double
   ED_CESS = (Val(TXTEDCESS) * Val(TXTCENVAT)) / 100
   H_ED_CESS = (Val(TXTHEDCESS) * Val(TXTCENVAT)) / 100
   TXTEXCISE = Val(TXTCENVAT) + ED_CESS + H_ED_CESS
End Sub

Private Sub SetCharges()

TXTCENVAT = GetCode("CHRGMST", "CENVAT", "NAME", "PERC")
TXTVAT = GetCode("CHRGMST", "VAT", "NAME", "PERC")
TXTAVAT = GetCode("CHRGMST", "AVAT", "NAME", "PERC")
TXTEDCESS = GetCode("CHRGMST", "EDUCESS", "NAME", "PERC")
TXTHEDCESS = GetCode("CHRGMST", "H_ED_CESS", "NAME", "PERC")

End Sub

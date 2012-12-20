VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmexciseopening 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excise Opening Balances"
   ClientHeight    =   10170
   ClientLeft      =   1155
   ClientTop       =   540
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   9090
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   10215
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   18018
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
      Begin VB.TextBox RG23APCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   5490
      End
      Begin VB.TextBox RG23APCESSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   3
         Top             =   1560
         Width           =   1290
      End
      Begin VB.TextBox PLAPCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   6240
         Width           =   5490
      End
      Begin VB.TextBox PLACESSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   18
         Top             =   6240
         Width           =   1290
      End
      Begin VB.TextBox rg23cdeferedac 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4680
         Width           =   6930
      End
      Begin VB.TextBox SRVHEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   28
         Top             =   8760
         Width           =   1290
      End
      Begin VB.TextBox SRVEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   26
         Top             =   8400
         Width           =   1290
      End
      Begin VB.TextBox SRVCENAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   24
         Top             =   8040
         Width           =   1290
      End
      Begin VB.TextBox SERVICECENVAT 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   8040
         Width           =   5490
      End
      Begin VB.TextBox SERVICEEDCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   8400
         Width           =   5490
      End
      Begin VB.TextBox SERVICEHEDCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   8760
         Width           =   5490
      End
      Begin VB.TextBox PLAHEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   22
         Top             =   6960
         Width           =   1290
      End
      Begin VB.TextBox PLAEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   20
         Top             =   6600
         Width           =   1290
      End
      Begin VB.TextBox PLACENAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   16
         Top             =   5880
         Width           =   1290
      End
      Begin VB.TextBox PLACENVAT 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5880
         Width           =   5490
      End
      Begin VB.TextBox PLAEDCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   6600
         Width           =   5490
      End
      Begin VB.TextBox PLAHEDCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   6960
         Width           =   5490
      End
      Begin VB.TextBox RG23CHEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   13
         Top             =   4200
         Width           =   1290
      End
      Begin VB.TextBox RG23CEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   11
         Top             =   3840
         Width           =   1290
      End
      Begin VB.TextBox RG23CCENAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   9
         Top             =   3480
         Width           =   1290
      End
      Begin VB.TextBox RG23CCENVAT 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3480
         Width           =   5490
      End
      Begin VB.TextBox RG23CEDCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3840
         Width           =   5490
      End
      Begin VB.TextBox RG23CHEDCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4200
         Width           =   5490
      End
      Begin VB.TextBox RG23AHEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   7
         Top             =   2280
         Width           =   1290
      End
      Begin VB.TextBox RG23AEDCSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   5
         Top             =   1920
         Width           =   1290
      End
      Begin VB.TextBox RG23ACENAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   1
         Top             =   1200
         Width           =   1290
      End
      Begin VB.TextBox RG23ACENVAT 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1200
         Width           =   5490
      End
      Begin VB.TextBox RG23AEDUCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1920
         Width           =   5490
      End
      Begin VB.TextBox RG23AHEDCESS 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   5490
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6480
         TabIndex        =   29
         Top             =   9480
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
         Image           =   "frmexciseopening.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7800
         TabIndex        =   30
         Top             =   9480
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
         Image           =   "frmexciseopening.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PAPER CESS"
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
         Index           =   18
         Left            =   240
         TabIndex        =   51
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PAPER CESS"
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
         Index           =   17
         Left            =   240
         TabIndex        =   50
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "RG23-C Deffered A/c"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   16
         Left            =   240
         TabIndex        =   49
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   48
         Top             =   8040
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ED. CESS"
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
         Index           =   14
         Left            =   240
         TabIndex        =   47
         Top             =   8400
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "HR. ED. CESS"
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
         Index           =   13
         Left            =   240
         TabIndex        =   46
         Top             =   8760
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   1455
         Left            =   120
         Top             =   7800
         Width           =   8775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICE TAX"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   45
         Top             =   7560
         Width           =   1815
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   44
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ED. CESS"
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
         Index           =   10
         Left            =   240
         TabIndex        =   43
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "HR. ED. CESS"
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
         Index           =   9
         Left            =   240
         TabIndex        =   42
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1815
         Left            =   120
         Top             =   5640
         Width           =   8775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PLA"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   40
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ED. CESS"
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
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "HR. ED. CESS"
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
         Index           =   5
         Left            =   240
         TabIndex        =   38
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   2055
         Left            =   120
         Top             =   3240
         Width           =   8775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "RG23-C-II"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
         Caption         =   "Excise Opening Balance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5760
         TabIndex        =   36
         Top             =   120
         Width           =   3015
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   420
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ED. CESS"
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
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "HR. ED. CESS"
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
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1815
         Left            =   120
         Top             =   960
         Width           =   8775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "RG23-A-II"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmexciseopening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  'Check Valid Code
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & RG23ACENVAT & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    RG23ACENAMT.SetFocus
    Exit Sub
  End If
  RG23ACENVAT.Tag = RS!code
  
  If UNT_ISPAPPER = "Y" Then
    
    If RS.State = 1 Then RS.Close
    RS.Open "select code from accmst where name='" & RG23APCESS & "'", CN, adOpenDynamic, adLockOptimistic
    If RS.EOF Then
      MsgBox "Invalid A/c"
      RG23APCESS.SetFocus
      Exit Sub
    End If
    RG23APCESS.Tag = RS!code
   Else
    RG23APCESS.Tag = "XXXXXX"
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & RG23AEDUCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    RG23AEDUCESS.SetFocus
    Exit Sub
  End If
  RG23AEDUCESS.Tag = RS!code
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & RG23AHEDCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    RG23AHEDCESS.SetFocus
    Exit Sub
  End If
  RG23AHEDCESS.Tag = RS!code
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & RG23CCENVAT & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    RG23CCENAMT.SetFocus
    Exit Sub
  End If
  RG23CCENVAT.Tag = RS!code
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & RG23CEDCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    RG23CEDCESS.SetFocus
    Exit Sub
  End If
  RG23CEDCESS.Tag = RS!code
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & RG23CHEDCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    RG23CHEDCESS.SetFocus
    Exit Sub
  End If
  RG23CHEDCESS.Tag = RS!code
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & rg23cdeferedac & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    rg23cdeferedac.SetFocus
    Exit Sub
  End If
  rg23cdeferedac.Tag = RS!code
  
  'PLA
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & PLACENVAT & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    PLACENAMT.SetFocus
    Exit Sub
  End If
  PLACENVAT.Tag = RS!code
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & PLAPCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    PLAPCESS.SetFocus
    Exit Sub
  End If
  PLAPCESS.Tag = RS!code
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & PLAEDCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    PLAEDCESS.SetFocus
    Exit Sub
  End If
  PLAEDCESS.Tag = RS!code
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & PLAHEDCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    PLAHEDCESS.SetFocus
    Exit Sub
  End If
  PLAHEDCESS.Tag = RS!code
  'Service Tax
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & SERVICECENVAT & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    SERVICECENVAT.SetFocus
    Exit Sub
  End If
  SERVICECENVAT.Tag = RS!code
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & SERVICEEDCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    SERVICEEDCESS.SetFocus
    Exit Sub
  End If
  SERVICEEDCESS.Tag = RS!code
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & SERVICEHEDCESS & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid A/c"
    SERVICEHEDCESS.SetFocus
    Exit Sub
  End If
  On Error GoTo LAST
  SERVICEHEDCESS.Tag = RS!code
  CN.BeginTrans
  CN.Execute "DELETE FROM EXCISEOPENING where comp='" & compPth & "' and unit='" & UNCD & "'"
  
  CN.Execute "delete from trnman where COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND EXTRA1='EXC' AND VTYP='OPN'"
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM EXCISEOPENING where comp='" & compPth & "' and unit='" & UNCD & "'"
  RS.AddNew
  RS!COMP = compPth
  RS!UNIT = UNCD
  RS!RG23ACENVAT = RG23ACENVAT.Tag
  RS!RG23ACENAMT = Val(RG23ACENAMT)
  
  RS!rg23acessac = RG23APCESS.Tag
  RS!rg23acessamt = Val(RG23APCESSAMT)
  
  RS!RG23AEDUCESS = RG23AEDUCESS.Tag
  RS!RG23AEDUCSAMT = Val(RG23AEDCSAMT)
  RS!RG23AHEDCESS = RG23AHEDCESS.Tag
  RS!RG23AHEDCESSAMT = Val(RG23AHEDCSAMT)
  
  RS!RG23CCENVAT = RG23CCENVAT.Tag
  RS!RG23CCENAMT = Val(RG23CCENAMT)
  RS!RG23CEDUCESS = RG23CEDCESS.Tag
  RS!RG23CEDUCSAMT = Val(RG23CEDCSAMT)
  RS!RG23CHEDCESS = RG23CHEDCESS.Tag
  RS!RG23CHEDCESSAMT = Val(RG23CHEDCSAMT)
  RS!rg23cdeffered = rg23cdeferedac.Tag
  
  RS!PLACENVAT = PLACENVAT.Tag
  RS!PLACENAMT = Val(PLACENAMT)
  
  RS!placessac = PLAPCESS.Tag
  RS!PLACESSAMT = Val(PLACESSAMT)
  
  RS!rg23acessamt = Val(RG23APCESSAMT)
  RS!PLAEDUCESS = PLAEDCESS.Tag
  RS!PLAEDUCSAMT = Val(PLAEDCSAMT)
  RS!PLAHEDCESS = PLAHEDCESS.Tag
  RS!PLAHEDCESSAMT = Val(PLAHEDCSAMT)
  
  
  RS!SRVCENVAT = SERVICECENVAT.Tag
  RS!SRVCENAMT = Val(SRVCENAMT)
  RS!SRVAEDUCESS = SERVICEEDCESS.Tag
  RS!SRVEDUCSAMT = Val(SRVEDCSAMT)
  RS!SRVHEDCESS = SERVICEHEDCESS.Tag
  RS!SRVHEDCESSAMT = Val(SRVHEDCSAMT)
  RS.Update
  
  
  'Effect In EGpman
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='EXC' AND (DBCD='RG23-A' OR DBCD='RG23-C' OR DBCD='PLAREG' OR DBCD='SRVREG') AND EXTRA5='Opening'"
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  
  If Val(RG23ACENAMT) <> 0 Or Val(RG23AEDCSAMT) <> 0 Or Val(RG23AHEDCSAMT) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXC"
    SAVDAT!SRNO = CStr(FSDT - 1)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!dbcd = "RG23-A"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = "OPNRG23-A"
    SAVDAT!chln = "OPNRG23-A"
    SAVDAT!CHDT = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!TTYP = "RG23-A"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(RG23ACENAMT)
    SAVDAT!CESS = Val(RG23APCESS)
    SAVDAT!EDUCESS = Val(RG23AEDCSAMT)
    SAVDAT!H_ED_CESS = Val(RG23AHEDCSAMT)
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "Opening"
    SAVDAT.Update
  End If
  
  
  If Val(RG23CCENAMT) <> 0 Or Val(RG23CEDCSAMT) <> 0 Or Val(RG23CHEDCSAMT) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXC"
    SAVDAT!SRNO = CStr(FSDT - 1)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!dbcd = "RG23-C"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = "OPNRG23-C"
    SAVDAT!chln = "OPNRG23-C"
    SAVDAT!CHDT = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!TTYP = "RG23-C"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(RG23CCENAMT)
    SAVDAT!EDUCESS = Val(RG23CEDCSAMT)
    SAVDAT!H_ED_CESS = Val(RG23CHEDCSAMT)
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "Opening"
    SAVDAT.Update
  End If
  
  
  If Val(PLACENAMT) <> 0 Or Val(PLAEDCSAMT) <> 0 Or Val(PLAHEDCSAMT) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXC"
    SAVDAT!SRNO = CStr(FSDT - 1)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!dbcd = "PLAREG"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = "OPNPLAREG"
    SAVDAT!chln = "OPNPLAREG"
    SAVDAT!CHDT = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!TTYP = "PLAREG"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(PLACENAMT)
    SAVDAT!CESS = Val(PLACESSAMT)
    SAVDAT!EDUCESS = Val(PLAEDCSAMT)
    SAVDAT!H_ED_CESS = Val(PLAHEDCSAMT)
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "Opening"
    SAVDAT.Update
  End If
  
  If Val(SRVCENAMT) <> 0 Or Val(SRVEDCSAMT) <> 0 Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!UNIT = UNCD
    SAVDAT!VTYP = "EXC"
    SAVDAT!SRNO = CStr(FSDT - 1)
    SAVDAT!SRCH = 1
    SAVDAT!Date = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!dbcd = "SRVREG"
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = ""
    SAVDAT!VBNO = "OPNSRVREG"
    SAVDAT!chln = "OPNSRVREG"
    SAVDAT!CHDT = Format(FSDT - 1, "YYYY/MM/DD")
    SAVDAT!TTYP = "SERVICE TAX"
    SAVDAT!RECSTAT = "A"
    SAVDAT!CENVAT = Val(SRVCENAMT)
    SAVDAT!EDUCESS = Val(SRVEDCSAMT)
    SAVDAT!H_ED_CESS = 0
    SAVDAT!EXTRA3 = "True"
    SAVDAT!EXTRA5 = "Opening"
    SAVDAT.Update
  End If
  
  
  'Effect In Trnman
  'rg23-a (Cenvat)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & RG23ACENVAT.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = RG23ACENVAT.Tag
  RS!ACOD = RG23ACENVAT.Tag
  RS!RCOD = RG23ACENVAT.Tag
  RS!damt = RS!damt + Val(RG23ACENAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  
  'RG23-a (Paper Cess)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & RG23APCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = RG23APCESS.Tag
  RS!ACOD = RG23APCESS.Tag
  RS!RCOD = RG23APCESS.Tag
  RS!damt = RS!damt + Val(RG23APCESSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  
  'rg23-a (EDUCESS)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & RG23AEDUCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = RG23AEDUCESS.Tag
  RS!ACOD = RG23AEDUCESS.Tag
  RS!RCOD = RG23AEDUCESS.Tag
  RS!damt = RS!damt + Val(RG23AEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  'RG23a-II (Hr. Edu Cess)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & RG23AHEDCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = RG23AHEDCESS.Tag
  RS!ACOD = RG23AHEDCESS.Tag
  RS!RCOD = RG23AHEDCESS.Tag
  RS!damt = RS!damt + Val(RG23AHEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  
  'rg23-C (Cenvat)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & RG23CCENVAT.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = RG23CCENVAT.Tag
  RS!ACOD = RG23CCENVAT.Tag
  RS!RCOD = RG23CCENVAT.Tag
  RS!damt = RS!damt + Val(RG23CCENAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  'rg23-a (EDUCESS)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & RG23CEDCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = RG23CEDCESS.Tag
  RS!ACOD = RG23CEDCESS.Tag
  RS!RCOD = RG23CEDCESS.Tag
  RS!damt = RS!damt + Val(RG23CEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  'RG23a-II (Hr. Edu Cess)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & RG23CHEDCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = RG23CHEDCESS.Tag
  RS!ACOD = RG23CHEDCESS.Tag
  RS!RCOD = RG23CHEDCESS.Tag
  RS!damt = RS!damt + Val(RG23CHEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  '---------------
  'PLA (Cenvat)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & PLACENVAT.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = PLACENVAT.Tag
  RS!ACOD = PLACENVAT.Tag
  RS!RCOD = PLACENVAT.Tag
  RS!damt = RS!damt + Val(PLACENAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  
  'PLA (Paper Cess)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & PLAPCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = PLAPCESS.Tag
  RS!ACOD = PLAPCESS.Tag
  RS!RCOD = PLAPCESS.Tag
  RS!damt = RS!damt + Val(PLACESSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  
  
  
  'PLA (EDUCESS)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & PLAEDCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = PLAEDCESS.Tag
  RS!ACOD = PLAEDCESS.Tag
  RS!RCOD = PLAEDCESS.Tag
  RS!damt = RS!damt + Val(PLAEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  'PLA (Hr. Edu Cess)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & PLAHEDCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = PLAHEDCESS.Tag
  RS!ACOD = PLAHEDCESS.Tag
  RS!RCOD = PLAHEDCESS.Tag
  RS!damt = RS!damt + Val(PLAHEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  '--------------------------------
  'SERVICE TAX
  'SERVICE (Cenvat)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & SERVICECENVAT.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = SERVICECENVAT.Tag
  RS!ACOD = SERVICECENVAT.Tag
  RS!RCOD = SERVICECENVAT.Tag
  RS!damt = RS!damt + Val(SRVCENAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  'SERVICE (EDUCESS)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & SERVICEEDCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = SERVICEEDCESS.Tag
  RS!ACOD = SERVICEEDCESS.Tag
  RS!RCOD = SERVICEEDCESS.Tag
  RS!damt = RS!damt + Val(SRVEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  'SERVICE (Hr. Edu Cess)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * from trnman where comp='" & compPth & "' and unit='" & UNCD & "' and acod='" & SERVICEHEDCESS.Tag & "' AND VTYP='OPN'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "OPN"
  RS!SRNO = "0"
  RS!SRCH = 1
  RS!UNIT = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(FSDT - 1, "YYYY/MM/DD")
  RS!dbcd = "OPNDAT"
  RS!vno = SERVICEHEDCESS.Tag
  RS!ACOD = SERVICEHEDCESS.Tag
  RS!RCOD = SERVICEHEDCESS.Tag
  RS!damt = RS!damt + Val(SRVHEDCSAMT)
  RS!camt = 0
  RS!VBNO = "XXXXXXXXXX"
  RS!RECSTAT = "A"
  RS!MLTENT = "N"
  RS!AUST = "A"
  RS!EXTRA1 = "EXC"
  RS.Update
  
  
  
  CN.CommitTrans
  MsgBox "Save Successful"
  Unload Me
  Exit Sub
LAST:
  MsgBox ERR.Description
'  Resume
End Sub

Private Sub Form_Activate()
  If Allow_view_only = "Y" Then
    Unload Me
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call FILDTL
  If UNT_ISPAPPER = "Y" Then
    Label8(18).Visible = True
    RG23APCESS.Visible = True
    RG23APCESSAMT.Visible = True
   Else
    Label8(18).Visible = False
    RG23APCESS.Visible = False
    RG23APCESSAMT.Visible = False
  End If
End Sub

Private Sub PLACENAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub PLACENAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, PLACENAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
   
End Sub

Private Sub PLAEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub PLAEDCSAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, PLAEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
   
End Sub

Private Sub PLAHEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub PLAHEDCSAMT_KeyPress(KeyAscii As Integer)

   If KeyAscii = 46 And InStr(1, PLAHEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub RG23ACENAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub RG23ACENAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, RG23ACENVAT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub RG23ACENVAT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(RG23ACENVAT.Text)) = Empty Then
    RG23ACENVAT = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, RG23ACENVAT, "SELECT RG23-A-II A/C (Cenvat)")
    RG23ACENVAT.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If RG23ACENVAT <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub RG23AEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub RG23AEDCSAMT_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 46 And InStr(1, RG23AEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub RG23AEDUCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(RG23AEDUCESS.Text)) = Empty Then
    RG23AEDUCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, RG23AEDUCESS, "SELECT RG23-A-II A/C (Edu. Cess)")
    RG23AEDUCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If RG23AEDUCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub RG23AHEDCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(RG23AHEDCESS.Text)) = Empty Then
    RG23AHEDCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, RG23AHEDCESS, "SELECT RG23-A-II A/C (Hr. Edu. Cess)")
    RG23AHEDCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If RG23AHEDCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub RG23AHEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub RG23AHEDCSAMT_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 46 And InStr(1, RG23AHEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub RG23CCENAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub RG23CCENAMT_KeyPress(KeyAscii As Integer)

   If KeyAscii = 46 And InStr(1, RG23CCENVAT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub RG23CCENVAT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(RG23CCENVAT.Text)) = Empty Then
    RG23CCENVAT = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, RG23CCENVAT, "SELECT RG23-C-II A/C (Cenvat)")
    RG23CCENVAT.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If RG23CCENVAT <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub rg23cdeferedac_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(rg23cdeferedac.Text)) = Empty Then
    rg23cdeferedac = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, rg23cdeferedac, "SELECT RG23-C-II A/C (Deffered A/c)")
    rg23cdeferedac.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If rg23cdeferedac <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub RG23CEDCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(RG23CEDCESS.Text)) = Empty Then
    RG23CEDCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, RG23CEDCESS, "SELECT RG23-C-II A/C (Edu. Cess)")
    RG23CEDCESS.Tag = Key
  End If
    If KeyCode = vbKeyReturn Then
    If RG23CEDCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub RG23CEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub RG23CEDCSAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, RG23CEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub RG23CHEDCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(RG23CHEDCESS.Text)) = Empty Then
    RG23CHEDCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, RG23CHEDCESS, "SELECT RG23-A-II A/C (Hr. Edu. Cess)")
    RG23CHEDCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If RG23CHEDCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub PLACENVAT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(PLACENVAT.Text)) = Empty Then
    PLACENVAT = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, PLACENVAT, "SELECT PLA A/C (Cenvat)")
    PLACENVAT.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If PLACENVAT <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub PLAEDCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(PLAEDCESS.Text)) = Empty Then
    PLAEDCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, PLAEDCESS, "SELECT PLA A/C (Edu. Cess)")
    PLAEDCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If PLAEDCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub PLAHEDCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(PLAHEDCESS.Text)) = Empty Then
    PLAHEDCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, PLAHEDCESS, "SELECT RG23-A-II A/C (Hr. Edu. Cess)")
    PLAHEDCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If PLAHEDCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub RG23CHEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub RG23CHEDCSAMT_KeyPress(KeyAscii As Integer)

   If KeyAscii = 46 And InStr(1, RG23CHEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub SERVICECENVAT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(SERVICECENVAT.Text)) = Empty Then
    SERVICECENVAT = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, SERVICECENVAT, "SELECT Serivce Tax (Cenvat)")
    SERVICECENVAT.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If SERVICECENVAT <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub SERVICEEDCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(SERVICEEDCESS.Text)) = Empty Then
    SERVICEEDCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, SERVICEEDCESS, "SELECT Service Tax (Edu. CEss)")
    SERVICEEDCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If SERVICEEDCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub SERVICEHEDCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(SERVICEHEDCESS.Text)) = Empty Then
    SERVICEHEDCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, SERVICEHEDCESS, "SELECT Serivce Tax (Hr. Edu Cess)")
    SERVICEHEDCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If SERVICEHEDCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub SRVCENAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub SRVCENAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, SRVCENAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub SRVEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub SRVEDCSAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, SRVEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub SRVHEDCSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub SRVHEDCSAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, SRVHEDCSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub
Private Sub FILDTL()
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM EXCISEOPENING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'"
  If Not RS.EOF Then
    RG23ACENVAT.Tag = RS!RG23ACENVAT
    RG23ACENAMT = RS!RG23ACENAMT
    
    RG23APCESS.Tag = RS!rg23acessac & ""
    RG23APCESSAMT = RS!rg23acessamt
    
    RG23AEDUCESS.Tag = RS!RG23AEDUCESS
    RG23AEDCSAMT = RS!RG23AEDUCSAMT
    RG23AHEDCESS.Tag = RS!RG23AHEDCESS
    RG23AHEDCSAMT = RS!RG23AHEDCESSAMT
    
    RG23CCENVAT.Tag = RS!RG23CCENVAT
    RG23CCENAMT = RS!RG23CCENAMT
    RG23CEDCESS.Tag = RS!RG23CEDUCESS
    RG23CEDCSAMT = RS!RG23CEDUCSAMT
    RG23CHEDCESS.Tag = RS!RG23CHEDCESS
    RG23CHEDCSAMT = RS!RG23CHEDCESSAMT
    rg23cdeferedac.Tag = RS!rg23cdeffered
    
    PLACENVAT.Tag = RS!PLACENVAT
    
    PLACENAMT = RS!PLACENAMT
    PLAPCESS.Tag = RS!placessac & ""
    PLACESSAMT = RS!PLACESSAMT
    PLAEDCESS.Tag = RS!PLAEDUCESS
    PLAEDCSAMT = RS!PLAEDUCSAMT
    PLAHEDCESS.Tag = RS!PLAHEDCESS
    PLAHEDCSAMT = RS!PLAHEDCESSAMT
    
    
    SERVICECENVAT.Tag = RS!SRVCENVAT
    SRVCENAMT = RS!SRVCENAMT
    SERVICEEDCESS.Tag = RS!SRVAEDUCESS
    SRVEDCSAMT = RS!SRVEDUCSAMT
    SERVICEHEDCESS.Tag = RS!SRVHEDCESS
    SRVHEDCSAMT = RS!SRVHEDCESSAMT
    'RG23-A
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & RG23ACENVAT.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      RG23ACENVAT.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & RG23APCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      RG23APCESS.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & RG23AEDUCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      RG23AEDUCESS.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & RG23AHEDCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      RG23AHEDCESS.Text = RS!NAME & ""
    End If
    'RG23-c
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & RG23CCENVAT.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      RG23CCENVAT.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & RG23CEDCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      RG23CEDCESS.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & RG23CHEDCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      RG23CHEDCESS.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & rg23cdeferedac.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      rg23cdeferedac.Text = RS!NAME & ""
    End If
    
    'pla
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & PLACENVAT.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      PLACENVAT.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & PLAPCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      PLAPCESS.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & PLAEDCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      PLAEDCESS.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & PLAHEDCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      PLAHEDCESS.Text = RS!NAME & ""
    End If
    
    'Service Tax
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & SERVICECENVAT.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      SERVICECENVAT.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & SERVICEEDCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      SERVICEEDCESS.Text = RS!NAME & ""
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & SERVICEHEDCESS.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      SERVICEHEDCESS.Text = RS!NAME & ""
    End If
    
  End If
End Sub


Private Sub RG23APCESS_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(RG23APCESS.Text)) = Empty Then
    RG23APCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, RG23APCESS, "SELECT RG23-A-II A/C (Paper Cess)")
    RG23APCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If RG23APCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub PLApcess_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or (Trim(PLAPCESS.Text)) = Empty Then
    PLAPCESS = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, PLAPCESS, "SELECT PLA A/C (Paper Cess)")
    PLAPCESS.Tag = Key
  End If
  If KeyCode = vbKeyReturn Then
    If PLAPCESS <> Empty Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub RG23APCESSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub RG23APCESSAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, RG23APCESSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0

End Sub

Private Sub PLACESSAMT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub PLACESSAMT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 46 And InStr(1, PLACESSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub


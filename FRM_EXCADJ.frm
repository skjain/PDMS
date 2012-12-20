VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FRM_EXCADJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjustment of Exicse As on Current Date"
   ClientHeight    =   7245
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8985
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   7215
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12726
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
      Begin VB.TextBox RG23APCESSDR 
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
         Left            =   6600
         TabIndex        =   3
         Top             =   1440
         Width           =   2010
      End
      Begin VB.TextBox RG23APCESSCR 
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
         Left            =   4080
         TabIndex        =   2
         Top             =   1440
         Width           =   2010
      End
      Begin VB.TextBox PLAPCESSDR 
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
         Left            =   6600
         TabIndex        =   17
         Top             =   5400
         Width           =   2010
      End
      Begin VB.TextBox PLAPCESSCR 
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
         Left            =   4080
         TabIndex        =   16
         Top             =   5400
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker txtdate 
         Height          =   375
         Left            =   7080
         TabIndex        =   38
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56426497
         CurrentDate     =   40786
      End
      Begin VB.TextBox rg23acencr 
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
         Left            =   4080
         TabIndex        =   0
         Top             =   1080
         Width           =   2010
      End
      Begin VB.TextBox rg23aedcscr 
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
         Left            =   4080
         TabIndex        =   4
         Top             =   1800
         Width           =   2010
      End
      Begin VB.TextBox rg23ahedcscr 
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
         Left            =   4080
         TabIndex        =   6
         Top             =   2160
         Width           =   2010
      End
      Begin VB.TextBox rg23ccencr 
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
         Left            =   4080
         TabIndex        =   8
         Top             =   3240
         Width           =   2010
      End
      Begin VB.TextBox rg23cedcscr 
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
         Left            =   4080
         TabIndex        =   10
         Top             =   3600
         Width           =   2010
      End
      Begin VB.TextBox rg23chedcscr 
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
         Left            =   4080
         TabIndex        =   12
         Top             =   3960
         Width           =   2010
      End
      Begin VB.TextBox plahedcscr 
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
         Left            =   4080
         TabIndex        =   20
         Top             =   6120
         Width           =   2010
      End
      Begin VB.TextBox placencr 
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
         Left            =   4080
         TabIndex        =   14
         Top             =   5040
         Width           =   2010
      End
      Begin VB.TextBox plaedcscr 
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
         Left            =   4080
         TabIndex        =   18
         Top             =   5760
         Width           =   2010
      End
      Begin VB.TextBox rg23acendr 
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
         Left            =   6600
         TabIndex        =   1
         Top             =   1080
         Width           =   2010
      End
      Begin VB.TextBox rg23aedcsdr 
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
         Left            =   6600
         TabIndex        =   5
         Top             =   1800
         Width           =   2010
      End
      Begin VB.TextBox rg23ahedcsdr 
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
         Left            =   6600
         TabIndex        =   7
         Top             =   2160
         Width           =   2010
      End
      Begin VB.TextBox rg23ccendr 
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
         Left            =   6600
         TabIndex        =   9
         Top             =   3240
         Width           =   2010
      End
      Begin VB.TextBox rg23cedcsdr 
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
         Left            =   6600
         TabIndex        =   11
         Top             =   3600
         Width           =   2010
      End
      Begin VB.TextBox rg23chedcsdr 
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
         Left            =   6600
         TabIndex        =   13
         Top             =   3960
         Width           =   2010
      End
      Begin VB.TextBox plahedcsdr 
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
         Left            =   6600
         TabIndex        =   21
         Top             =   6120
         Width           =   2010
      End
      Begin VB.TextBox placendr 
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
         Left            =   6600
         TabIndex        =   15
         Top             =   5040
         Width           =   2010
      End
      Begin VB.TextBox plaedcsdr 
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
         Left            =   6600
         TabIndex        =   19
         Top             =   5760
         Width           =   2010
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6480
         TabIndex        =   23
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
         Image           =   "FRM_EXCADJ.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7800
         TabIndex        =   24
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
         Image           =   "FRM_EXCADJ.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "P.CESS"
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
         TabIndex        =   42
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "P.CESS"
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
         TabIndex        =   41
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Debit"
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
         Index           =   13
         Left            =   6720
         TabIndex        =   40
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Credit"
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
         Left            =   4200
         TabIndex        =   39
         Top             =   600
         Width           =   1815
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
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1815
         Left            =   120
         Top             =   840
         Width           =   8775
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
         TabIndex        =   36
         Top             =   2160
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
         TabIndex        =   35
         Top             =   1800
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
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         Height          =   420
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8775
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cr / Db as on Date"
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
         Left            =   3600
         TabIndex        =   33
         Top             =   120
         Width           =   3375
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
         TabIndex        =   32
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1455
         Left            =   120
         Top             =   3000
         Width           =   8775
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
         TabIndex        =   31
         Top             =   3960
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
         TabIndex        =   30
         Top             =   3600
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
         TabIndex        =   29
         Top             =   3240
         Width           =   1455
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
         TabIndex        =   28
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1575
         Left            =   120
         Top             =   4920
         Width           =   8775
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
         TabIndex        =   27
         Top             =   6120
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
         TabIndex        =   26
         Top             =   5760
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
         Index           =   11
         Left            =   240
         TabIndex        =   25
         Top             =   5040
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FRM_EXCADJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  'Check Valid Code
  If Not IsNumeric(rg23acencr) Then MsgBox "Invalid Amount ": rg23acencr.SetFocus: Exit Sub
  If UNT_ISPAPPER = "Y" Then
    If Not IsNumeric(RG23APCESSCR) Then MsgBox "Invalid Amount ": RG23APCESSCR.SetFocus: Exit Sub
   Else
    RG23APCESSCR.Text = 0
  End If
  If Not IsNumeric(rg23aedcscr) Then MsgBox "Invalid Amount ": rg23aedcscr.SetFocus: Exit Sub
  If Not IsNumeric(rg23ahedcscr) Then MsgBox "Invalid Amount ": rg23ahedcscr.SetFocus: Exit Sub
  
  If Not IsNumeric(rg23ccencr) Then MsgBox "Invalid Amount ": rg23ccencr.SetFocus: Exit Sub
  If Not IsNumeric(rg23cedcscr) Then MsgBox "Invalid Amount ": rg2caedcscr.SetFocus: Exit Sub
  If Not IsNumeric(rg23chedcscr) Then MsgBox "Invalid Amount ": rg23chedcscr.SetFocus: Exit Sub
  
  If Not IsNumeric(placencr) Then MsgBox "Invalid Amount ": placencr.SetFocus: Exit Sub
  If UNT_ISPAPPER = "Y" Then
  If Not IsNumeric(PLAPCESSCR) Then MsgBox "Invalid Amount ": PLAPCESSCR.SetFocus: Exit Sub
  Else
  PLAPCESSCR.Text = 0
  End If
  If Not IsNumeric(plaedcscr) Then MsgBox "Invalid Amount ": plaedcscr.SetFocus: Exit Sub
  If Not IsNumeric(plahedcscr) Then MsgBox "Invalid Amount ": plahedcscr.SetFocus: Exit Sub
  
  If Not IsNumeric(rg23acendr) Then MsgBox "Invalid Amount ": rg23acendr.SetFocus: Exit Sub
  If UNT_ISPAPPER = "Y" Then
  If Not IsNumeric(RG23APCESSDR) Then MsgBox "Invalid Amount ": RG23APCESSDR.SetFocus: Exit Sub
  Else
  RG23APCESSDR.Text = 0
  End If
  If Not IsNumeric(rg23aedcsdr) Then MsgBox "Invalid Amount ": rg23aedcsdr.SetFocus: Exit Sub
  If Not IsNumeric(rg23ahedcsdr) Then MsgBox "Invalid Amount ": rg23ahedcsdr.SetFocus: Exit Sub
  
  If Not IsNumeric(rg23ccendr) Then MsgBox "Invalid Amount ": rg23ccendr.SetFocus: Exit Sub
  If Not IsNumeric(rg23cedcsdr) Then MsgBox "Invalid Amount ": rg2caedcsdr.SetFocus: Exit Sub
  If Not IsNumeric(rg23chedcsdr) Then MsgBox "Invalid Amount ": rg23chedcsdr.SetFocus: Exit Sub
  
  If Not IsNumeric(placendr) Then MsgBox "Invalid Amount ": placendr.SetFocus: Exit Sub
  If UNT_ISPAPPER = "Y" Then
  If Not IsNumeric(PLAPCESSDR) Then MsgBox "Invalid Amount ": PLAPCESSDR.SetFocus: Exit Sub
  Else
  PLAPCESSDR.Text = 0
  End If
  If Not IsNumeric(plaedcsdr) Then MsgBox "Invalid Amount ": plaedcsdr.SetFocus: Exit Sub
  If Not IsNumeric(plahedcsdr) Then MsgBox "Invalid Amount ": plahedcsdr.SetFocus: Exit Sub
  
  On Error GoTo LAST
  
  'Effect In EGpman
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and EXTRA5='Adjustment'"
  
  
  
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  
  
  'RG23-A II
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!UNIT = UNCD
  SAVDAT!VTYP = "EXC"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!dbcd = "RG23-A"
  SAVDAT!CRAC = ""
  SAVDAT!DRAC = ""
  SAVDAT!VBNO = "ACRRG23-A"
  SAVDAT!chln = "ACRRG23-A"
  SAVDAT!CHDT = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!TTYP = "RG23-A"
  SAVDAT!RECSTAT = "A"
  SAVDAT!CENVAT = Val(rg23acencr)
  SAVDAT!EDUCESS = Val(rg23aedcscr)
  SAVDAT!H_ED_CESS = Val(rg23ahedcscr)
  SAVDAT!CESS = Val(RG23APCESSCR)
  SAVDAT!EXTRA3 = "True"
  SAVDAT!EXTRA5 = "Adjustment"
  SAVDAT.Update
  
  
  'Debit Entry
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!UNIT = UNCD
  SAVDAT!VTYP = "EXD"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!dbcd = "RG23-A"
  SAVDAT!CRAC = ""
  SAVDAT!DRAC = ""
  SAVDAT!VBNO = "ADBRG23-A"
  SAVDAT!chln = "ADBRG23-A"
  SAVDAT!CHDT = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!TTYP = "RG23-A"
  SAVDAT!RECSTAT = "A"
  SAVDAT!CENVAT = Val(rg23acendr)
  SAVDAT!CESS = Val(RG23APCESSDR)
  SAVDAT!EDUCESS = Val(rg23aedcsdr)
  SAVDAT!H_ED_CESS = Val(rg23ahedcsdr)
  SAVDAT!EXTRA3 = "True"
  SAVDAT!EXTRA5 = "Adjustment"
  SAVDAT.Update
  
  
  
  'RG23-C II
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!UNIT = UNCD
  SAVDAT!VTYP = "EXC"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!dbcd = "RG23-C"
  SAVDAT!CRAC = ""
  SAVDAT!DRAC = ""
  SAVDAT!VBNO = "ACRRG23-C"
  SAVDAT!chln = "ACRRG23-C"
  SAVDAT!CHDT = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!TTYP = "RG23-C"
  SAVDAT!RECSTAT = "A"
  SAVDAT!CENVAT = Val(rg23ccencr)
  SAVDAT!EDUCESS = Val(rg23cedcscr)
  SAVDAT!H_ED_CESS = Val(rg23chedcscr)
  SAVDAT!EXTRA3 = "True"
  SAVDAT!EXTRA5 = "Adjustment"
  SAVDAT.Update
  
  
  'Debit Entry
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!UNIT = UNCD
  SAVDAT!VTYP = "EXD"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!dbcd = "RG23-C"
  SAVDAT!CRAC = ""
  SAVDAT!DRAC = ""
  SAVDAT!VBNO = "ADBRG23-C"
  SAVDAT!chln = "ADBRG23-C"
  SAVDAT!CHDT = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!TTYP = "RG23-C"
  SAVDAT!RECSTAT = "A"
  SAVDAT!CENVAT = Val(rg23ccendr)
  SAVDAT!EDUCESS = Val(rg23cedcsdr)
  SAVDAT!H_ED_CESS = Val(rg23chedcsdr)
  SAVDAT!EXTRA3 = "True"
  SAVDAT!EXTRA5 = "Adjustment"
  SAVDAT.Update
  
  
  
  'PLA
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!UNIT = UNCD
  SAVDAT!VTYP = "EXC"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!dbcd = "PLAREG"
  SAVDAT!CRAC = ""
  SAVDAT!DRAC = ""
  SAVDAT!VBNO = "ACRPLAREG"
  SAVDAT!chln = "ACRPLAREG"
  SAVDAT!CHDT = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!TTYP = "PLAREG"
  SAVDAT!RECSTAT = "A"
  SAVDAT!CENVAT = Val(placencr)
  SAVDAT!CESS = Val(PLAPCESSCR)
  SAVDAT!EDUCESS = Val(plaedcscr)
  SAVDAT!H_ED_CESS = Val(plahedcscr)
  SAVDAT!EXTRA3 = "True"
  SAVDAT!EXTRA5 = "Adjustment"
  SAVDAT.Update
  
  
  'Debit Entry
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!UNIT = UNCD
  SAVDAT!VTYP = "EXD"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!dbcd = "PLAREG"
  SAVDAT!CRAC = ""
  SAVDAT!DRAC = ""
  SAVDAT!VBNO = "ADBPLAREG"
  SAVDAT!chln = "ADBPLAREG"
  SAVDAT!CHDT = Format(TXTDATE, "YYYY/MM/DD")
  SAVDAT!TTYP = "PLAREG"
  SAVDAT!RECSTAT = "A"
  SAVDAT!CENVAT = Val(placendr)
  SAVDAT!CESS = Val(PLAPCESSDR)
  SAVDAT!EDUCESS = Val(plaedcsdr)
  SAVDAT!H_ED_CESS = Val(plahedcsdr)
  SAVDAT!EXTRA3 = "True"
  SAVDAT!EXTRA5 = "Adjustment"
  SAVDAT.Update
  
  
  
  
  
  
  
  MsgBox "Save Successful"
  Unload Me
  Exit Sub
LAST:
  MsgBox ERR.Description
  Resume
End Sub

Private Sub Form_Activate()
  If Allow_view_only = "Y" Then
     Unload Me
     Exit Sub
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call FILDTL
  Me.KeyPreview = True
  TXTDATE.MinDate = FSDT
  TXTDATE.MaxDate = FEDT
  If UNT_ISPAPPER = "Y" Then
    Label8(15).Visible = True
    Label8(14).Visible = True
    RG23APCESSCR.Visible = True
    RG23APCESSDR.Visible = True
    PLAPCESSCR.Visible = True
    PLAPCESSDR.Visible = True
   Else
    Label8(15).Visible = False
    Label8(14).Visible = False
    RG23APCESSCR.Visible = False
    RG23APCESSDR.Visible = False
    PLAPCESSCR.Visible = False
    PLAPCESSDR.Visible = False
  End If
End Sub




Private Sub FILDTL()
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and EXTRA5='Adjustment'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    TXTDATE = Now
  End If
  Do While Not RS.EOF
   TXTDATE.Value = RS!Date
   Select Case RS!TTYP
    
    
    Case "RG23-A"
     Select Case RS!VTYP
      Case "EXC"
       rg23acencr = RS!CENVAT
       RG23APCESSCR = RS!CESS
       rg23aedcscr = RS!EDUCESS
       rg23ahedcscr = RS!H_ED_CESS
      Case "EXD"
       rg23acendr = RS!CENVAT
       RG23APCESSDR = RS!CESS
       rg23aedcsdr = RS!EDUCESS
       rg23ahedcsdr = RS!H_ED_CESS
     End Select
    Case "RG23-C"
     Select Case RS!VTYP
      Case "EXC"
       rg23ccencr = RS!CENVAT
       rg23cedcscr = RS!EDUCESS
       rg23chedcscr = RS!H_ED_CESS
      Case "EXD"
       rg23ccendr = RS!CENVAT
       rg23cedcsdr = RS!EDUCESS
       rg23chedcsdr = RS!H_ED_CESS
     End Select
    Case "PLAREG"
     Select Case RS!VTYP
      Case "EXC"
       placencr = RS!CENVAT
       PLAPCESSCR = RS!CESS
       plaedcscr = RS!EDUCESS
       plahedcscr = RS!H_ED_CESS
      Case "EXD"
       placendr = RS!CENVAT
       PLAPCESSDR = RS!CESS
       plaedcsdr = RS!EDUCESS
       plahedcsdr = RS!H_ED_CESS
     End Select
   End Select
   RS.MoveNext
  Loop
End Sub


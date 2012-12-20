VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHB~1.OCX"
Begin VB.Form frmOutwardForJob 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Store Division Issue Raw Material for Jobwork / Returnable Gate Pass / Non-Returnable Gate Pass"
   ClientHeight    =   6750
   ClientLeft      =   375
   ClientTop       =   435
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11190
   Begin FramePlusCtl.FramePlus Frm1 
      Height          =   6735
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   11880
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
      Begin VB.TextBox TXTCOST 
         Appearance      =   0  'Flat
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox TXTISSCOPS 
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
         Left            =   9120
         MaxLength       =   10
         TabIndex        =   22
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox TXTCOPS 
         Appearance      =   0  'Flat
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
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   21
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox MERGE 
         Appearance      =   0  'Flat
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox TXTISSPCS 
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
         Left            =   9120
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox TXTPCS 
         Appearance      =   0  'Flat
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
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtParty 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtExpDays 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   25
         Top             =   5400
         Width           =   975
      End
      Begin VB.OptionButton optJob 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&JOB WORK"
         Enabled         =   0   'False
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
         Left            =   3720
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optRGP 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&RGP"
         Enabled         =   0   'False
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
         Left            =   6120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optNRGP 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&NRGP"
         Enabled         =   0   'False
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
         Left            =   7920
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TXTRMRK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   26
         Top             =   5400
         Width           =   6135
      End
      Begin VB.TextBox TXTISSQTY 
         Appearance      =   0  'Flat
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
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TXTSTOCK 
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TXTINAM 
         Appearance      =   0  'Flat
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox TXTIDNO 
         Appearance      =   0  'Flat
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
         Left            =   3240
         TabIndex        =   16
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox TXTVBNO 
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
         Left            =   9240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   35
         Top             =   1440
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   9240
         TabIndex        =   12
         Top             =   1800
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   18350081
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   6000
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
         Image           =   "frmOutwardForJob.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1680
         TabIndex        =   28
         Top             =   6000
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
         Image           =   "frmOutwardForJob.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   6000
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
         Image           =   "frmOutwardForJob.frx":1124
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   6000
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
         Image           =   "frmOutwardForJob.frx":1576
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDOK 
         Height          =   375
         Left            =   10200
         TabIndex        =   23
         Top             =   2640
         Width           =   855
         _ExtentX        =   1508
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
         Image           =   "frmOutwardForJob.frx":19C8
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDITMDEL 
         Height          =   375
         Left            =   10200
         TabIndex        =   31
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Remove"
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
      Begin MSFlexGridLib.MSFlexGrid ITMFLEX 
         Height          =   1335
         Left            =   240
         TabIndex        =   24
         Top             =   3840
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2355
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
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   5280
         TabIndex        =   3
         Top             =   6000
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
         Image           =   "frmOutwardForJob.frx":1D62
         cBack           =   -2147483633
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Head :"
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
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label LBLISSCOPS 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
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
         Height          =   195
         Left            =   8400
         TabIndex        =   43
         Top             =   3240
         Width           =   435
      End
      Begin VB.Label LBLSTKCOPS 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Cops Stock"
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
         Left            =   5040
         TabIndex        =   42
         Top             =   3240
         Width           =   990
      End
      Begin VB.Line Line8 
         X1              =   7680
         X2              =   7680
         Y1              =   2280
         Y2              =   3600
      End
      Begin VB.Label LBLMRGN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Merge No. :"
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
         Left            =   240
         TabIndex        =   41
         Top             =   3240
         Width           =   1020
      End
      Begin VB.Line Line13 
         X1              =   120
         X2              =   10080
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label LBLISSPCS 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Pcs/Box"
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
         Left            =   9120
         TabIndex        =   40
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label LBLSTKPCS 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock (Pcs)"
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
         Left            =   7800
         TabIndex        =   39
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Line Line12 
         X1              =   9000
         X2              =   9000
         Y1              =   2280
         Y2              =   3600
      End
      Begin VB.Label LBLFIFO 
         BackStyle       =   0  'Transparent
         Caption         =   "Note : Edit && Delete are not allowed.  (FIFO Is Applied in Unit Configuration)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   7680
         TabIndex        =   38
         Top             =   6000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   11160
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "&Party Name "
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
         TabIndex        =   8
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Expected Return Days :"
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
         TabIndex        =   32
         Top             =   5400
         Width           =   2040
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   11160
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of &Transaction    (A)                                   (B)                         (C)"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   6855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11160
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "&Remarks  :"
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
         TabIndex        =   33
         Top             =   5400
         Width           =   930
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   10080
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line9 
         X1              =   10080
         X2              =   10080
         Y1              =   2280
         Y2              =   3600
      End
      Begin VB.Line Line7 
         X1              =   6240
         X2              =   6240
         Y1              =   2280
         Y2              =   3600
      End
      Begin VB.Line Line6 
         X1              =   4800
         X2              =   4800
         Y1              =   2280
         Y2              =   3600
      End
      Begin VB.Line Line5 
         X1              =   3120
         X2              =   3120
         Y1              =   2280
         Y2              =   3120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   11160
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "&Quantity"
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
         Left            =   6600
         TabIndex        =   30
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label lblGRN_ID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Identi&fication No."
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
         Left            =   3240
         TabIndex        =   27
         Top             =   2400
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Item Description"
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
         TabIndex        =   13
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Stock"
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
         Left            =   5040
         TabIndex        =   29
         Top             =   2400
         Width           =   930
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11160
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label LBLNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Annexture No. "
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
         Left            =   7920
         TabIndex        =   37
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label LBLDT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Date :"
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
         Left            =   8640
         TabIndex        =   11
         Top             =   1800
         Width           =   540
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         Height          =   6495
         Left            =   120
         Top             =   120
         Width           =   11175
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Outward for JobWork / RGP / NRGP from Store Division"
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
         Left            =   2160
         TabIndex        =   36
         Top             =   240
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmOutwardForJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public M_DBCD As String
Public M_DVCD As String
Public M_DVNM As String
Public M_SRNO As String
Dim SAVEFLAG As Boolean
Dim ROWNO As Long
Dim SWITCH As Boolean
Public M_VTYP As String
Dim LASTSLIPNO As String, LASTVTYP As String
'-------------------------------------------------------------------------------------------
' FORM EVENTS
'-------------------------------------------------------------------------------------------

Private Sub cmdCancel_Click()
  ClsData (Me)
  ITMFLEX.Clear
  ITMFLEX.Rows = 2
  btn_sts (True)
  Call SETFLEX
  cmdAdd.SetFocus
  M_SRNO = Empty
  cmdOk.Caption = "&Add"
  SWITCH = False
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub CMDITMDEL_Click()
Dim CURSOR As Long
Dim J As Long

For J = ROWNO To ITMFLEX.Rows - 2

 ITMFLEX.TextMatrix(J, 0) = ITMFLEX.TextMatrix(J + 1, 0)
 ITMFLEX.TextMatrix(J, 1) = ITMFLEX.TextMatrix(J + 1, 1)
 ITMFLEX.TextMatrix(J, 2) = ITMFLEX.TextMatrix(J + 1, 2)
 ITMFLEX.TextMatrix(J, 3) = ITMFLEX.TextMatrix(J + 1, 3)
 ITMFLEX.TextMatrix(J, 4) = ITMFLEX.TextMatrix(J + 1, 4)
 ITMFLEX.TextMatrix(J, 5) = ITMFLEX.TextMatrix(J + 1, 5)
 ITMFLEX.TextMatrix(J, 6) = ITMFLEX.TextMatrix(J + 1, 6)
 ITMFLEX.TextMatrix(J, 7) = ITMFLEX.TextMatrix(J + 1, 7)
 ITMFLEX.TextMatrix(J, 8) = ITMFLEX.TextMatrix(J + 1, 8)
Next J

ITMFLEX.Rows = ITMFLEX.Rows - 1
Call CLEARDATA

If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
        TXTINAM.SetFocus
    Else
        TXTRMRK.Enabled = True: TXTRMRK.SetFocus
    End If

SWITCH = False
If TXTINAM.Enabled Then TXTINAM.SetFocus
cmdOk.Caption = "&Add"
CMDITMDEL.Enabled = False

End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

 If CHKSAVEDATA = True Then
    Exit Sub
 End If
   
 If SAVEFLAG = True Then
    Call SetChln
 End If
 
 Call SAVEISS
 
 If SAVEFLAG = True Then
    MsgBox "Your " & LBLNO & " is " + TXTVBNO.Text
 End If
    Call cmdCancel_Click
 Exit Sub
    
LAST:
    MsgBox ERR.Description
    If RS.State = 1 Then
        RS.CancelUpdate
    End If
    CN.RollbackTrans
    Exit Sub

End Sub

Private Sub cmdSavePrint_Click()
Call cmdSave_Click

If LASTVTYP = "RGP" Or LASTVTYP = "NGP" Then
    frmRPT_ReturanableGP.ONLINEPRINT = True
    frmRPT_ReturanableGP.M_VBNO = "'" & LASTSLIPNO & "'"
    frmRPT_ReturanableGP.Tag = LASTVTYP
    frmRPT_ReturanableGP.Visible = False
    Call frmRPT_ReturanableGP.cmdpreview_Click
End If
End Sub

Private Sub Form_Activate()
  btn_sts (True)
End Sub

Private Sub Form_Load()
 FIFOREQ = "Y"
 Call CenterChild(frm_Main, Me)
 Me.KeyPreview = True
 Me.Tag = zoomflag
 TXTVBDT = Date
 TXTVBDT.MaxDate = FEDT
 TXTVBDT.MinDate = FSDT
 Call SETFLEX
 CMDITMDEL.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
 End If
End Sub

' BUTTON EVENTS
'-------------------------------------------------------------------------------------------
Private Sub cmdAdd_Click()
    zoomflag = False
    btn_sts (False)
    M_SRNO = Empty
    Call SetChln
    SAVEFLAG = True
    txtExpDays.Enabled = True
    optJob.SetFocus
End Sub

Private Sub CMDOK_Click()
 Dim INDEX As Long
 
 If Not SWITCH Then
      ROWNO = ITMFLEX.Rows - 1
 End If
 
 If CheckData(ROWNO) Then Exit Sub
 
 
    ITMFLEX.TextMatrix(ROWNO, 0) = Trim(TXTINAM)
    ITMFLEX.TextMatrix(ROWNO, 1) = Trim(TXTIDNO)
    ITMFLEX.TextMatrix(ROWNO, 2) = Trim(TXTSTOCK)
    ITMFLEX.TextMatrix(ROWNO, 3) = Trim(nstr(Val(TXTISSQTY), 12, 3))
    ITMFLEX.TextMatrix(ROWNO, 4) = Trim(nstr(Val(TXTPCS), 9, 0))
    ITMFLEX.TextMatrix(ROWNO, 5) = Trim(nstr(Val(TXTISSPCS), 9, 0))
    ITMFLEX.TextMatrix(ROWNO, 6) = Trim(nstr(Val(txtCops), 9, 0))
    ITMFLEX.TextMatrix(ROWNO, 7) = Trim(nstr(Val(TXTISSCOPS), 9, 0))
    ITMFLEX.TextMatrix(ROWNO, 8) = Trim(MERGE)
    
    'CONDITION FOR ONLY ALLOW ONE ITEM
    If optJob.Value = True Then
       txtExpDays.Enabled = True: txtExpDays.SetFocus
       Call CLEARDATA
       cmdOk.Caption = "&Add"
       SWITCH = False
       Exit Sub
    End If
    
    If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
        TXTINAM.SetFocus
    Else
        txtExpDays.Enabled = True: txtExpDays.SetFocus
    End If
           
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
    Call CLEARDATA
    cmdOk.Caption = "&Add"
    SWITCH = False
End Sub
'-------------------------------------------------------------------------------------------
' LOCAL PROCEDURE
'-------------------------------------------------------------------------------------------
Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    TXTVBDT.Enabled = Not Yes
    txtParty.Enabled = Not Yes
    TXTCOST.Enabled = Not Yes
    optJob.Enabled = Not Yes
    optNRGP.Enabled = Not Yes
    optRGP.Enabled = Not Yes
    TXTINAM.Enabled = Not Yes
    TXTIDNO.Enabled = Not Yes
    TXTSTOCK.Enabled = Not Yes
    TXTISSQTY.Enabled = Not Yes
    txtExpDays.Enabled = Not Yes
    TXTRMRK.Enabled = Not Yes
End Sub
'-------------------------------------------------------------------------------------------

Private Sub ITMFLEX_Click()
   If ITMFLEX.Rows > 1 And ITMFLEX.TextMatrix(ITMFLEX.ROW, 0) <> Empty Then
    cmdOk.Caption = "Upd&ate"
    CMDITMDEL.Enabled = True
    ROWNO = ITMFLEX.ROW
    TXTINAM = ITMFLEX.TextMatrix(ROWNO, 0)
    TXTIDNO = ITMFLEX.TextMatrix(ROWNO, 1)
    TXTSTOCK = ITMFLEX.TextMatrix(ROWNO, 2)
    TXTISSQTY = ITMFLEX.TextMatrix(ROWNO, 3)
    TXTPCS = ITMFLEX.TextMatrix(ROWNO, 4)
    TXTISSPCS = ITMFLEX.TextMatrix(ROWNO, 5)
    txtCops = ITMFLEX.TextMatrix(ROWNO, 6)
    TXTISSCOPS = ITMFLEX.TextMatrix(ROWNO, 7)
    MERGE = ITMFLEX.TextMatrix(ROWNO, 8)
    
    SWITCH = True
  End If
    
   If Val(ITMFLEX.ROW) > 0 Then
      If TXTINAM.Enabled Then TXTINAM.SetFocus
   End If
   
End Sub

Private Sub MERGE_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        MERGE = Empty
    ElseIf KeyCode = vbKeyF2 Or MERGE = Empty Then
        M_DESC = Empty
        Key = Empty
        NEW_VISIBLE = False
        MERGE = SearchList1("Select DISTINCT MRGN,MRGN  From MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD = '" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "'", 0, Empty, "Select MERGE FROM MASTER")
        'Me.Tag = Key
        'MERGE = Key
    End If
  
  
  'TXTSTOCK = GetItemStock(TXTICOD)
  'TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
  Call FindSpecification
  Call StockDisplay
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
  Me.KeyPreview = True


End Sub

'-------------------------------------------------------------------------------------------
' CODE FOR CURSOR POSITION ON MODULE
'-------------------------------------------------------------------------------------------

Private Sub optJob_Click()
  Call SetChln
End Sub

Private Sub optNRGP_Click()
  Call SetChln
  If optNRGP.Value = True Then
     txtExpDays.Enabled = False
  Else
     txtExpDays.Enabled = True
  End If
End Sub

Private Sub optRGP_Click()
  Call SetChln
End Sub

Private Sub TXTCOST_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTCOST = Empty
    ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTCOST = Empty) Then
        M_DESC = Empty
        NEW_VISIBLE = True
        TXTCOST = SearchList1("Select  TOP 20 Code,Name From REFMST WHERE CATA='N' AND NAME NOT LIKE '%DISABLE%'", 0, Empty, "Select COSTING HEAD FROM MASTER")
        If key_PressNew = True Then
            M_DESC = ""
            Ref_Cat = "N"
            LOAD Frm_Ref_FAS
            Frm_Ref_FAS.Tag = Ref_Cat
            Frm_Ref_FAS.Show
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub txtExpDays_GotFocus()
   txtExpDays.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtExpDays_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, txtExpDays, Me, False) = 0 Then KeyAscii = 0
End Sub

Private Sub txtExpDays_LostFocus()
   txtExpDays.BackColor = vbWhite
End Sub

Private Sub TXTIDNO_GotFocus()
   TXTIDNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTIDNO_LostFocus()
   TXTIDNO.BackColor = vbWhite
End Sub

Private Sub TXTINAM_Change()
Call STOCKCLEAR
End Sub

Private Sub TXTISSQTY_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTISSQTY, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtParty_GotFocus()
  txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or Trim(txtParty.Text) = Empty Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtParty.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, txtParty, "List of Party")
        txtParty.Tag = Key
    ElseIf KeyCode = vbKeyDelete Then
        txtParty = Empty
    End If
    Me.KeyPreview = True
End Sub

Private Sub txtParty_LostFocus()
  txtParty.BackColor = vbWhite
End Sub

Private Sub TXTINAM_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
If KeyCode = vbKeyF2 Or (Trim(TXTINAM) = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: Key = Empty
   TXTINAM.Text = SearchITEMLIST("SELECT TOP 20 CODE, NAME FROM ITMMST", 0, TXTINAM.Text, "SELECT ITEM FROM LIST")
   TXTSTOCK = GetItemStock(Key)
   TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
   TXTINAM.Tag = Key
End If
Call FindSpecification
Call StockDisplay
Me.KeyPreview = True
End Sub

Private Sub txtINAM_GotFocus()
 TXTINAM.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub txtINAM_LostFocus()
 TXTINAM.BackColor = vbWhite
End Sub

Private Sub TXTRMRK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  'SendKeys "{TAB}"
End If
End Sub

Private Sub TXTSTOCK_GotFocus()
 TXTSTOCK.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTSTOCK_LostFocus()
 TXTSTOCK.BackColor = vbWhite
End Sub

Private Sub TXTISSQTY_GotFocus()
 TXTISSQTY.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTISSQTY_LostFocus()
 TXTISSQTY.BackColor = vbWhite
End Sub

Private Sub TXTRMRK_GotFocus()
 TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRMRK_LostFocus()
 TXTRMRK.BackColor = vbWhite
End Sub

Private Sub SETFLEX()
  ITMFLEX.Clear
  ITMFLEX.ColWidth(0) = 2600
  ITMFLEX.ColWidth(1) = 1500
  ITMFLEX.ColWidth(2) = 1300
  ITMFLEX.ColWidth(3) = 1200
  ITMFLEX.ColWidth(4) = 1000
  ITMFLEX.ColWidth(5) = 900
  ITMFLEX.ColWidth(6) = 1000
  ITMFLEX.ColWidth(7) = 900
  ITMFLEX.ColWidth(8) = 1000
  ITMFLEX.ColWidth(9) = 0
  
  ITMFLEX.Clear
  ITMFLEX.TextMatrix(0, 0) = "Item Description"
  ITMFLEX.TextMatrix(0, 1) = "Identification No"
  ITMFLEX.TextMatrix(0, 2) = "Item Stock"
  ITMFLEX.TextMatrix(0, 3) = "Quantity"
  ITMFLEX.TextMatrix(0, 4) = "Pcs/Stock"
  ITMFLEX.TextMatrix(0, 5) = "Pcs."
  ITMFLEX.TextMatrix(0, 6) = "Cops Stock"
  ITMFLEX.TextMatrix(0, 7) = "Cops"
  ITMFLEX.TextMatrix(0, 8) = "MRGN"
  ITMFLEX.TextMatrix(0, 9) = "Amount"
  
  ITMFLEX.ColAlignment(0) = vbLeftJustify
  ITMFLEX.ColAlignment(1) = vbLeftJustify
  ITMFLEX.ColAlignment(2) = vbRightJustify
  ITMFLEX.ColAlignment(3) = vbRightJustify
  
  End Sub

Private Sub CLEARDATA()
        TXTIDNO = Empty
        TXTINAM = Empty
        TXTSTOCK = Empty
        TXTISSQTY = Empty
        
        MERGE = Empty
        TXTPCS = Empty
        TXTISSPCS = Empty
        txtCops = Empty
        TXTISSCOPS = Empty
        LBLSTKPCS.Enabled = True
        LBLSTKCOPS.Enabled = True
        LBLISSPCS.Enabled = True
        LBLISSCOPS.Enabled = True
        LBLMRGN.Enabled = True
        MERGE.Enabled = True
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
    If Trim(TXTINAM) = Empty Then
        MsgBox "Please Select Items From List !!", vbInformation
        'TXTINAM.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If Val(TXTISSQTY) < 1 Then
        MsgBox "Please Enter Valid Quantity !!", vbInformation, "Quantity Missing !!"
        TXTISSQTY.SetFocus
        CheckData = True
        Exit Function
    End If
                
    If Val(TXTSTOCK) < Val(TXTISSQTY) Then
        MsgBox "Stock Doesn't Support !!", vbInformation, "Stock Exceed !!"
        TXTISSQTY.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If MERGE.Enabled = True Then
    If Trim(MERGE) = Empty Then
        MsgBox "Please Select Merge No. From List !!", vbInformation
        If MERGE.Enabled Then MERGE.SetFocus
        CheckData = True
        Exit Function
    End If
    End If

    If Val(TXTPCS) < Val(TXTISSPCS) Then
        MsgBox "Stock Doesn't Support !!", vbInformation, "Stock Exceed !!"
       ' TXTISSQTY.SetFocus
        CheckData = True
        Exit Function
    End If

    If Val(txtCops) < Val(TXTISSCOPS) Then
        MsgBox "Stock Doesn't Support !!", vbInformation, "Stock Exceed !!"
       ' TXTISSQTY.SetFocus
        CheckData = True
        Exit Function
    End If

    

    
    For INDEX = 1 To ITMFLEX.Rows - 1
        If Trim(ITMFLEX.TextMatrix(INDEX, 0)) = TXTIDNO And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
           MsgBox "Invalid Item Detail"
           TXTINAM.SetFocus
           CheckData = True
           Exit Function
        End If
    Next INDEX
    
    'For Negative Stock
    'FIFO : NRGP
    If SAVEFLAG And FIFOREQ = "Y" And optNRGP.Value = True Then
       If Trim(TXTIDNO) = "" Then
            MsgBox "GRN NO. SHOULD NOT EMPTY !!! "
            TXTIDNO.SetFocus
            Exit Function
       End If
       
       Dim BLQTY As Double
       BLQTY = BALQNTY(TXTINAM, Trim(TXTIDNO))
       If TXTISSQTY > BLQTY Then
          MsgBox "Balanced Qnty. is " & CStr(BLQTY) & " / GRN Not Exist!!!"
          TXTISSQTY.SetFocus
          CheckData = True
          Exit Function
       End If
    End If
    '--------------------------------------
    
End Function


Private Function CHKSAVEDATA() As Boolean
If txtParty = Empty Then
  MsgBox "Enter Party then Save"
  CHKSAVEDATA = True
  If txtParty.Enabled Then txtParty.SetFocus
  Exit Function
End If

If txtExpDays = Empty Then
  MsgBox "Enter Valid Expected Days !!!", vbInformation
  CHKSAVEDATA = True
  If txtExpDays.Enabled Then txtExpDays.SetFocus
  Exit Function
End If

If ITMFLEX.TextMatrix(1, 0) = Empty Then
  MsgBox "Enter Data then Save"
  CHKSAVEDATA = True
  TXTINAM.Enabled = True
  TXTINAM.SetFocus
  Exit Function
End If

'FIFO
If SAVEFLAG And FIFOREQ = "Y" And optNRGP.Value = True Then
   If Not ALLOWFLEX Then
      CHKSAVEDATA = True
      Exit Function
   End If
End If
'------------------------

End Function

Private Sub SAVEISS()
  
  On Error GoTo LAST
  Dim SQL As String
  
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
   
  CN.BeginTrans
  Call DELETEISS
  SQL = Empty
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' AND  VTYP='" & M_VTYP & "' AND DBCD='000001' AND VBNO='" & Trim(TXTVBNO) & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  
  Dim AI As String
  Dim BQ As Double
  Dim RATE As Double
  Dim i As Long
  Dim DVCOD As String
      
  i = 1
  Dim FIFORATE As Double
  
  For i = 1 To ITMFLEX.Rows - 1
    If ITMFLEX.TextMatrix(i, 0) <> Empty Then
    SAVDAT.AddNew
    
     '-----------------------------------------------------------------------------------
     'FIFO
      If SAVEFLAG = True And FIFOREQ = "Y" Then
         FIFORATE = FindFIFORate(ITMFLEX.TextMatrix(i, 0), Val(ITMFLEX.TextMatrix(i, 3)), i)
      End If
     '-------------------------------------------------------------------------------------
    
     SAVDAT!COMP = compPth
     SAVDAT!VTYP = M_VTYP
     SAVDAT!SRNO = 0
     SAVDAT!SRCH = i
     SAVDAT!VBNO = TXTVBNO.Text
     LASTSLIPNO = TXTVBNO.Text
     LASTVTYP = M_VTYP
     SAVDAT!chln = TXTVBNO
     SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
     SAVDAT!dbcd = M_DBCD
     AI = GetCode("ITMMST", Trim(ITMFLEX.TextMatrix(i, 0)), "NAME", "CODE")
     SAVDAT!ICOD = AI
     SAVDAT!PCES = Val(ITMFLEX.TextMatrix(i, 5))
     SAVDAT!QNTY = Val(ITMFLEX.TextMatrix(i, 3)): BQ = Val(ITMFLEX.TextMatrix(i, 3))
     SAVDAT!COPS = Val(ITMFLEX.TextMatrix(i, 7))
     SAVDAT!MRGN = Trim(ITMFLEX.TextMatrix(i, 8))
     SAVDAT!ltno = Trim(ITMFLEX.TextMatrix(i, 8))
    'FIFO
     If SAVEFLAG = True And FIFOREQ = "Y" Then
        SAVDAT!RATE = FIFORATE
     Else
       SAVDAT!RATE = 0
       RATE = 0
     End If
    
     SAVDAT!AMNT = SAVDAT!RATE * Val(ITMFLEX.TextMatrix(i, 3))
     SAVDAT!QORP = "Q"
     SAVDAT![User] = cUName
     If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
     End If
     SAVDAT!OPER = "-"
     SAVDAT!PCOD = GetCode("ACCMST", txtParty, "NAME", "CODE")
     SAVDAT!DVCD = "000001"
    
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
    SAVDAT!SPECIFICATION = GetSpeci(AI)
    SAVDAT!EXTRA2 = Trim(TXTRMRK)
    
    SAVDAT.Update
    
   End If
  Next
  
  
 'JOBOUT
 If SAVDAT.State = 1 Then SAVDAT.Close
 SAVDAT.Open "SELECT * FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND  VTYP='" & M_VTYP & _
             "' AND DBCD='" & M_DBCD & "' AND VBNO='" & Trim(TXTVBNO) & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
 i = 1
  For i = 1 To ITMFLEX.Rows - 1
    If ITMFLEX.TextMatrix(i, 0) <> Empty Then
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!unit = UNCD
    SAVDAT!VTYP = M_VTYP
    SAVDAT!dbcd = M_DBCD
    SAVDAT!VBNO = TXTVBNO
    SAVDAT!SRCH = i
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!PCOD = GetCode("ACCMST", txtParty, "NAME", "CODE")
    SAVDAT!ICOD = GetCode("ITMMST", Trim(ITMFLEX.TextMatrix(i, 0)), "NAME", "CODE")
    SAVDAT!IDNO = ITMFLEX.TextMatrix(i, 1)
    SAVDAT!PCES = Val(ITMFLEX.TextMatrix(i, 5))
    SAVDAT!QNTY = Val(ITMFLEX.TextMatrix(i, 3))
    SAVDAT!COPS = Val(ITMFLEX.TextMatrix(i, 7))
    SAVDAT!ltno = Trim(ITMFLEX.TextMatrix(i, 8))
    
    If Val(ITMFLEX.TextMatrix(i, 3)) > 0 Then
       RATE = Val(ITMFLEX.TextMatrix(i, 9)) / Val(ITMFLEX.TextMatrix(i, 3))
       SAVDAT!RATE = RATE
    Else
       SAVDAT!RATE = 0
       RATE = 0
    End If
    SAVDAT!RATE = RATE
    SAVDAT!AMNT = Val(ITMFLEX.TextMatrix(i, 9))
    SAVDAT!OPER = "+"
    SAVDAT!GRNNO = ITMFLEX.TextMatrix(i, 1)
    SAVDAT!XDAYS = txtExpDays
    SAVDAT!RMRK = Trim(TXTRMRK)
    SAVDAT!User = cUName
    SAVDAT!SYSR = "N"
    SAVDAT!RECSTAT = "A"
    SAVDAT.Update
  End If
 Next
 
 'FIFO
    If SAVEFLAG = True And FIFOREQ = "Y" Then
       If optNRGP.Value = True Then
          Call SetFIFOConsumptionNRGP
       Else
          Call SetFIFOConsumptionRGP
       End If
    End If
 '----------------------
    
 'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
    Call SetSRNO(TXTVBNO, M_VTYP, M_DBCD)
  End If
'--------------------------------
'DAILYSTATUS ENTRY
  Call DAILYSTATUS(M_VTYP, GetCode("ACCMST", txtParty, "NAME", "CODE"), M_DBCD, Val(ITMFLEX.TextMatrix(1, 3)), TXTVBNO, Val(ITMFLEX.TextMatrix(1, 4)), cUName, "N", Now, TXTVBDT)
'--------------------------------
  CN.CommitTrans
  Exit Sub
LAST:
 MsgBox ERR.Description
 If SAVDAT.State = 1 Then
 Resume
   SAVDAT.CancelUpdate
   SAVDAT.Close
 End If
 CN.RollbackTrans
End Sub


Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = M_VTYP
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = 0
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = 0
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

Private Sub DELETEISS()
  Dim L As Long
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND  VTYP='" & M_VTYP & "' AND DBCD='000001' AND VBNO='" & Trim(TXTVBNO) & "' AND RECSTAT<>'D'", L
  CN.Execute "DELETE FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND  VTYP='" & M_VTYP & "' AND DBCD='000001' AND VBNO='" & Trim(TXTVBNO) & "' AND RECSTAT<>'D'", L
End Sub

Private Sub SetChln()
On Error GoTo LAST

M_DBCD = "000001"
If optJob.Value = True Then
   LBLNO = "Annexture No. "
   M_VTYP = "ANX"
   TXTVBNO = GenVNO("ANX", M_DBCD)
   If ITMFLEX.Rows - 1 > 1 Then ITMFLEX.Rows = 2
   lblGRN_ID.Caption = "Identi&fication No."
   ITMFLEX.TextMatrix(0, 1) = "Identification No"
ElseIf optRGP.Value = True Then
   LBLNO = " Challan No. "
   M_VTYP = "RGP"
   TXTVBNO = GenVNO("RGP", M_DBCD)
   lblGRN_ID.Caption = "Identi&fication No."
   ITMFLEX.TextMatrix(0, 1) = "Identification No"
ElseIf optNRGP.Value = True Then
   LBLNO = "    Slip No. "
   M_VTYP = "NGP"
   TXTVBNO = GenVNO("NGP", M_DBCD)
   lblGRN_ID.Caption = "G&RN No."
   ITMFLEX.TextMatrix(0, 1) = "GRN No."
End If

If Format(Now, "DD/MM/YYYY") <= FEDT And Format(Now, "DD/MM/YYYY") > FSDT Then
   TXTVBDT = Now
Else
   MsgBox "Check System Date and Transaction Date"
End If
     
   Exit Sub
LAST:
   MsgBox ERR.Description
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

'FIFO----------------------
Private Sub SetFIFOConsumptionNRGP()
On Error GoTo FIFOERR

'VARIABLE DECLARATION
Dim ICOD As String, Item As String, INDEX As Long
Dim BALQNTY As Double, TMPQTY As Double
Dim FIFORS As ADODB.Recordset: Set FIFORS = New ADODB.Recordset
Dim MRGN As String
'-------------------------------------------------------------
For INDEX = 1 To ITMFLEX.Rows - 1
'-------------------------------------------------------------
'INITIALISE
 Item = ITMFLEX.TextMatrix(INDEX, 0)
 BALQNTY = Val(ITMFLEX.TextMatrix(INDEX, 3))
 MRGN = Trim(ITMFLEX.TextMatrix(INDEX, 8))
'-------------------------------------------------------------

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then ICOD = Trim(FIFORS!CODE & "")
FIFORS.Close

If FIFORS.State = 1 Then FIFORS.Close
If MRGN = Empty Then
   FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
Else
   FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND MRGN = '" & MRGN & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
End If
  
If Not FIFORS.EOF Then
Do While Not FIFORS.EOF
        
        TMPQTY = Val(FIFORS!BAL_QNTY)
            
        If BALQNTY > TMPQTY Then
           FIFORS!RET_PTY_QNTY = Val(FIFORS!RET_PTY_QNTY) + TMPQTY
           FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
           FIFORS.Update
           BALQNTY = BALQNTY - TMPQTY
        ElseIf BALQNTY > 0 Or BALQNTY = TMPQTY Then
           FIFORS!RET_PTY_QNTY = Val(FIFORS!RET_PTY_QNTY) + BALQNTY
           FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
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

Private Function BALQNTY(Item As String, MRN As String) As Double
On Error GoTo ALLOWERR
Dim ICOD As String
Dim ALLOWRS As ADODB.Recordset
Set ALLOWRS = New ADODB.Recordset
    
BALQNTY = 0
    
If ALLOWRS.State = 1 Then ALLOWRS.Close
ALLOWRS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not ALLOWRS.EOF Then ICOD = Trim(ALLOWRS!CODE & "")
ALLOWRS.Close
  
If ALLOWRS.State = 1 Then ALLOWRS.Close
ALLOWRS.Open "SELECT ISNULL(SUM(BAL_QNTY),0) AS BAL_QNTY FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VBNO='" & MRN & "' AND ICOD='" & ICOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not ALLOWRS.EOF Then
   BALQNTY = Val(ALLOWRS!BAL_QNTY)
End If
    
Exit Function
ALLOWERR:
MsgBox ERR.Description
End Function

Private Function ALLOWFLEX() As Boolean
On Error GoTo ALLOWERR
Dim BLQTY As Double, INDEX As Long
 
ALLOWFLEX = True
 
 For INDEX = 1 To ITMFLEX.Rows - 1
     If Trim(ITMFLEX.TextMatrix(INDEX, 0)) = Empty Or Trim(ITMFLEX.TextMatrix(INDEX, 1)) = Empty Then
        ITMFLEX.SetFocus
        ALLOWFLEX = False
        Exit Function
     End If
     
     BLQTY = BALQNTY(ITMFLEX.TextMatrix(INDEX, 0), ITMFLEX.TextMatrix(INDEX, 1))
     If Val(ITMFLEX.TextMatrix(ROWNO, 3)) > BLQTY Then
        MsgBox "Item Balanced Qnty. is " & CStr(BLQTY) & " !!! "
        ITMFLEX.ROW = INDEX
        ITMFLEX.COL = 4
        ITMFLEX.SetFocus
        ALLOWFLEX = False
        Exit Function
     End If
 Next INDEX
 
Exit Function
ALLOWERR:
MsgBox ERR.Description
End Function

'FIFO
Private Function FindFIFORate(Item As String, QNTY As Double, i As Long) As Double
On Error GoTo FIFOERR
Dim ICOD As String
Dim Top As Double
Dim Bottom As Double
Dim BALQNTY As Double
Dim FIFORS As ADODB.Recordset
Set FIFORS = New ADODB.Recordset
Dim MRGN As String

FindFIFORate = 0
BALQNTY = QNTY

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then ICOD = Trim(FIFORS!CODE & "")
FIFORS.Close


If FIFORS.State = 1 Then FIFORS.Close
If Trim(ITMFLEX.TextMatrix(i, 8)) = Empty Then
   FIFORS.Open "SELECT BAL_QNTY AS QNTY,RATE,NETRATE FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
Else
   FIFORS.Open "SELECT BAL_QNTY AS QNTY,RATE,NETRATE FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND MRGN = '" & Trim(ITMFLEX.TextMatrix(i, 8)) & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
End If

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

'FIFO---------------------------------------------------------
 Private Sub SetFIFOConsumptionRGP()
 On Error GoTo FIFOERR

'VARIABLE DECLARATION
  Dim ICOD As String, Item As String, INDEX As Long
  Dim BALQNTY As Double, TMPQTY As Double
  Dim FIFORS As ADODB.Recordset: Set FIFORS = New ADODB.Recordset
  Dim MRGN As String
 '-------------------------------------------------------------
  For INDEX = 1 To ITMFLEX.Rows - 1
 '-------------------------------------------------------------
 'INITIALISE
  Item = ITMFLEX.TextMatrix(INDEX, 0)
  BALQNTY = Val(ITMFLEX.TextMatrix(INDEX, 3))
  MRGN = Trim(ITMFLEX.TextMatrix(INDEX, 8))
 '-------------------------------------------------------------

If FIFORS.State = 1 Then FIFORS.Close
   FIFORS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then ICOD = Trim(FIFORS!CODE & "")
FIFORS.Close

If FIFORS.State = 1 Then FIFORS.Close
If MRGN = Empty Then
    FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
Else
    FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND  MRGN = '" & MRGN & "'  AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
End If

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

Private Function GetItemStock(ICOD As String) As Double
Dim DIVISIONCODE As String
DIVISIONCODE = "000001"

Dim STKRS As ADODB.Recordset
Set STKRS = New ADODB.Recordset
Dim ISSQTY As Double
Dim RTIQTY As Double

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND ICOD='" & ICOD & "' AND OPER='+' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   ISSQTY = STKRS!QNTY
Else
   ISSQTY = 0
End If

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND ICOD='" & ICOD & "' AND OPER='-' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   RTIQTY = STKRS!QNTY
Else
   RTIQTY = 0
End If

GetItemStock = ISSQTY - RTIQTY

End Function

Public Sub FindSpecification()
'----------------------------------------------------------------------
    'SPECIFICATION ACCORDING ITEM GROUP
    Dim RS As New ADODB.Recordset
    Dim SPECI As String
    Dim MRGN As String
    Dim igcd As String
    If RS.State = 1 Then RS.Close
       RS.Open "SELECT *  FROM ITMMST WHERE CODE = '" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       igcd = RS!igcd
    End If
    If RS.State = 1 Then RS.Close
       RS.Open "SELECT * FROM IGMMST WHERE CODE = '" & igcd & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       SPECI = RS!SPECIFICATION
       MRGN = RS!MERGE
    End If
    
    
    If MRGN = "Y" Then
       MERGE.Enabled = True
       LBLMRGN.Enabled = True
      If MERGE = Empty And MERGE.Enabled = True Then MERGE.SetFocus
    Else
       MERGE.Enabled = False
       LBLMRGN.Enabled = False
    End If
       
    If Val(SPECI) = 0 Then
       TXTISSPCS.Enabled = True
       TXTISSQTY.Enabled = True
       TXTISSCOPS.Enabled = False
       LBLSTKCOPS.Enabled = False
       LBLISSCOPS.Enabled = False
    ElseIf Val(SPECI) = 1 Then
       TXTISSQTY.Enabled = True
       TXTISSPCS.Enabled = False
       TXTISSCOPS.Enabled = False
       LBLSTKPCS.Enabled = False
       LBLISSPCS.Enabled = False
       LBLSTKCOPS.Enabled = False
       LBLISSCOPS.Enabled = False
    ElseIf Val(SPECI) = 2 Then
       TXTISSPCS.Enabled = True
       TXTISSCOPS.Enabled = True
       TXTISSQTY.Enabled = True
       LBLSTKCOPS.Enabled = True
       LBLISSCOPS.Enabled = True
       LBLISSPCS.Enabled = True
    ElseIf Val(SPECI) = 3 Then
       TXTISSCOPS.Enabled = True
       TXTISSQTY.Enabled = True
       LBLSTKCOPS.Enabled = True
       LBLISSCOPS.Enabled = True
       TXTISSPCS.Enabled = False
       LBLSTKPCS.Enabled = False
       LBLISSPCS.Enabled = False
    End If
End Sub

Private Sub StockDisplay()

Dim RS As New ADODB.Recordset
Dim SPECI As String
Dim MRGN As String
Dim igcd As String
Dim TXTICOD As String
If RS.State = 1 Then RS.Close
    RS.Open "SELECT *  FROM ITMMST WHERE CODE = '" & GetCode("ITMMST", TXTINAM, "NAME", "CODE") & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
    igcd = RS!igcd
End If

If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM IGMMST WHERE CODE = '" & igcd & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
    SPECI = RS!SPECIFICATION
    MRGN = RS!MERGE
End If

Call FindSpecification

    TXTICOD = GetCode("ITMMST", TXTINAM, "NAME", "CODE")
 '------------------------------------------------------------------
  '1. 'SPECIFICATION BOX/PCS + QUANTITY AND MERGENO/WITHOUT MERGENO
  
     If Val(SPECI) = 0 And MRGN = "Y" Then
        TXTSTOCK = GetItemStockMrgn(TXTICOD)
        TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
        TXTPCS = GetItemStockPcsMrgn(TXTICOD)
        TXTPCS = nstr(Val(TXTPCS), 9, 0)
     ElseIf Val(SPECI) = 0 And MRGN <> "Y" Then
        TXTSTOCK = GetItemStock(TXTICOD)
        TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
        TXTPCS = GetItemStockPcs(TXTICOD)
        TXTPCS = nstr(Val(TXTPCS), 9, 0)
     End If
 '------------------------------------------------------------------
 '2. 'SPECIFICATION QUANTITY

     If Val(SPECI) = 1 And MRGN = "Y" Then
        TXTSTOCK = GetItemStockMrgn(TXTICOD)
        TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
     ElseIf Val(SPECI) = 1 And MRGN <> "Y" Then
        TXTSTOCK = GetItemStock(TXTICOD)
        TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
     End If
 '-------------------------------------------------------------------
 '3 'SPECIFICATION COPS + BOX/PCS + QUANTITY
    
    
    If Val(SPECI) = 2 And MRGN = "Y" Then
       TXTSTOCK = GetItemStockMrgn(TXTICOD)
       TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
       TXTPCS = GetItemStockPcsMrgn(TXTICOD)
       TXTPCS = nstr(Val(TXTPCS), 9, 0)
       txtCops = GetItemStockCopsMrgn(TXTICOD)
       txtCops = nstr(Val(txtCops), 9, 0)
    ElseIf Val(SPECI) = 2 And MRGN <> "Y" Then
       TXTSTOCK = GetItemStock(TXTICOD)
       TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
       TXTPCS = GetItemStockPcs(TXTICOD)
       TXTPCS = nstr(Val(TXTPCS), 9, 0)
       txtCops = GetItemStockCops(TXTICOD)
       txtCops = nstr(Val(txtCops), 9, 0)
    End If
 '---------------------------------------------------------------------
 '4 SPECIFICATION COPS + QUANTITY
 
    
    If Val(SPECI) = 3 And MRGN = "Y" Then
       TXTSTOCK = GetItemStockMrgn(TXTICOD)
       TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
       txtCops = GetItemStockCopsMrgn(TXTICOD)
       txtCops = nstr(Val(txtCops), 9, 0)
    ElseIf Val(SPECI) = 3 And MRGN <> "Y" Then
       TXTSTOCK = GetItemStock(TXTICOD)
       TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
       txtCops = GetItemStockCops(TXTICOD)
       txtCops = nstr(Val(txtCops), 9, 0)
    End If
    
End Sub

Private Function GetItemStockPcs(ICOD As String) As Double
Dim DIVISIONCODE As String
DIVISIONCODE = "000001"

Dim STKRS As ADODB.Recordset
Set STKRS = New ADODB.Recordset

Dim ISSPCS As Double
Dim RTIPCS As Double

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(PCES),0) AS PCS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "'  AND ICOD='" & ICOD & "' AND OPER='+' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   ISSPCS = STKRS!PCS
Else
   ISSPCS = 0
End If

If STKRS.State = 1 Then STKRS.Close
STKRS.Open " SELECT ISNULL(SUM(PCES),0) AS PCS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND  ICOD='" & ICOD & "' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   RTIPCS = STKRS!PCS
Else
   RTIPCS = 0
End If

GetItemStockPcs = ISSPCS - RTIPCS

TXTVBDT.Enabled = False

End Function

Private Function GetItemStockCops(ICOD As String) As Double
Dim DIVISIONCODE As String
DIVISIONCODE = "000001"

Dim STKRS As ADODB.Recordset
Set STKRS = New ADODB.Recordset

Dim RTICOPS As Double
Dim ISSCOPS As Double


If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(COPS),0) AS COPS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "'  AND ICOD='" & ICOD & "' AND OPER='+' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   ISSCOPS = STKRS!COPS
Else
   ISSCOPS = 0
End If

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT  ISNULL(SUM(COPS),0) AS COPS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND  ICOD='" & ICOD & "' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   RTICOPS = STKRS!COPS
Else
   RTICOPS = 0
End If

GetItemStockCops = ISSCOPS - RTICOPS

TXTVBDT.Enabled = False

End Function

Private Function GetItemStockCopsMrgn(ICOD As String) As Double
Dim DIVISIONCODE As String
DIVISIONCODE = "000001"

Dim STKRS As ADODB.Recordset
Set STKRS = New ADODB.Recordset

Dim RTICOPS As Double
Dim ISSCOPS As Double


If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(COPS),0) AS COPS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND MRGN = '" & MERGE & "' AND ICOD='" & ICOD & "' AND OPER='+' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   ISSCOPS = STKRS!COPS
Else
   ISSCOPS = 0
End If

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT  ISNULL(SUM(COPS),0) AS COPS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND MRGN = '" & MERGE & "' AND ICOD='" & ICOD & "' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   RTICOPS = STKRS!COPS
Else
   RTICOPS = 0
End If

GetItemStockCopsMrgn = ISSCOPS - RTICOPS

TXTVBDT.Enabled = False

End Function


Private Function GetItemStockPcsMrgn(ICOD As String) As Double
Dim DIVISIONCODE As String
DIVISIONCODE = "000001"

Dim STKRS As ADODB.Recordset
Set STKRS = New ADODB.Recordset
Dim ISSQTY As Double
Dim ISSPCS As Double
Dim RTIPCS As Double
Dim RTICOPS As Double
Dim ISSCOPS As Double
Dim RTIQTY As Double

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(PCES),0) AS PCS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "'  AND MRGN = '" & MERGE & "' AND ICOD='" & ICOD & "' AND OPER='+' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   ISSPCS = STKRS!PCS
Else
   ISSPCS = 0
End If

If STKRS.State = 1 Then STKRS.Close
STKRS.Open " SELECT ISNULL(SUM(PCES),0) AS PCS FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND MRGN = '" & MERGE & "' AND ICOD='" & ICOD & "' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   RTIPCS = STKRS!PCS
Else
   RTIPCS = 0
End If

GetItemStockPcsMrgn = ISSPCS - RTIPCS

TXTVBDT.Enabled = False

End Function


Private Function GetItemStockMrgn(ICOD As String) As Double
Dim DIVISIONCODE As String
DIVISIONCODE = "000001"

Dim STKRS As ADODB.Recordset
Set STKRS = New ADODB.Recordset
Dim ISSQTY As Double
Dim RTIQTY As Double

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND  MRGN = '" & Trim(MERGE) & "' AND ICOD='" & ICOD & "' AND OPER='+' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   ISSQTY = STKRS!QNTY
Else
   ISSQTY = 0
End If

If STKRS.State = 1 Then STKRS.Close
STKRS.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVISIONCODE & "' AND  MRGN = '" & Trim(MERGE) & "' AND ICOD='" & ICOD & "' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not STKRS.EOF Then
   RTIQTY = STKRS!QNTY
Else
   RTIQTY = 0
End If

GetItemStockMrgn = ISSQTY - RTIQTY

TXTVBDT.Enabled = False

End Function

Private Sub STOCKCLEAR()

        TXTSTOCK = Empty
        TXTISSQTY = Empty
        MERGE = Empty
        TXTPCS = Empty
        TXTISSPCS = Empty
        txtCops = Empty
        TXTISSCOPS = Empty
        LBLSTKPCS.Enabled = True
        LBLSTKCOPS.Enabled = True
        LBLISSPCS.Enabled = True
        LBLISSCOPS.Enabled = True
        LBLMRGN.Enabled = True
        MERGE.Enabled = True

End Sub

Private Function GetSpeci(ICOD) As String
GetSpeci = ""
Dim SPECI As String
Dim MRGN As String
Dim IGCOD As String
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset
Dim SPRS As ADODB.Recordset
Set SPRS = New ADODB.Recordset
If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT * FROM ITMMST WHERE CODE = '" & ICOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GETRS.EOF Then
   IGCOD = GETRS!igcd
End If

If SPRS.State = 1 Then SPRS.Close
SPRS.Open "SELECT * FROM IGMMST WHERE CODE = '" & Trim(IGCOD) & "'", CN, adOpenDynamic, adLockOptimistic
If Not SPRS.EOF Then
MRGN = SPRS!MERGE
GetSpeci = SPRS!SPECIFICATION
End If
End Function


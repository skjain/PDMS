VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmReturnToStore 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return to Store Division from Another Division"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleMode       =   0  'User
   ScaleWidth      =   10905.17
   Begin FramePlusCtl.FramePlus Frm1 
      Height          =   6255
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   11033
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
         TabIndex        =   15
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox TXTRMRK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   16
         Top             =   5040
         Width           =   5175
      End
      Begin VB.TextBox TXTRATE 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TXTISSQTY 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TXTSTOCK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox TXTICOD 
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
         TabIndex        =   32
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TXTREQSLIP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TXTMACHINE 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox TXTFROMDIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox TXTTODIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox TXTVBNO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   8400
         TabIndex        =   19
         Top             =   1080
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
         Format          =   16449537
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   5520
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
         Image           =   "frmReturnToStore.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   5520
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
         Image           =   "frmReturnToStore.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   5520
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
         Image           =   "frmReturnToStore.frx":1124
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   3960
         TabIndex        =   3
         Top             =   5520
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
         Image           =   "frmReturnToStore.frx":1576
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDOK 
         Height          =   375
         Left            =   9600
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmReturnToStore.frx":19C8
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDITMDEL 
         Height          =   375
         Left            =   9600
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmReturnToStore.frx":1D62
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid ITMFLEX 
         Height          =   1815
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   6
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
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   37
         Top             =   5040
         Width           =   1020
      End
      Begin VB.Label LBLFIFO 
         BackStyle       =   0  'Transparent
         Caption         =   "Note : Edit && Delete are not allowed.             (FIFO Is Applied)"
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
         Left            =   7560
         TabIndex        =   36
         Top             =   5520
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "to Store"
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
         Left            =   6720
         TabIndex        =   35
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "in Division"
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
         Left            =   4920
         TabIndex        =   34
         Top             =   2160
         Width           =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   10800
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks  :"
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
         Left            =   4560
         TabIndex        =   33
         Top             =   5040
         Width           =   930
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   9480
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line9 
         X1              =   9480
         X2              =   9480
         Y1              =   1920
         Y2              =   3000
      End
      Begin VB.Line Line8 
         Visible         =   0   'False
         X1              =   7920
         X2              =   7920
         Y1              =   1920
         Y2              =   3000
      End
      Begin VB.Line Line7 
         X1              =   6240
         X2              =   6240
         Y1              =   1920
         Y2              =   3000
      End
      Begin VB.Line Line6 
         X1              =   4560
         X2              =   4560
         Y1              =   1920
         Y2              =   3000
      End
      Begin VB.Line Line5 
         X1              =   1680
         X2              =   1680
         Y1              =   1920
         Y2              =   3000
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   10800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
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
         Height          =   195
         Left            =   8040
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Return Qnty"
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
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code "
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
         TabIndex        =   29
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1800
         TabIndex        =   28
         Top             =   2160
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
         Left            =   4920
         TabIndex        =   27
         Top             =   1920
         Width           =   930
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   10800
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Slip No. "
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
         Left            =   6960
         TabIndex        =   26
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "From Division "
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
         Left            =   480
         TabIndex        =   25
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Division    "
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
         Left            =   480
         TabIndex        =   24
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return SlipNo. "
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
         Left            =   6960
         TabIndex        =   23
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Date "
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
         Left            =   6960
         TabIndex        =   22
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   6015
         Left            =   120
         Top             =   120
         Width           =   10695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   10800
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Return to Store Division from Another Division"
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
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No."
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
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmReturnToStore"
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
'-------------------------------------------------------------------------------------------
' FORM EVENTS
'-------------------------------------------------------------------------------------------

Private Sub cmdCancel_Click()
  TXTTODIV.Tag = TXTTODIV
  ClsData (Me)
  TXTTODIV = TXTTODIV.Tag
  ITMFLEX.Clear
  ITMFLEX.Rows = 2
  btn_sts (True)
  Call SETFLEX
  cmdAdd.SetFocus
  M_SRNO = Empty
  cmdOk.Caption = "&Add"
  SWITCH = False
End Sub

Private Sub cmdExit_Click()
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
Next J

ITMFLEX.Rows = ITMFLEX.Rows - 1
Call CLEARDATA

If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
           If TXTINAM.Enabled Then TXTINAM.SetFocus
    Else
           If TXTCOST.Enabled = True Then TXTCOST.SetFocus
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
  
'Genrate Sr. No.
 If M_SRNO = Empty Then
    M_SRNO = pubGenSrNoSTR(TXTVBDT, "RTI")
 End If
    
 If SAVEFLAG = True Then
    TXTVBNO = GenVNO("RTI", M_DBCD)
 End If
    
 Call SAVERTI
 
 If SAVEFLAG = True Then
    MsgBox "Your Return Slip No. is " + TXTVBNO.Text
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

Private Sub Form_Activate()
 Call ColorComponent(Me)
 Me.BackColor = RGB(RED, GREEN, BLUE)
 btn_sts (True)
End Sub

Private Sub Form_Load()
 FIFOREQ = "Y"
 Call CenterChild(frm_Main, Me)
 Me.KeyPreview = True
 Me.Tag = zoomflag
 M_DBCD = "000001"
 If Not zoomflag = True Then
    M_SRNO = Empty
 End If
 M_DVCD = "000001"
 TXTVBDT = Date
 TXTVBDT.MaxDate = FEDT
 TXTVBDT.MinDate = FSDT
 Call SETFLEX
 TXTTODIV = GETDIVNAME("000001")
 CMDITMDEL.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If UCase(ActiveControl.NAME) = "TXTRMRK" And KeyAscii = vbKeyReturn Then cmdSave.SetFocus: Exit Sub
 If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
 End If
End Sub
'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
' BUTTON EVENTS
'-------------------------------------------------------------------------------------------
Private Sub cmdAdd_Click()
    zoomflag = False
    btn_sts (False)
    M_SRNO = Empty
    TXTVBNO = GenVNO("RTI", M_DBCD)
    SAVEFLAG = True
    TXTFROMDIV.Enabled = True
    TXTFROMDIV.SetFocus
End Sub

Private Sub cmdOk_Click()
 Dim INDEX As Long
 
 If Not SWITCH Then
      ROWNO = ITMFLEX.Rows - 1
 End If
 
 If CheckData(ROWNO) Then Exit Sub
 
    ITMFLEX.TextMatrix(ROWNO, 0) = Trim(TXTICOD)
    ITMFLEX.TextMatrix(ROWNO, 1) = Trim(TXTINAM)
    ITMFLEX.TextMatrix(ROWNO, 2) = Trim(TXTSTOCK)
    ITMFLEX.TextMatrix(ROWNO, 3) = Trim(nstr(Val(TXTISSQTY), 12, 3))
    ITMFLEX.TextMatrix(ROWNO, 4) = FindFIFORate(GetCode("ITMMST", Trim(TXTINAM), "NAME", "CODE"))
    'ITMFLEX.TextMatrix(ROWNO, 4) = Trim(nstr(Val(TXTRATE), 12, 2))
    ITMFLEX.TextMatrix(ROWNO, 5) = nstr(Val(TXTISSQTY) * Val(ITMFLEX.TextMatrix(ROWNO, 4)), 10, 2)
               
    If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
        TXTINAM.SetFocus
    Else
        TXTCOST.Enabled = True: TXTCOST.SetFocus
    End If
           
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
    Call CLEARDATA
    cmdOk.Caption = "&Add"
    SWITCH = False
End Sub

'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
' LOCAL PROCEDURE
'-------------------------------------------------------------------------------------------
Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    TXTMACHINE.Enabled = Not Yes
    TXTVBDT.Enabled = Not Yes
    TXTICOD.Enabled = Not Yes
    TXTREQSLIP.Enabled = Not Yes
    TXTINAM.Enabled = Not Yes
    TXTISSQTY.Enabled = Not Yes
    'TXTRATE.Enabled = Not Yes
    TXTRMRK.Enabled = Not Yes
    TXTCOST.Enabled = Not Yes
End Sub
'-------------------------------------------------------------------------------------------

Private Sub ITMFLEX_Click()
   If ITMFLEX.Rows > 1 And ITMFLEX.TextMatrix(ITMFLEX.ROW, 1) <> Empty Then
    cmdOk.Caption = "Upd&ate"
    CMDITMDEL.Enabled = True
    ROWNO = ITMFLEX.ROW
    TXTICOD = ITMFLEX.TextMatrix(ROWNO, 0)
    TXTINAM = ITMFLEX.TextMatrix(ROWNO, 1)
    TXTSTOCK = ITMFLEX.TextMatrix(ROWNO, 2)
    TXTISSQTY = ITMFLEX.TextMatrix(ROWNO, 3)
    'TXTRATE = ITMFLEX.TextMatrix(ROWNO, 4)
    SWITCH = True
  End If
    
   If Val(ITMFLEX.ROW) > 0 Then
      If TXTINAM.Enabled Then TXTINAM.SetFocus
   End If
   
End Sub

Private Sub TXTCOST_GotFocus()
 TXTCOST.BackColor = RGB(BRED, BGREEN, BBLUE)
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

Private Sub TXTCOST_LostFocus()
 TXTCOST.BackColor = vbWhite
End Sub

'-------------------------------------------------------------------------------------------
' CODE FOR CURSOR POSITION ON MODULE
'-------------------------------------------------------------------------------------------

Private Sub TXTFROMDIV_GotFocus()
 TXTFROMDIV.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}":
End Sub

Private Sub TXTFROMDIV_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
          
    If KeyCode = vbKeyF2 Or (Trim(TXTFROMDIV) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTFROMDIV.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, TXTFROMDIV.Text, "SELECT DIVISION FROM LIST")
        TXTFROMDIV.Tag = Key
        M_DVNM = TXTFROMDIV
        M_DVCD = Key
    End If
        
    Me.KeyPreview = True
End Sub

Private Sub TXTFROMDIV_LostFocus()
 TXTFROMDIV.BackColor = vbWhite
End Sub

Private Sub TXTINAM_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
If KeyCode = vbKeyF2 Or (Trim(TXTINAM) = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False
   If ItemSearchField = 0 Then
      M_DESC = TXTICOD.Text
      If TXTICOD <> Empty Then TXTSTOCK = GetItemStock(GetDivCode(TXTFROMDIV), TXTICOD)
      TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
   Else
      M_DESC = TXTINAM.Text
   End If
   Key = Empty
   If SAVEFLAG Then
      TXTINAM.Text = SearchITEMLIST("SELECT TOP 20 CODE, NAME FROM ITMMST", 0, TXTINAM.Text, "SELECT ITEM FROM LIST")
   Else
      TXTINAM.Text = SearchITEMLIST("SELECT TOP 20 CODE, NAME FROM ITMMST", 0, TXTINAM.Text, "SELECT ITEM FROM LIST")
   End If
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTINAM.Text = ""
            frm_Item.Show
        Else
            TXTICOD.Text = Key
            TXTSTOCK = GetItemStock(M_DVCD, TXTICOD)
            TXTSTOCK = nstr(Val(TXTSTOCK), 9, 3)
        End If
    Else
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTTODIV_GotFocus()
 TXTTODIV.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTTODIV_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
          
    If KeyCode = vbKeyF2 Or (Trim(TXTTODIV) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTTODIV.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001'", 0, TXTTODIV.Text, "SELECT DIVISION FROM LIST")
        TXTTODIV.Tag = Key
        M_DVNM = TXTTODIV
        M_DVCD = Key
    End If
        
    Me.KeyPreview = True
End Sub

Private Sub TXTTODIV_LostFocus()
 TXTTODIV.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
 TXTMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub txtMACHINE_LostFocus()
 TXTMACHINE.BackColor = vbWhite
End Sub

Private Sub TXTREQSLIP_GotFocus()
 TXTREQSLIP.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTREQSLIP_LostFocus()
 TXTREQSLIP.BackColor = vbWhite
End Sub

Private Sub txtINAM_GotFocus()
 TXTINAM.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub txtINAM_LostFocus()
 TXTINAM.BackColor = vbWhite
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
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTRMRK_LostFocus()
 TXTRMRK.BackColor = vbWhite
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTMACHINE = Empty
    ElseIf KeyCode = vbKeyF2 Then
        M_DESC = Empty
        NEW_VISIBLE = False
        TXTMACHINE = SearchList1("Select  TOP 20 Code,Name From MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & M_DVCD & "'", 0, Empty, "Select M/C FROM MASTER")
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTMACHINE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TXTMACHINE = Empty Then
        Call txtMachine_KeyDown(vbKeyF2, 0)
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub SETFLEX()
  ITMFLEX.Clear
  ITMFLEX.ColWidth(0) = 1440
  ITMFLEX.ColWidth(1) = 3200
  ITMFLEX.ColWidth(2) = 1350
  ITMFLEX.ColWidth(3) = 1350
  ITMFLEX.ColWidth(4) = 1350
  ITMFLEX.ColWidth(5) = 1350
  
  ITMFLEX.Clear
  ITMFLEX.TextMatrix(0, 0) = "Item Code"
  ITMFLEX.TextMatrix(0, 1) = "Item Description"
  ITMFLEX.TextMatrix(0, 2) = "Item Stock"
  ITMFLEX.TextMatrix(0, 3) = "Return Qty"
  ITMFLEX.TextMatrix(0, 4) = "Rate"
  ITMFLEX.TextMatrix(0, 5) = "Amount"
  
  ITMFLEX.ColAlignment(0) = vbLeftJustify
  ITMFLEX.ColAlignment(1) = vbLeftJustify
  ITMFLEX.ColAlignment(2) = vbRightJustify
  ITMFLEX.ColAlignment(3) = vbRightJustify
  ITMFLEX.ColAlignment(4) = vbRightJustify
End Sub

Private Sub CLEARDATA()
        TXTICOD = Empty
        TXTINAM = Empty
        TXTSTOCK = Empty
        TXTISSQTY = Empty
        'TXTRATE = Empty
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
    If Trim(TXTINAM) = Empty Then
        MsgBox "Please Select Items From List !!", vbInformation
        If TXTINAM.Enabled Then TXTINAM.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If Val(TXTISSQTY) < 1 Then
        MsgBox "Please Enter Valid Quantity !!", vbInformation, "Quantity Missing !!"
        If TXTISSQTY.Enabled Then TXTISSQTY.SetFocus
        CheckData = True
        Exit Function
    End If
    
    'If Val(TXTRATE) <= 0 Then
    '    MsgBox "Check Item Rate !!", vbInformation, "Rate Is Missing"
    '    If TXTRATE.Enabled Then TXTRATE.SetFocus
    '    CheckData = True
    '    Exit Function
    'End If
        
    If Val(TXTSTOCK) < Val(TXTISSQTY) Then
        MsgBox "Stock Doesn't Support !!", vbInformation, "Stock Exceed !!"
        If TXTISSQTY.Enabled Then TXTISSQTY.SetFocus
        CheckData = True
        Exit Function
    End If

    For INDEX = 1 To ITMFLEX.Rows - 1
        If Trim(ITMFLEX.TextMatrix(INDEX, 0)) = TXTICOD And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
           MsgBox "Invalid Item Detail"
           If TXTINAM.Enabled Then TXTINAM.SetFocus
           CheckData = True
           Exit Function
        End If
    Next INDEX
    
End Function


Private Function CHKSAVEDATA() As Boolean
If TXTFROMDIV = Empty Then
  MsgBox "Enter Source Division then Save"
  CHKSAVEDATA = True
  If TXTFROMDIV.Enabled Then TXTFROMDIV.SetFocus
  Exit Function
End If

If TXTCOST = Empty Then
  MsgBox "Costing Head is Missing", vbCritical
  CHKSAVEDATA = True
  If TXTCOST.Enabled Then TXTCOST.SetFocus
  Exit Function
End If

If TXTTODIV = Empty Then
  MsgBox "Enter Destination Division then Save"
  CHKSAVEDATA = True
  If TXTTODIV.Enabled Then TXTTODIV.SetFocus
  Exit Function
End If

If TXTMACHINE = Empty Then
  MsgBox "Enter Machine Number then Save"
  CHKSAVEDATA = True
  If TXTMACHINE.Enabled Then TXTMACHINE.SetFocus
  Exit Function
End If

If TXTREQSLIP = Empty Then
  MsgBox "Enter Requision Slip Number !!!", vbInformation
  CHKSAVEDATA = True
  If TXTREQSLIP.Enabled Then TXTREQSLIP.SetFocus
  Exit Function
End If

If ITMFLEX.TextMatrix(1, 0) = Empty Then
  MsgBox "Enter Data then Save"
  CHKSAVEDATA = True
  TXTINAM.Enabled = True
  TXTINAM.SetFocus
  Exit Function
End If

End Function


Private Sub SAVERTI()
  
  On Error GoTo LAST
  Dim SQL As String, CSHD As String
  
  Dim SAVDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  
  Set SAVDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
      
  
  CN.BeginTrans
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM REFMST WHERE NAME='" & TXTCOST & "' AND CATA='N'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
     CSHD = Trim(RS!CODE & "")
  End If
  RS.Close
  
  Call DELETERTI
  SQL = Empty
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND VTYP='RTI' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
  
  Dim AI As String
  Dim BQ As Double
  Dim I As Long
  Dim DVCOD As String
  DVCOD = GetDivCode(TXTFROMDIV)
    
  I = 1
  For I = 1 To ITMFLEX.Rows - 1
    If ITMFLEX.TextMatrix(I, 0) <> Empty Then
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!unit = UNCD
    SAVDAT!DVCD = DVCOD
    SAVDAT!VTYP = "RTI"
    SAVDAT!SRNO = M_SRNO
    SAVDAT!SRCH = I
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = TXTREQSLIP
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!CSHD = CSHD
    SAVDAT!CHEAD = CSHD
    SAVDAT!ICOD = ITMFLEX.TextMatrix(I, 0): AI = ITMFLEX.TextMatrix(I, 0)
    SAVDAT!PCES = 0
    SAVDAT!QNTY = Val(ITMFLEX.TextMatrix(I, 3)): BQ = Val(ITMFLEX.TextMatrix(I, 3))
    SAVDAT!RATE = Val(ITMFLEX.TextMatrix(I, 4))
    SAVDAT!AMNT = Val(ITMFLEX.TextMatrix(I, 5))
    SAVDAT!QORP = "Q"
    SAVDAT![User] = cUName
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    SAVDAT!OPER = "-"
     
    SAVDAT!PCOD = GetMachineCode(DVCOD, TXTMACHINE)
    
    SAVDAT!RECSTAT = "A"
    SAVDAT.Update
        
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "RTI"
    SAVDAT!SRNO = M_SRNO
    SAVDAT!SRCH = I + (ITMFLEX.Rows - 1)
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = TXTREQSLIP
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!CSHD = CSHD
    SAVDAT!CHEAD = CSHD
    SAVDAT!ICOD = ITMFLEX.TextMatrix(I, 0): AI = ITMFLEX.TextMatrix(I, 0)
    SAVDAT!PCES = 0
    SAVDAT!QNTY = Val(ITMFLEX.TextMatrix(I, 3)): BQ = Val(ITMFLEX.TextMatrix(I, 3))
    SAVDAT!RATE = Val(ITMFLEX.TextMatrix(I, 4))
    SAVDAT!AMNT = Val(ITMFLEX.TextMatrix(I, 5))
    SAVDAT!QORP = "Q"
    SAVDAT![User] = cUName
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    SAVDAT!OPER = "+"
    SAVDAT!PCOD = GetMachineCode(DVCOD, TXTMACHINE)
    SAVDAT!DVCD = GetDivCode(TXTTODIV)
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
    SAVDAT.Update
   End If
  Next
  
  'FIFO
    If SAVEFLAG = True And FIFOREQ = "Y" Then
       Call SetFIFOUP
    End If
  '----------------
 
 'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
    Call SetSRNO(TXTVBNO, "RTI", M_DBCD)
  End If
'----------------------------
'DAILYSTATUS ENTRY

 Call DAILYSTATUS("RTI", GetMachineCode(DVCOD, TXTMACHINE), M_DBCD, Val(ITMFLEX.TextMatrix(1, 3)), TXTVBNO, Val(ITMFLEX.TextMatrix(1, 5)), cUName, "N", Now, TXTVBDT)
'---------------------------
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

Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "RTI"
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

Private Sub DELETERTI()
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND DBCD='" & M_DBCD & "' AND VTYP='RTI' AND VBNO='" & TXTVBNO & "' "
End Sub

Private Function GetItemStock(DIVISIONCODE As String, ICOD As String) As Double
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

'FIFO----------------------
Private Sub SetFIFOUP()
On Error GoTo FIFOERR

'VARIABLE DECLARATION
Dim ICOD As String, Item As String, INDEX As Long
Dim BALQNTY As Double, TMPQTY As Double
Dim FIFORS As ADODB.Recordset: Set FIFORS = New ADODB.Recordset

'-------------------------------------------------------------
For INDEX = 1 To ITMFLEX.Rows - 1
'-------------------
'INITIALISE
Item = ITMFLEX.TextMatrix(INDEX, 1)
BALQNTY = Val(ITMFLEX.TextMatrix(INDEX, 3))
'-------------------

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then ICOD = Trim(FIFORS!CODE & "")
FIFORS.Close

'EITHER CASE :IF PENDING GRN EXIST
If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then
       If BALQNTY > 0 Then
          FIFORS!RET_DPT_QNTY = Val(FIFORS!RET_DPT_QNTY) + BALQNTY
          FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
          BALQNTY = 0
       End If
          FIFORS.Update
End If
'----------------------------------------------
'OR CASE : IF NO PENDING GRN EXIST
If BALQNTY > 0 Then
    If FIFORS.State = 1 Then FIFORS.Close
    FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' ORDER BY DATE,VBNO DESC", CN, adOpenDynamic, adLockOptimistic
    If Not FIFORS.EOF Then
       If BALQNTY > 0 Then
          FIFORS!RET_DPT_QNTY = Val(FIFORS!RET_DPT_QNTY) + BALQNTY
          FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
       End If
          FIFORS.Update
    End If
 End If
'-----------------------------

Next INDEX

Exit Sub
FIFOERR:
MsgBox ERR.Description
End Sub


Private Function FindFIFORate(ICOD As String) As Double
'DEFAULT
FindFIFORate = 0
'---------------

Dim F1RS As ADODB.Recordset
Set F1RS = New ADODB.Recordset

If F1RS.State = 1 Then F1RS.Close
F1RS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
If Not F1RS.EOF Then
   FindFIFORate = Val(F1RS!RATE)
   F1RS.Close
   Exit Function
Else 'SPECIAL CASE : : IF NO PENDING GRN EXIST

    If F1RS.State = 1 Then F1RS.Close
    F1RS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND ICOD='" & ICOD & "' ORDER BY DATE,VBNO DESC", CN, adOpenDynamic, adLockOptimistic
    If Not F1RS.EOF Then
       FindFIFORate = Val(F1RS!RATE)
       F1RS.Close
    End If
    
End If

End Function

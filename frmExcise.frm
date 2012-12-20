VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmExcise 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXCISE ENTRY"
   ClientHeight    =   7995
   ClientLeft      =   2985
   ClientTop       =   2025
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8065.909
   ScaleMode       =   0  'User
   ScaleWidth      =   10322.32
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   6960
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   7995
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   14102
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
      Begin VB.TextBox txtcess 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   5640
         Width           =   2055
      End
      Begin VB.TextBox txtvbno 
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   43
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtpurac 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4320
         Width           =   6495
      End
      Begin VB.TextBox TXTEXTRA5 
         Height          =   1575
         Left            =   4320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   5520
         Width           =   4215
      End
      Begin VB.TextBox CENVAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox EDUCESS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   6000
         Width           =   2055
      End
      Begin VB.TextBox HEDUCESS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   6360
         Width           =   2055
      End
      Begin VB.TextBox ADUTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   6720
         Width           =   2055
      End
      Begin VB.TextBox SUPNAM 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2760
         Width           =   6495
      End
      Begin VB.TextBox MFGNAM 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3120
         Width           =   6495
      End
      Begin VB.TextBox ITMDESC 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3480
         Width           =   6495
      End
      Begin VB.TextBox TQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox ITOT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6600
         TabIndex        =   16
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox EXCREG 
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
         ItemData        =   "frmExcise.frx":0000
         Left            =   2040
         List            =   "frmExcise.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox SUPTYPE 
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
         ItemData        =   "frmExcise.frx":002E
         Left            =   2040
         List            =   "frmExcise.frx":003B
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox VBNO 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6120
         TabIndex        =   4
         Top             =   8280
         Visible         =   0   'False
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
         Image           =   "frmExcise.frx":0069
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   3720
         TabIndex        =   1
         Top             =   7320
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
         Image           =   "frmExcise.frx":0403
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4920
         TabIndex        =   2
         Top             =   7320
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
         Image           =   "frmExcise.frx":118D
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   7320
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
         Image           =   "frmExcise.frx":15DF
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker ENTDAT 
         Height          =   315
         Left            =   7320
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   18284545
         CurrentDate     =   39347
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   2520
         TabIndex        =   0
         Top             =   7320
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
         Image           =   "frmExcise.frx":1A31
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   7320
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
         Image           =   "frmExcise.frx":1DCB
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker txtodat 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   18284545
         CurrentDate     =   39347
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Cess"
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
         Left            =   480
         TabIndex        =   45
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   240
         X2              =   8760
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Journal V. No."
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
         Left            =   5520
         TabIndex        =   44
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase A/c"
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
         Left            =   480
         TabIndex        =   42
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   8295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Narration (If Any)"
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
         Left            =   5520
         TabIndex        =   41
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Height          =   2415
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   8535
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Cenvat"
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
         Left            =   480
         TabIndex        =   40
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Edu. Cess"
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
         Left            =   480
         TabIndex        =   39
         Top             =   6000
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Hr. Edu Cess"
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
         Left            =   480
         TabIndex        =   38
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Add. Duty"
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
         Left            =   480
         TabIndex        =   37
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Supplier"
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
         Left            =   480
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Mfg."
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
         Left            =   480
         TabIndex        =   35
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Quantity"
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
         Left            =   480
         TabIndex        =   33
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Assessable Value"
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
         Left            =   4800
         TabIndex        =   32
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Excise Register"
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
         Left            =   480
         TabIndex        =   31
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Type"
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
         Left            =   480
         TabIndex        =   30
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Date"
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
         TabIndex        =   29
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document No. "
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
         TabIndex        =   28
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Date "
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
         Left            =   6000
         TabIndex        =   27
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label LBLCHHEAD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No. "
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
         Left            =   6000
         TabIndex        =   26
         Top             =   720
         Width           =   1005
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   540
         Left            =   5040
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "EXCISE CREDIT ENTRY "
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
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label ENTRYNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXX"
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
         Left            =   7320
         TabIndex        =   24
         Top             =   750
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmExcise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVFLAG As Boolean
Public M_VBNO As String

Private Sub ADUTY_GotFocus()
ADUTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ADUTY_LostFocus()
ADUTY.BackColor = vbWhite
End Sub

Private Sub CENVAT_GotFocus()
CENVAT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub CENVAT_LostFocus()
CENVAT.BackColor = vbWhite
End Sub

Private Sub CENVAT_Validate(Cancel As Boolean)
  If Not IsNumeric(CENVAT) Then
    MsgBox "Invalid Cenvate Amount"
    Cancel = True
  End If
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo LAST
  If Allow_view_only = "Y" Then
     cmdCancel_Click
     btn_sts (True)
     Exit Sub
  End If
  SAVFLAG = True
  btn_sts False
  
  ENTRYNO.Caption = GenVNO("EXC", "000001")
  Call genJOURNAL
  VBNO.SetFocus
  EXCREG.ListIndex = 0
  SUPTYPE.ListIndex = 0
  ENTDAT.MinDate = FSDT
  ENTDAT.MaxDate = FEDT
  txtodat.MinDate = FSDT
  txtodat.MaxDate = FEDT
  Exit Sub
LAST:
  MsgBox ERR.Description
End Sub

Private Sub cmdCancel_Click()
  Call ClsData(Me)
  btn_sts True
  If cmdAdd.Enabled = True Then
    cmdAdd.SetFocus
  End If
  ENTRYNO.Caption = GenVNO("EXC", "000001")
  SAVFLAG = True
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000074", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  SAVFLAG = False
  M_VBNO = Empty
  frm_EXCISECRList.Show 1
  If Allow_view_only = "Y" Then
     cmdCancel_Click
     btn_sts (True)
     Exit Sub
  End If
  If M_VBNO <> Empty Then
     Dim AYS
     AYS = MsgBox("Are You Sure To Delete the Entry ? ", vbYesNo)
     If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE EGPMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND SRNO='" & M_VBNO & "' AND VTYP='EXC'"
        CN.CommitTrans
        btn_sts (False)
    End If
  End If
  
  btn_sts (True)
  cmdAdd.SetFocus
  Call cmdCancel_Click
  Exit Sub
End Sub

Private Sub cmdEdit1_Click()
  
  Dim EDIT_ENTRYNO As String
  EDIT_ENTRYNO = InputBox("Enter Entry No to Edit Excise Entry", "Excise Entry Editing")
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND SRNO='" & EDIT_ENTRYNO & "' AND VTYP='EXC'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Entry Does Not Exist"
    Call cmdCancel_Click
    Exit Sub
  End If
  SAVFLAG = False
  btn_sts False
  'Store All The Data In Respective Box
  Select Case RS!TTYP
   Case "RG23-A"
    EXCREG.ListIndex = 0
   Case "RG23-C"
    EXCREG.ListIndex = 1
   Case Else
    EXCREG.ListIndex = 2
  End Select
  Select Case RS!EXTRA1 & ""
   Case "Manufacturer"
    SUPTYPE.ListIndex = 0
   Case "1st Stage Dealer"
    SUPTYPE.ListIndex = 1
   Case Else
    SUPTYPE.ListIndex = 2
  End Select
  ENTRYNO = Trim(RS!SRNO)
  ENTDAT.Value = RS!Date
  TQTY.Text = RS!QNTY
  VBNO.Text = Trim(RS!VBNO & "")
  ITOT.Text = RS!ITOT
  CENVAT.Text = RS!CENVAT
  EDUCESS.Text = RS!EDUCESS
  HEDUCESS.Text = RS!H_ED_CESS
  ADUTY.Text = RS!A_DUTY
  TXTEXTRA5.Text = RS!EXTRA5 & ""
  If IsNull(RS!CHDT) Then
    txtodat = RS!Date
   Else
    txtodat = RS!CHDT
  End If
  Dim SUP_COD As String
  Dim MFG_COD As String
  Dim ITM_COD As String
  SUP_COD = RS!CRAC
  MFG_COD = RS!DRAC
  ITM_COD = RS!ICOD
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & SUP_COD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    SUPNAM = RS!NAME & ""
   Else
    SUPNAM = Empty
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & MFG_COD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    MFGNAM = RS!NAME & ""
   Else
    MFGNAM = Empty
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM ITMMST WHERE CODE='" & ITM_COD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    ITMDESC = RS!NAME & ""
   Else
    ITMDESC = Empty
  End If
  EXCREG.SetFocus
End Sub

Private Sub cmdEdit_Click()
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000074", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
  SAVFLAG = False
  M_VBNO = Empty
  frm_EXCISECRList.Show 1
  If ALLOW_EDIT(compPth, UNCD, "JOU", "000001", Trim(txtvbno)) = "C" Then
      cmdCancel_Click
      btn_sts (True)
      Exit Sub
  End If
  If Allow_view_only = "Y" Then
     cmdCancel_Click
     btn_sts (True)
     Exit Sub
  End If
  If M_VBNO <> Empty Then
     btn_sts (False)
     EXCREG.Enabled = True
     EXCREG.SetFocus
  Else
     btn_sts (True)
     cmdAdd.SetFocus
     Call cmdCancel_Click
     Exit Sub
  End If
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub EDUCESS_GotFocus()
EDUCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub EDUCESS_LostFocus()
EDUCESS.BackColor = vbWhite
End Sub

Private Sub EDUCESS_Validate(Cancel As Boolean)
  If Not IsNumeric(EDUCESS) Then
    MsgBox "Invalid Edu. Cess Amount"
    Cancel = True
  End If
End Sub

Private Sub ENTDAT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub EXCREG_GotFocus()
EXCREG.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub EXCREG_LostFocus()
EXCREG.BackColor = vbWhite
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If ActiveControl.NAME = "SUPNAM" Or ActiveControl.NAME = "MFGNAM" Then Exit Sub
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
On Error GoTo LOADFORMERR
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  EXCREG.ListIndex = 0
  SUPTYPE.ListIndex = 0
  ENTDAT.Value = Now
  txtodat.Value = Now
  btn_sts True
  Me.KeyPreview = True
  If UNT_ISPAPPER = "Y" Then
    Label18.Visible = True
    txtcess.Visible = True
   Else
    Label18.Visible = False
    txtcess.Visible = False
  End If
  Exit Sub
LOADFORMERR:
  If ERR.Description = "Only one MDI Form allowed" Then
  MsgBox "System Virtual Memory Full"
  End If
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
End Sub

Private Sub HEDUCESS_GotFocus()
HEDUCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub HEDUCESS_LostFocus()
HEDUCESS.BackColor = vbWhite
End Sub

Private Sub HEDUCESS_Validate(Cancel As Boolean)
  If Not IsNumeric(HEDUCESS) Then
    MsgBox "Invalid Hr. Edu. Cess Amount"
    Cancel = True
  End If
End Sub

Private Sub ITMDESC_GotFocus()
ITMDESC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ITMDESC_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
  If ITMDESC = Empty Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    ITMDESC = SearchList1("SELECT TOP 20 CODE,NAME FROM ITMMST", 0, ITMDESC, "SELECT ITEM NAME FROM LIST")
  End If
  If KeyCode = vbKeyReturn And ITMDESC <> Empty Then
    txtpurac.SetFocus
  End If
End Sub

Private Sub ITMDESC_LostFocus()
ITMDESC.BackColor = vbWhite
End Sub

Private Sub ITOT_GotFocus()
ITOT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ITOT_LostFocus()
ITOT.BackColor = vbWhite
End Sub

Private Sub ITOT_Validate(Cancel As Boolean)
  If Not IsNumeric(ITOT) Then
    MsgBox "Invalid Ass.Value"
    Cancel = True
  End If
End Sub

Private Sub MFGNAM_GotFocus()
MFGNAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub MFGNAM_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
  If MFGNAM = Empty Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    MFGNAM = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, MFGNAM, "SELECT MFG. NAME FROM LIST")
  End If
  If KeyCode = vbKeyReturn And MFGNAM <> Empty Then
    ITMDESC.SetFocus
  End If
End Sub

Private Sub MFGNAM_LostFocus()
 MFGNAM.BackColor = vbWhite
End Sub

Private Sub SUPNAM_GotFocus()
SUPNAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SUPNAM_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
  If SUPNAM = Empty Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    SUPNAM = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, SUPNAM, "SELECT SUPPLIER NAME FROM LIST")
  End If
  If KeyCode = vbKeyReturn And SUPNAM <> Empty Then
    MFGNAM.SetFocus
  End If
End Sub

Private Sub SUPNAM_LostFocus()
SUPNAM.BackColor = vbWhite
End Sub

Private Sub SUPTYPE_GotFocus()
SUPTYPE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SUPTYPE_LostFocus()
SUPTYPE.BackColor = vbWhite
End Sub

Private Sub TimerBillNo1_Timer()
    Static ctr As Integer
    If ctr Mod 45 = 0 And ctr <= 45 Then
       lblAlert.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
       BORDER.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
       ENTRYNO.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
       LBLCHHEAD.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
    ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
       lblAlert.ForeColor = vbRed
       BORDER.BorderColor = vbRed
       ENTRYNO.ForeColor = vbRed
       LBLCHHEAD.ForeColor = vbRed
    ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
       lblAlert.ForeColor = vbBlue
       BORDER.BorderColor = vbBlue
       ENTRYNO.ForeColor = vbBlue
       LBLCHHEAD.ForeColor = vbBlue
       ctr = 0
    End If
    ctr = ctr + 15
End Sub


Private Sub TQTY_GotFocus()
  TQTY.BackColor = RGB(BRED, BGREEN, BBLUE)
  Me.KeyPreview = True
End Sub

Private Sub TQTY_LostFocus()
TQTY.BackColor = vbWhite
End Sub

Private Sub TQTY_Validate(Cancel As Boolean)
  Me.KeyPreview = True
  If Not IsNumeric(TQTY) Then
    MsgBox "Invalid Quantity"
    Cancel = True
  End If
End Sub

Private Sub cmdSave_Click()
  
  On Error GoTo LAST
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ACCMST WHERE NAME='" & txtpurac & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Purchase A/c"
    txtpurac.SetFocus
    Exit Sub
  End If
  txtpurac.Tag = RS!CODE
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM EXCISEOPENING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Excise A/c"
    VBNO.SetFocus
    Exit Sub
  End If
  
  Dim CENVATAC As String
  Dim PCESSAC As String
  Dim ECESSAC As String
  Dim HEDCSAC As String
  Dim DIFFEDAC As String
  Dim PERCAC As Double
  Dim SQL As String
  SQL = Empty
  Select Case EXCREG.Text
   Case "RG23-A"
    SQL = "SELECT RG23ACENVAT AS CENVATAC,RG23ACESSAC AS PCESSAC,RG23AEDUCESS AS ECESSAC,RG23AHEDCESS AS HEDCSAC, '' AS DEFFEDAC FROM EXCISEOPENING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'"
   Case "RG23-C"
    SQL = "SELECT RG23CCENVAT AS CENVATAC,'' AS PCESSAC,RG23CEDUCESS AS ECESSAC,RG23CHEDCESS AS HEDCSAC,rg23cdeffered AS DEFFEDAC FROM EXCISEOPENING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'"
   Case "SERVICE TAX"
    SQL = "SELECT SRVCENVAT AS CENVATAC,'' AS PCESSAC,SRVAEDUCESS AS ECESSAC,SRVHEDCESS AS HEDCSAC, '' AS DEFFEDAC FROM EXCISEOPENING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'"
   Case Else
    MsgBox "Invalid Excise Type"
    EXCREG.SetFocus
    Exit Sub
  End Select
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT EXCCPERC FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    PERCAC = RS!exccperc
  End If
  If RS.State = 1 Then RS.Close
  RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Excise Detail"
    VBNO.SetFocus
    Exit Sub
  End If
  
  
  CENVATAC = RS!CENVATAC & ""
  PCESSAC = RS!PCESSAC & ""
  ECESSAC = RS!ECESSAC & ""
  HEDCSAC = RS!HEDCSAC & ""
  DIFFEDAC = RS!DEFFEDAC & ""
  
  If SAVFLAG = True Then
     ENTRYNO.Caption = GenVNO("EXC", "000001")
  End If
  
  TXTEXTRA5 = Replace(Trim(TXTEXTRA5), vbCrLf, "")
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT CODE FROM ACCMST WHERE NAME='" & txtpurac & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Purchase A/c"
    txtpurac.SetFocus
    Exit Sub
  End If
  txtpurac.Tag = RS!CODE & ""
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT SRNO,VBNO FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT<>'D' AND VBNO='" & VBNO & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    If Trim(RS!SRNO) = Trim(ENTRYNO) Then
      'O.k
     Else
      MsgBox "Duplicate ARE-1"
      VBNO.SetFocus
      Exit Sub
    End If
  End If
  
  Dim SAVDAT As New ADODB.Recordset
  Dim SUP_COD As String
  Dim MFG_COD As String
  Dim ITM_COD As String
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT CODE FROM ACCMST WHERE NAME='" & SUPNAM & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
    MsgBox "Invalid Supplier Name"
    SUPNAM.SetFocus
    Exit Sub
  End If
  SUP_COD = SAVDAT!CODE
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT CODE FROM ACCMST WHERE NAME='" & MFGNAM & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
    MsgBox "Invalid Mfg. Name"
    MFGNAM.SetFocus
    Exit Sub
  End If
  MFG_COD = SAVDAT!CODE
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT CODE FROM ITMMST WHERE NAME='" & ITMDESC & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
    MsgBox "Invalid Item Name"
    ITMDESC.SetFocus
    Exit Sub
  End If
  ITM_COD = SAVDAT!CODE
  
  If EXCREG = Empty Then
    MsgBox "Invalid Excise Register"
    EXCREG.SetFocus
    Exit Sub
  End If
  
  If SUPTYPE = Empty Then
    MsgBox "Invalid Supplier Type"
    SUPTYPE.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(TQTY) Then
    MsgBox "Invalid Quantity"
    TQTY.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(ITOT) Then
    MsgBox "Invalid Ass.Value"
    ITOT.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(CENVAT) Then
    MsgBox "Invalid Cenvat Amount"
    CENVAT.SetFocus
    Exit Sub
  End If
  If UNT_ISPAPPER = "N" Then
    txtcess = 0
  End If
  If Not IsNumeric(txtcess) Then
    MsgBox "Invalid P.CESS Amount"
    txtcess.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(ADUTY) Then
    MsgBox "Invalid Add. Duty Amount"
    ADUTY.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(EDUCESS) Then
    MsgBox "Invalid Edu. Cess Amount"
    EDUCESS.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(HEDUCESS) Then
    MsgBox "Invalid Hr. Edu. Cess Amount"
    HEDUCESS.SetFocus
    Exit Sub
  End If
  Dim C_CENVAT As Double
  Dim C_EDCESS As Double
  Dim C_HEDCES As Double
  Dim C_PCESS As Double
  C_CENVAT = 0
  C_EDCESS = 0
  C_HEDCES = 0
  C_PCESS = 0
  If EXCREG.Text = "RG23-C" Then
    C_CENVAT = (Val(CENVAT.Text) * PERCAC) / 100
    C_PCESS = (Val(txtcess.Text) * PERCAC) / 100
    C_EDCESS = (Val(EDUCESS.Text) * PERCAC) / 100
    C_HEDCES = (Val(HEDUCESS.Text) * PERCAC) / 100
    
  End If
  'Delete Old Entries
  CN.BeginTrans
  If SAVFLAG = True Then
    Call genJOURNAL
  End If
  If Trim(txtvbno) = "" Then
    Call genJOURNAL
  End If
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' and unit='" & UNCD & "' AND VTYP='EXC' AND SRNO='" & ENTRYNO & "'"
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND VTYP='EXC' AND unit='" & UNCD & "' and SRNO='" & ENTRYNO & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
    SAVDAT.AddNew
  End If
  
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "EXC"
  SAVDAT!SRNO = ENTRYNO
  SAVDAT!SRCH = "1"
  SAVDAT!Date = Format(ENTDAT, "YYYY/MM/DD")
  SAVDAT!dbcd = "EXCREG"
  SAVDAT!CRAC = SUP_COD
  SAVDAT!DRAC = MFG_COD
  SAVDAT!VBNO = Trim(VBNO)
  SAVDAT!chln = Trim(VBNO)
  SAVDAT!ICOD = ITM_COD
  SAVDAT!QNTY = Val(TQTY)
  SAVDAT!AMNT = Val(ITOT)
  SAVDAT!ITOT = Val(ITOT)
  SAVDAT!TTYP = EXCREG.Text
  SAVDAT!RECSTAT = "A"
  SAVDAT!unit = UNCD
  'If EXCREG.Text = "RG23-C" Then
  '  SAVDAT!CENVAT = Val(C_CENVAT)
  '  SAVDAT!CESS = Val(C_PCESS)
  '  SAVDAT!EDUCESS = Val(C_EDCESS)
  '  SAVDAT!H_ED_CESS = Val(C_HEDCES)
  '  SAVDAT!A_DUTY = Val(ADUTY)
  'Else
    SAVDAT!CENVAT = Val(CENVAT) - Val(C_CENVAT)
    SAVDAT!CESS = Val(txtcess) - Val(C_PCESS)
    SAVDAT!EDUCESS = Val(EDUCESS) - Val(C_EDCESS)
    SAVDAT!H_ED_CESS = Val(HEDUCESS) - Val(C_HEDCES)
    SAVDAT!A_DUTY = Val(ADUTY)
  'End If
  
  SAVDAT!EXTRA4 = SUPTYPE.Text
  SAVDAT!EXTRA3 = "True"
  SAVDAT!CHDT = Format(txtodat.Value, "YYYY/MM/DD")
  SAVDAT!EXTRA5 = Mid(Trim(TXTEXTRA5.Text), 1, 50)
  SAVDAT!EXTRA1 = txtpurac.Tag
  SAVDAT!EXTRA2 = txtvbno
  SAVDAT.Update
  CN.CommitTrans
  Dim TOTEXC As Double
  TOTEXC = Val(CENVAT) + Val(EDUCESS) + Val(HEDUCESS) + Val(ADUTY) + Val(txtcess)
  
  'Save Record in trnman
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=1", CN, adOpenForwardOnly, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "JOU"
  RS!SRNO = "0"
  RS!SRCH = "1"
  RS!unit = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(ENTDAT, "YYYY/MM/DD")
  RS!dbcd = "000001"
  RS!vno = txtvbno
  RS!ACOD = txtpurac.Tag
  RS!RCOD = ""
  RS!damt = 0
  RS!camt = TOTEXC
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = TOTEXC
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXC"
  RS!EXTRA2 = ENTRYNO
  RS.Update
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=2", CN, adOpenForwardOnly, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "JOU"
  RS!SRNO = "0"
  RS!SRCH = "2"
  RS!unit = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(ENTDAT, "YYYY/MM/DD")
  RS!dbcd = "000001"
  RS!vno = txtvbno
  RS!ACOD = CENVATAC
  RS!RCOD = txtpurac.Tag
  RS!damt = Val(CENVAT.Text) + Val(ADUTY)
  RS!camt = 0
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(CENVAT.Text) + Val(ADUTY)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXC"
  RS!EXTRA2 = ENTRYNO
  RS.Update
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=3", CN, adOpenForwardOnly, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "JOU"
  RS!SRNO = "0"
  RS!SRCH = "3"
  RS!unit = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(ENTDAT, "YYYY/MM/DD")
  RS!dbcd = "000001"
  RS!vno = txtvbno
  RS!ACOD = PCESSAC
  RS!RCOD = txtpurac.Tag
  RS!damt = Val(txtcess.Text)
  RS!camt = 0
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(CENVAT.Text) + Val(ADUTY)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXC"
  RS!EXTRA2 = ENTRYNO
  RS.Update
  
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=4", CN, adOpenForwardOnly, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "JOU"
  RS!SRNO = "0"
  RS!SRCH = "4"
  RS!unit = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(ENTDAT, "YYYY/MM/DD")
  RS!dbcd = "000001"
  RS!vno = txtvbno
  RS!ACOD = ECESSAC
  RS!RCOD = txtpurac.Tag
  RS!damt = Val(EDUCESS.Text)
  RS!camt = 0
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(EDUCESS.Text)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXC"
  RS!EXTRA2 = ENTRYNO
  RS.Update
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=5", CN, adOpenForwardOnly, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!VTYP = "JOU"
  RS!SRNO = "0"
  RS!SRCH = "5"
  RS!unit = UNCD
  RS!MSDVCD = "000001"
  RS!User = cUName
  RS!Date = Format(ENTDAT, "YYYY/MM/DD")
  RS!dbcd = "000001"
  RS!vno = txtvbno
  RS!ACOD = HEDCSAC
  RS!RCOD = txtpurac.Tag
  RS!damt = Val(HEDUCESS.Text)
  RS!camt = 0
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(HEDUCESS.Text)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXC"
  RS!EXTRA2 = ENTRYNO
  RS.Update
  
  CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU'  AND VNO='" & txtvbno & "' AND SRCH=6"
  CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU'  AND VNO='" & txtvbno & "' AND SRCH=7"
  CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU'  AND VNO='" & txtvbno & "' AND SRCH=8"
  CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU'  AND VNO='" & txtvbno & "' AND SRCH=9"
  
  If EXCREG.Text = "RG23-C" And C_CENVAT > 0 Then
    'For Cenvat Reserve
    Dim C_TOTEXC As Double
    C_TOTEXC = C_CENVAT + C_EDCESS + C_HEDCES
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=6", CN, adOpenForwardOnly, adLockOptimistic
    If RS.EOF Then
      RS.AddNew
    End If
    RS!COMP = compPth
    RS!VTYP = "JOU"
    RS!SRNO = "0"
    RS!SRCH = "6"
    RS!unit = UNCD
    RS!MSDVCD = "000001"
    RS!User = cUName
    RS!Date = Format(ENTDAT, "YYYY/MM/DD")
    RS!dbcd = "000001"
    RS!vno = txtvbno
    RS!ACOD = DIFFEDAC
    RS!RCOD = ""
    RS!damt = C_TOTEXC
    RS!camt = 0
    RS!VBNO = Trim(txtvbno)
    RS!narr = TXTEXTRA5
    RS!AMNT = C_TOTEXC
    RS!RECSTAT = "A"
    RS!MLTENT = "Y"
    RS!EXTRA1 = "EXC"
    RS!EXTRA2 = ENTRYNO
    RS.Update
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=7", CN, adOpenForwardOnly, adLockOptimistic
    If RS.EOF Then
      RS.AddNew
    End If
    RS!COMP = compPth
    RS!VTYP = "JOU"
    RS!SRNO = "0"
    RS!SRCH = "7"
    RS!unit = UNCD
    RS!MSDVCD = "000001"
    RS!User = cUName
    RS!Date = Format(ENTDAT, "YYYY/MM/DD")
    RS!dbcd = "000001"
    RS!vno = txtvbno
    RS!ACOD = CENVATAC
    RS!RCOD = DIFFEDAC
    RS!damt = 0
    RS!camt = C_CENVAT
    RS!VBNO = Trim(txtvbno)
    RS!narr = TXTEXTRA5
    RS!AMNT = C_CENVAT
    RS!RECSTAT = "A"
    RS!MLTENT = "Y"
    RS!EXTRA1 = "EXC"
    RS!EXTRA2 = ENTRYNO
    RS.Update
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=8", CN, adOpenForwardOnly, adLockOptimistic
    If RS.EOF Then
      RS.AddNew
    End If
    RS!COMP = compPth
    RS!VTYP = "JOU"
    RS!SRNO = "0"
    RS!SRCH = "8"
    RS!unit = UNCD
    RS!MSDVCD = "000001"
    RS!User = cUName
    RS!Date = Format(ENTDAT, "YYYY/MM/DD")
    RS!dbcd = "000001"
    RS!vno = txtvbno
    RS!ACOD = ECESSAC
    RS!RCOD = DIFFEDAC
    RS!damt = 0
    RS!camt = C_EDCESS
    RS!VBNO = Trim(txtvbno)
    RS!narr = TXTEXTRA5
    RS!AMNT = C_EDCESS
    RS!RECSTAT = "A"
    RS!MLTENT = "Y"
    RS!EXTRA1 = "EXC"
    RS!EXTRA2 = ENTRYNO
    RS.Update
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='000001' AND VTYP='JOU' AND VNO='" & txtvbno & "' AND SRCH=9", CN, adOpenForwardOnly, adLockOptimistic
    If RS.EOF Then
      RS.AddNew
    End If
    RS!COMP = compPth
    RS!VTYP = "JOU"
    RS!SRNO = "0"
    RS!SRCH = "9"
    RS!unit = UNCD
    RS!MSDVCD = "000001"
    RS!User = cUName
    RS!Date = Format(ENTDAT, "YYYY/MM/DD")
    RS!dbcd = "000001"
    RS!vno = txtvbno
    RS!ACOD = HEDCSAC
    RS!RCOD = DIFFEDAC
    RS!damt = 0
    RS!camt = C_HEDCES
    RS!VBNO = Trim(txtvbno)
    RS!narr = TXTEXTRA5
    RS!AMNT = C_HEDCES
    RS!RECSTAT = "A"
    RS!MLTENT = "Y"
    RS!EXTRA1 = "EXC"
    RS!EXTRA2 = ENTRYNO
    RS.Update
    
    
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND VTYP='EXC' AND unit='" & UNCD & "' and SRNO='" & ENTRYNO & "' AND SRCH=2", CN, adOpenDynamic, adLockOptimistic
    If SAVDAT.EOF Then
      SAVDAT.AddNew
    End If
  
    Dim NEWFSDT As Date
    NEWFSDT = ("01/04/" + Trim(STR(Year(FSDT) + 1)))
    
  
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "EXC"
    SAVDAT!SRNO = ENTRYNO
    SAVDAT!SRCH = "2"
    SAVDAT!Date = Format(NEWFSDT, "YYYY/MM/DD")
    SAVDAT!dbcd = "EXCRG1"
    SAVDAT!CRAC = SUP_COD
    SAVDAT!DRAC = MFG_COD
    SAVDAT!VBNO = Trim(VBNO)
    SAVDAT!chln = Trim(VBNO)
    SAVDAT!ICOD = ITM_COD
    SAVDAT!QNTY = Val(TQTY)
    SAVDAT!AMNT = Val(ITOT)
    SAVDAT!ITOT = Val(ITOT)
    SAVDAT!TTYP = EXCREG.Text
    SAVDAT!RECSTAT = "A"
    SAVDAT!unit = UNCD
    SAVDAT!CENVAT = Val(C_CENVAT)
    SAVDAT!CESS = Val(C_PCESS)
    SAVDAT!EDUCESS = Val(C_EDCESS)
    SAVDAT!H_ED_CESS = Val(C_HEDCES)
    SAVDAT!A_DUTY = Val(ADUTY)
    
    SAVDAT!EXTRA4 = SUPTYPE.Text
    SAVDAT!EXTRA3 = "True"
    SAVDAT!CHDT = Format(txtodat.Value, "YYYY/MM/DD")
    SAVDAT!EXTRA5 = Mid(Trim(TXTEXTRA5.Text), 1, 50)
    SAVDAT!EXTRA1 = txtpurac.Tag
    SAVDAT!EXTRA2 = txtvbno
    SAVDAT.Update
    
  End If
  
  'Save New Record
  'Update excno in untmst
  If SAVFLAG = True Then
    
    CN.Execute "UPDATE SERIALMASTER SET [SRNO]='" & ENTRYNO & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
          "' AND VTYP='EXC' AND CODE='000001' AND FYCD='" & FYCD & "'"
          
    CN.Execute "UPDATE SERIALMASTER SET SRNO='" & Mid(txtvbno, 1, 6) & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND FYCD='" & FYCD & "' AND VTYP='JOU' AND CODE='000001'"
  
  End If
  
  Call UPDATESTATUS
  Call ClsData(Me)
  Call cmdCancel_Click
  btn_sts (True)
  cmdAdd.SetFocus
  Exit Sub
LAST:
  MsgBox ERR.Description
  Resume
End Sub

Private Sub TXTEXTRA5_GotFocus()
 TXTEXTRA5.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTEXTRA5_LostFocus()
TXTEXTRA5.BackColor = vbWhite
End Sub

Private Sub txtodat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub VBNO_GotFocus()
VBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub VBNO_LostFocus()
 VBNO.BackColor = vbWhite
End Sub

Private Sub VBNO_Validate(Cancel As Boolean)
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT SRNO,VBNO FROM EGPMAN WHERE VBNO='" & VBNO & "' AND  COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF = False Then
    If Trim(RS!SRNO) = Trim(ENTRYNO) Then
      'O.k
     Else
      MsgBox "Duplicate ARE-1"
      Cancel = True
    End If
  End If
End Sub

Private Sub TXTPURAC_GotFocus()
txtpurac.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPURAC_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
  If txtpurac = Empty Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    txtpurac = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, txtpurac, "SELECT PURCHASE A/C FROM LIST")
  End If
  If KeyCode = vbKeyReturn And txtpurac <> Empty Then
    TQTY.SetFocus
  End If
End Sub

Private Sub TXTPURAC_LostFocus()
txtpurac.BackColor = vbWhite
End Sub
Private Sub genJOURNAL()
    Dim TEMPRS As New ADODB.Recordset
    Dim ctr As Double
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "select SRNO from SERIALMASTER where comp='" & compPth & "' AND UNIT='" & UNCD & "' AND FYCD='" & FYCD & "' AND VTYP='JOU'", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS![SRNO] = "" Or IsNull(TEMPRS![SRNO]) = True Then
        txtvbno = "000001" + Mid(FYCD, 1, 4)
        Exit Sub
    End If
    ctr = Val(Trim(Mid(TEMPRS![SRNO], 1, 6)))
    ctr = ctr + 1
    If ctr < 10 Then
        txtvbno = "00000" + CStr(ctr)
    ElseIf ctr >= 10 And ctr < 100 Then
        txtvbno = "0000" + CStr(ctr)
    ElseIf ctr >= 100 And ctr < 1000 Then
        txtvbno = "000" + CStr(ctr)
    ElseIf ctr >= 1000 And ctr < 10000 Then
        txtvbno = "00" + CStr(ctr)
    ElseIf ctr >= 100 And ctr < 100000 Then
        txtvbno = "0" + CStr(ctr)
    Else
        txtvbno = CStr(ctr)
    End If
    txtvbno = txtvbno + Mid(FYCD, 1, 4)
End Sub

Private Sub TXTCESS_GotFocus()
txtcess.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCESS_LostFocus()
txtcess.BackColor = vbWhite
End Sub

Private Sub TXTCESS_Validate(Cancel As Boolean)
  If Not IsNumeric(txtcess) Then
    MsgBox "Invalid Paper Cess Amount"
    Cancel = True
  End If
End Sub


Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM FASDAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "EXC"
  DLYSTA!PCOD = "Direct Cr Entry"
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = 0
  DLYSTA!VBNO = ENTRYNO
  DLYSTA!AMNT = Val(CENVAT) + Val(EDUCESS) + Val(HEDUCESS)
  DLYSTA!CUSR = cUName
  If SAVFLAG = True Then
    DLYSTA!ACTN = "E"
   Else
    DLYSTA!ACTN = "M"
  End If
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM FASDAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "EXC"
  DLYSTA!PCOD = "Direct Cr Entry"
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = 0
  DLYSTA!VBNO = ENTRYNO
  DLYSTA!AMNT = Val(CENVAT) + Val(EDUCESS) + Val(HEDUCESS)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub


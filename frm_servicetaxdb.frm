VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_servicetaxdb 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excise Debit Entry"
   ClientHeight    =   5490
   ClientLeft      =   2325
   ClientTop       =   2640
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9000
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   5475
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9657
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
      Begin VB.TextBox TXTCESS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox VBNO 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   1575
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
         ItemData        =   "frm_servicetaxdb.frx":0000
         Left            =   2040
         List            =   "frm_servicetaxdb.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox ADUTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox HEDUCESS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox EDUCESS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox CENVAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox TXTEXTRA5 
         Height          =   1455
         Left            =   4440
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox txtpurac 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   6495
      End
      Begin VB.TextBox txtvbno 
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1800
         Width           =   1575
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6120
         TabIndex        =   29
         Top             =   7560
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
         Image           =   "frm_servicetaxdb.frx":002E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   3720
         TabIndex        =   23
         Top             =   4800
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
         Image           =   "frm_servicetaxdb.frx":03C8
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4920
         TabIndex        =   24
         Top             =   4800
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
         Image           =   "frm_servicetaxdb.frx":1152
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   26
         Top             =   4800
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
         Image           =   "frm_servicetaxdb.frx":15A4
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
         Format          =   54591489
         CurrentDate     =   39347
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   2520
         TabIndex        =   0
         Top             =   4800
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
         Image           =   "frm_servicetaxdb.frx":19F6
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   6120
         TabIndex        =   25
         Top             =   4800
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
         Image           =   "frm_servicetaxdb.frx":1D90
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker txtodat 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
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
         Format          =   54591489
         CurrentDate     =   39347
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "P.Cess"
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
         Top             =   3000
         Width           =   855
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
         TabIndex        =   6
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "EXCISE DEBIT ENTRY "
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
         TabIndex        =   32
         Top             =   240
         Width           =   3375
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   540
         Left            =   5040
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3735
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
         TabIndex        =   5
         Top             =   720
         Width           =   1005
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1125
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
         TabIndex        =   1
         Top             =   720
         Width           =   1470
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1530
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
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9000
         Y1              =   1560
         Y2              =   1560
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
         TabIndex        =   20
         Top             =   4080
         Width           =   1215
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
         TabIndex        =   18
         Top             =   3720
         Width           =   1215
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
         TabIndex        =   16
         Top             =   3360
         Width           =   975
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
         TabIndex        =   13
         Top             =   2640
         Width           =   855
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
         Left            =   5640
         TabIndex        =   31
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   4680
         Width           =   8295
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp. A/c"
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
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
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
         TabIndex        =   30
         Top             =   1800
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frm_servicetaxdb"
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
  
  ENTRYNO.Caption = GenVNO("EXD", "000001")
  Call genJOURNAL
  VBNO.SetFocus
  EXCREG.ListIndex = 0
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
  ENTRYNO.Caption = GenVNO("EXD", "000001")
  SAVFLAG = True
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000075", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  SAVFLAG = False
  M_VBNO = Empty
  'frm_EXCISEdbList.Show 1
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
        "' AND SRNO='" & M_VBNO & "' AND VTYP='EXD'"
        CN.CommitTrans
        btn_sts (False)
    End If
  End If
  
  btn_sts (True)
  cmdAdd.SetFocus
  Call cmdCancel_Click
  Exit Sub
End Sub


Private Sub cmdEdit_Click()
  'If M_USRSECLEVL = "1" Then
  '   If ReadConfigMaster("0017", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  'End If
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000075", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  Else
       Call chk_ldt
  End If
  SAVFLAG = False
  M_VBNO = Empty
  frm_excisedblist.Show 1
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
  
  ENTDAT.Value = Now
  txtodat.Value = Now
  btn_sts True
  Me.KeyPreview = True
  If UNT_ISPAPPER = "Y" Then
    Label1.Visible = True
    TXTCESS.Visible = True
   Else
    Label1.Visible = False
    TXTCESS.Visible = False
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
Private Sub cmdSave_Click()
  
  On Error GoTo LAST
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ACCMST WHERE NAME='" & txtpurac & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Exp. A/c"
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
    SQL = "SELECT RG23ACENVAT AS CENVATAC,RG23ACESSAC AS PCESSAC, RG23AEDUCESS AS ECESSAC,RG23AHEDCESS AS HEDCSAC, '' AS DEFFEDAC FROM EXCISEOPENING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'"
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
     ENTRYNO.Caption = GenVNO("EXD", "000001")
  End If
  
  TXTEXTRA5 = Replace(Trim(TXTEXTRA5), vbCrLf, "")
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT CODE FROM ACCMST WHERE NAME='" & txtpurac & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Exp. A/c"
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
  
  
  
  
  If EXCREG = Empty Then
    MsgBox "Invalid Excise Register"
    EXCREG.SetFocus
    Exit Sub
  End If
  
  
  
  
  
  
  
  If Not IsNumeric(CENVAT) Then
    MsgBox "Invalid Cenvat Amount"
    CENVAT.SetFocus
    Exit Sub
  End If
  If UNT_ISPAPPER = "N" Then
    TXTCESS.Text = 0
  End If
  If Not IsNumeric(TXTCESS) Then
    MsgBox "Invalid P.Cess Amount"
    TXTCESS.SetFocus
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
    C_CENVAT = Val(CENVAT.Text)
    C_EDCESS = Val(EDUCESS.Text)
    C_HEDCES = Val(HEDUCESS.Text)
    C_PCESS = Val(TXTCESS.Text)
  End If
  'Delete Old Entries
  CN.BeginTrans
  If SAVFLAG = True Then
    Call genJOURNAL
  End If
  If Trim(txtvbno) = "" Then
    Call genJOURNAL
  End If
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' and unit='" & UNCD & "' AND VTYP='EXD' AND SRNO='" & ENTRYNO & "'"
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND VTYP='EXD' AND unit='" & UNCD & "' and SRNO='" & ENTRYNO & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
    SAVDAT.AddNew
  End If
  
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "EXD"
  SAVDAT!SRNO = ENTRYNO
  SAVDAT!SRCH = "1"
  SAVDAT!Date = Format(ENTDAT, "YYYY/MM/DD")
  SAVDAT!dbcd = "EXCREG"
  SAVDAT!CRAC = SUP_COD & ""
  SAVDAT!DRAC = MFG_COD & ""
  SAVDAT!VBNO = Trim(VBNO)
  SAVDAT!chln = Trim(VBNO)
  SAVDAT!ICOD = ITM_COD & ""
  SAVDAT!QNTY = 0
  SAVDAT!AMNT = 0
  SAVDAT!ITOT = 0
  SAVDAT!TTYP = EXCREG.Text
  SAVDAT!RECSTAT = "A"
  SAVDAT!unit = UNCD

  SAVDAT!CENVAT = Val(CENVAT)
  SAVDAT!CESS = Val(TXTCESS)
  SAVDAT!EDUCESS = Val(EDUCESS)
  SAVDAT!H_ED_CESS = Val(HEDUCESS)
  SAVDAT!A_DUTY = Val(ADUTY)

  SAVDAT!EXTRA1 = "Manufacture"
  SAVDAT!EXTRA3 = "True"
  SAVDAT!CHDT = Format(txtodat.Value, "YYYY/MM/DD")
  SAVDAT!EXTRA5 = Mid(Trim(TXTEXTRA5.Text), 1, 50)
  SAVDAT!EXTRA1 = txtpurac.Tag
  SAVDAT!EXTRA2 = txtvbno
  SAVDAT.Update
  CN.CommitTrans
  Dim TOTEXC As Double
  TOTEXC = Val(CENVAT) + Val(EDUCESS) + Val(HEDUCESS) + Val(ADUTY) + Val(TXTCESS)
  
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
  RS!damt = TOTEXC
  RS!camt = 0
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = TOTEXC
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXD"
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
  RS!damt = 0
  RS!camt = Val(CENVAT.Text) + Val(ADUTY)
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(CENVAT.Text) + Val(ADUTY)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXD"
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
  RS!damt = 0
  RS!camt = Val(TXTCESS.Text)
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(TXTCESS.Text)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXD"
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
  RS!damt = 0
  RS!camt = Val(EDUCESS.Text)
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(EDUCESS.Text)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXD"
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
  RS!damt = 0
  RS!camt = Val(HEDUCESS.Text)
  RS!VBNO = Trim(txtvbno)
  RS!narr = TXTEXTRA5
  RS!AMNT = Val(HEDUCESS.Text)
  RS!RECSTAT = "A"
  RS!MLTENT = "Y"
  RS!EXTRA1 = "EXD"
  RS!EXTRA2 = ENTRYNO
  RS.Update
  

  'Save New Record
  'Update excno in untmst
  If SAVFLAG = True Then
    
    CN.Execute "UPDATE SERIALMASTER SET [SRNO]='" & ENTRYNO & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
          "' AND VTYP='EXD' AND CODE='000001' AND FYCD='" & FYCD & "'"
          
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
    SendKeys "{tab}"
    Me.KeyPreview = True
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



Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM FASDAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "EXD"
  DLYSTA!PCOD = "Direct Debit Entry"
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



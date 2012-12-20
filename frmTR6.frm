VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmTR6 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TR-6 PAYMENT DETAILS OF SERVICE TAX (CHALLAN). "
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9120
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   7920
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   8235
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   14526
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
         Left            =   6000
         TabIndex        =   10
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtvbno 
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   35
         Top             =   6360
         Width           =   1455
      End
      Begin VB.TextBox txtchqno 
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   15
         Top             =   6360
         Width           =   1455
      End
      Begin VB.TextBox txtdbac 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2280
         Width           =   5895
      End
      Begin VB.TextBox txtcrac 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1800
         Width           =   5895
      End
      Begin VB.TextBox BSRCODE 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   16
         Top             =   6840
         Width           =   5895
      End
      Begin VB.TextBox CENVAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   9
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox NCCD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6000
         TabIndex        =   14
         Top             =   9480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox ADUTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   12
         Top             =   4800
         Width           =   1815
      End
      Begin VB.TextBox EDUCESS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   4320
         Width           =   1815
      End
      Begin VB.TextBox HEDCESS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   13
         Top             =   5280
         Width           =   1815
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6000
         TabIndex        =   20
         Top             =   8160
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
         Image           =   "frmTR6.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   7440
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
         Image           =   "frmTR6.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4800
         TabIndex        =   18
         Top             =   7440
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
         Image           =   "frmTR6.frx":1124
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7200
         TabIndex        =   21
         Top             =   7440
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
         Image           =   "frmTR6.frx":1576
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker TDAT 
         Height          =   315
         Left            =   6720
         TabIndex        =   4
         Top             =   1320
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
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   2400
         TabIndex        =   0
         Top             =   7440
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
         Image           =   "frmTR6.frx":19C8
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   6000
         TabIndex        =   19
         Top             =   7440
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
         Image           =   "frmTR6.frx":1D62
         cBack           =   -2147483633
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "PAPER CESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   36
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Db A/c Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cr A/c Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label LBLCHLN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXX"
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
         Left            =   2160
         TabIndex        =   2
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "B.S.R Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   33
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Line Line5 
         X1              =   2400
         X2              =   7920
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CENVAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NCCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   9480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "A.DUTY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "EDUCATION CESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   29
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Hr. EDU CESS."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         Height          =   3615
         Left            =   2400
         Top             =   2640
         Width           =   5535
      End
      Begin VB.Line Line3 
         X1              =   2400
         X2              =   7920
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line2 
         X1              =   5400
         X2              =   5400
         Y1              =   2640
         Y2              =   6240
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Details Of Duty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label TOTAMT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   24
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   8055
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8895
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   240
         X2              =   8880
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   9000
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "TR-6 PAYMENT DETAILS"
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
         TabIndex        =   23
         Top             =   240
         Width           =   3615
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label LBLCHHEAD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Challan No. "
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
         Left            =   720
         TabIndex        =   1
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Challan Date "
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
         Left            =   5280
         TabIndex        =   3
         Top             =   1320
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmTR6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SAVEFLAG As Boolean
Dim MCHLN As String
Dim bankcod As String

Private Sub ADUTY_Change()
  TOTAMT = Val(CENVAT) + Val(NCCD) + Val(ADUTY) + Val(EDUCESS) + Val(HEDCESS) + Val(TXTCESS)
End Sub

Private Sub ADUTY_GotFocus()
  ADUTY.BackColor = RGB(BRED, BGREEN, BBLUE)
  ADUTY.SelStart = 0
  ADUTY.SelLength = Len(ADUTY)
End Sub

Private Sub ADUTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And InStr(1, ADUTY, ".", vbTextCompare) > 0 Then KeyAscii = 0
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub ADUTY_LostFocus()
 ADUTY.BackColor = vbWhite
End Sub

Private Sub BSRCODE_GotFocus()
BSRCODE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub BSRCODE_LostFocus()
BSRCODE.BackColor = vbWhite
End Sub

Private Sub CENVAT_Change()
  TOTAMT = Val(CENVAT) + Val(NCCD) + Val(ADUTY) + Val(EDUCESS) + Val(HEDCESS) + Val(TXTCESS)
End Sub

Private Sub CENVAT_GotFocus()
  CENVAT.BackColor = RGB(BRED, BGREEN, BBLUE)
  CENVAT.SelStart = 0
  CENVAT.SelLength = Len(CENVAT)
End Sub

Private Sub CENVAT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And InStr(1, CENVAT, ".", vbTextCompare) > 0 Then KeyAscii = 0
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub CENVAT_LostFocus()
    CENVAT.BackColor = vbWhite
End Sub

Private Sub cmdAdd_Click()
    Call ClsData(Me)
    If Allow_view_only = "Y" Then
       cmdCancel_Click
       btn_sts (True)
       Exit Sub
    End If
    Call btn_sts(False)
    txtcrac.Enabled = True
    
    TDAT.MinDate = FSDT
    TDAT.MaxDate = FEDT
    
    TDAT.SetFocus
    
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000073", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    SAVEFLAG = False
    NEW_VISIBLE = False
    cmdCancel.Cancel = True
    Call btn_sts(False)
    
  MCHLN = SearchList1("select DISTINCT CHLN, CHLN from EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='TR6' AND RECSTAT='A'", 0, "", "List Of CHALLAN OF TR6")
  If MCHLN <> Empty Then Call FILLDET
  If ALLOW_EDIT(compPth, UNCD, "PAY", GetCode("ACCMST", txtcrac, "NAME", "CODE"), Trim(txtvbno)) = "C" Then
      cmdCancel_Click
      btn_sts (True)
      Exit Sub
  End If
  If Allow_view_only = "Y" Then
     cmdCancel_Click
     btn_sts (True)
     Exit Sub
  End If
  txtcrac.Enabled = False
  txtdbac.SetFocus
End Sub

Private Sub EDUCESS_Change()
  TOTAMT = Val(CENVAT) + Val(NCCD) + Val(ADUTY) + Val(EDUCESS) + Val(HEDCESS) + Val(TXTCESS)
End Sub

Private Sub EDUCESS_GotFocus()
  EDUCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
  EDUCESS.SelStart = 0
  EDUCESS.SelLength = Len(EDUCESS)
End Sub

Private Sub EDUCESS_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 And InStr(1, EDUCESS, ".", vbTextCompare) > 0 Then KeyAscii = 0
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub EDUCESS_LostFocus()
  EDUCESS.BackColor = vbWhite
End Sub

Private Sub HEDCESS_Change()
  TOTAMT = Val(CENVAT) + Val(NCCD) + Val(ADUTY) + Val(EDUCESS) + Val(HEDCESS) + Val(TXTCESS)
End Sub

Private Sub HEDCESS_GotFocus()
  HEDCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
  HEDCESS.SelStart = 0
  HEDCESS.SelLength = Len(HEDCESS)
End Sub

Private Sub HEDCESS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And InStr(1, HEDCESS, ".", vbTextCompare) > 0 Then KeyAscii = 0
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub HEDCESS_LostFocus()
 HEDCESS.BackColor = vbWhite
End Sub

Private Sub NCCD_Change()
  TOTAMT = Val(CENVAT) + Val(NCCD) + Val(ADUTY) + Val(EDUCESS) + Val(HEDCESS) + Val(TXTCESS)
End Sub
Private Sub NCCD_GotFocus()
  NCCD.BackColor = RGB(BRED, BGREEN, BBLUE)
  NCCD.SelStart = 0
  NCCD.SelLength = Len(NCCD)
End Sub

Private Sub NCCD_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 And InStr(1, NCCD, ".", vbTextCompare) > 0 Then KeyAscii = 0
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub NCCD_LostFocus()
 NCCD.BackColor = vbWhite
End Sub

Private Sub cmdCancel_Click()
  ClsData (Me)
  cmdExit.Cancel = True
  Call btn_sts(True)
  LBLCHLN = GenVNO("TR6", "000001")
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000073", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  If Allow_view_only = "Y" Then
     cmdCancel_Click
     btn_sts (True)
     Exit Sub
  End If
  Dim AYS
  AYS = MsgBox("Are You Sure To Delete the TR-6 Challan", vbYesNo)
  If AYS = vbYes Then
    CN.BeginTrans
    CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CHLN='" & Trim(LBLCHLN) & "' AND VTYP='TR6' AND RECSTAT='A'"
    Call UPDATEDELSTATUS
    CN.CommitTrans
  End If
  Call cmdCancel_Click
  CENVAT.Enabled = True
  CENVAT.SetFocus
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  On Error GoTo LAST
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & txtcrac & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    bankcod = RS!CODE & ""
    txtcrac.Tag = RS!CODE & ""
   Else
    MsgBox "Invalid Bank Name"
    txtcrac.SetFocus
    Exit Sub
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from accmst where name='" & txtdbac & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    txtdbac.Tag = RS!CODE & ""
   Else
    MsgBox "Invalid P.L.A A/c"
    txtdbac.SetFocus
    Exit Sub
  End If
  
  If SAVEFLAG = True Then
    If M_CUNT = "Y" Then
      txtvbno = genBANKVBNo("Select SRNO AS PVOU from SERIALMASTER where COMP='" & compPth & "' AND [CODE]='" & bankcod & "' AND UNIT='" & UNCD & "' AND FYCD='" & FYCD & "' AND VTYP='PAY'")
     Else
      txtvbno = genBANKVBNo("Select SRNO AS PVOU from SERIALMASTER where COMP='" & compPth & "' AND [CODE]='" & bankcod & "' AND FYCD='" & FYCD & "' AND VTYP='PAY'")
    End If
  End If
  If SAVEFLAG = True Then
    LBLCHLN = GenVNO("TR6", "000001")
  End If
  
  
  
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "select * from EGPMAN where comp='" & compPth & "' AND UNIT='" & UNCD & "' AND CHLN='" & Trim(LBLCHLN) & "' AND VTYP='TR6'", CN, adOpenDynamic, adLockOptimistic
  CN.BeginTrans
  If SAVDAT.EOF Then
    SAVDAT.AddNew
  End If
  SAVDAT!COMP = compPth
  SAVDAT!unit = UNCD
  SAVDAT!VTYP = "TR6"
  SAVDAT!TTYP = "PLAREG"
  SAVDAT!Date = Format(TDAT, "YYYY/MM/DD")
  SAVDAT!CHDT = Format(TDAT, "YYYY/MM/DD")
  SAVDAT!chln = Trim(LBLCHLN)
  SAVDAT!VBNO = Trim(LBLCHLN)
  SAVDAT!dbcd = "XXXXXX"
  SAVDAT!SRNO = "XXXXXXXXX"
  SAVDAT!SRCH = "0"
  SAVDAT!CENVAT = Val(CENVAT)
  SAVDAT!A_DUTY = Val(ADUTY)
  SAVDAT!NCCD = Val(NCCD)
  SAVDAT!CESS = Val(TXTCESS)
  SAVDAT!EDUCESS = Val(EDUCESS)
  SAVDAT!H_ED_CESS = Val(HEDCESS)
  SAVDAT!BRMK = Trim(BSRCODE)
  SAVDAT!RECSTAT = "A"
  SAVDAT!EXTRA1 = "PAY"
  SAVDAT!EXTRA2 = txtvbno
  SAVDAT!EXTRA3 = txtcrac.Tag
  SAVDAT!EXTRA4 = txtdbac.Tag
  SAVDAT!EXTRA5 = txtchqno
  
  SAVDAT.Update
   If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & txtcrac.Tag & "' AND VNO='" & txtvbno & "' AND VTYP='PAY' AND SRCH=1", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
    SAVDAT.AddNew
  End If
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "PAY"
  SAVDAT!SRNO = "0"
  SAVDAT!SRCH = 1
  SAVDAT!unit = UNCD
  SAVDAT!MSDVCD = "000001"
  SAVDAT!User = cUName
  SAVDAT!Date = Format(TDAT.Value, "YYYY/MM/DD")
  SAVDAT!dbcd = txtcrac.Tag
  SAVDAT!vno = txtvbno
  SAVDAT!ACOD = txtcrac.Tag
  SAVDAT!RCOD = txtdbac.Tag
  SAVDAT!BRCD = ""
  SAVDAT!damt = 0
  SAVDAT!camt = Val(TOTAMT)
  SAVDAT!VBNO = Trim(txtvbno)
  SAVDAT!cdno = txtchqno
  SAVDAT!cddt = Format(TDAT.Value, "YYYY/MM/DD")
  SAVDAT!narr = "Being Amount Paid In PLA A/c"
  SAVDAT!AMNT = Val(TOTAMT)
  SAVDAT!DUDT = Format(TDAT.Value, "YYYY/MM/DD")
  SAVDAT!RCON = "N"
  SAVDAT!RECSTAT = "A"
  SAVDAT!MLTENT = "N"
  SAVDAT!RTYP = "EXC"
  SAVDAT!RSRN = Trim(LBLCHLN)
  SAVDAT.Update
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & txtcrac.Tag & "' AND VNO='" & txtvbno & "' AND VTYP='PAY' AND SRCH=2", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
    SAVDAT.AddNew
  End If
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "PAY"
  SAVDAT!SRNO = "0"
  SAVDAT!SRCH = 2
  SAVDAT!unit = UNCD
  SAVDAT!MSDVCD = "000001"
  SAVDAT!User = cUName
  SAVDAT!Date = Format(TDAT.Value, "YYYY/MM/DD")
  SAVDAT!dbcd = txtcrac.Tag
  SAVDAT!vno = txtvbno
  SAVDAT!ACOD = txtdbac.Tag
  SAVDAT!RCOD = txtcrac.Tag
  SAVDAT!BRCD = ""
  SAVDAT!damt = Val(TOTAMT)
  SAVDAT!camt = 0
  SAVDAT!VBNO = Trim(txtvbno)
  SAVDAT!cdno = txtchqno
  SAVDAT!cddt = Format(TDAT.Value, "YYYY/MM/DD")
  SAVDAT!narr = "Being Amount Paid In PLA A/c"
  SAVDAT!AMNT = Val(TOTAMT)
  SAVDAT!DUDT = Format(TDAT.Value, "YYYY/MM/DD")
  SAVDAT!RCON = "N"
  SAVDAT!RECSTAT = "A"
  SAVDAT!MLTENT = "N"
  SAVDAT!RTYP = "EXC"
  SAVDAT!RSRN = Trim(LBLCHLN)
  SAVDAT.Update
  
  If SAVEFLAG Then
    CN.Execute "UPDATE SERIALMASTER SET [SRNO]='" & LBLCHLN & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
        "' AND VTYP='TR6' AND CODE='000001' AND FYCD='" & FYCD & "'"
    If M_CUNT = "Y" Then
      CN.Execute "UPDATE SERIALMASTER SET SRNO='" & Trim(Mid(txtvbno, 1, 5)) & "' WHERE COMP='" & compPth & "' AND CODE='" & txtcrac.Tag & "' AND UNIT='" & UNCD & "' AND VTYP='PAY' AND FYCD='" & FYCD & "'"
     Else
      CN.Execute "UPDATE SERIALMASTER SET SRNO='" & Trim(Mid(txtvbno, 1, 5)) & "' WHERE COMP='" & compPth & "' AND CODE='" & txtcrac.Tag & "' AND VTYP='PAY' AND FYCD='" & FYCD & "'"
    End If
  End If
  
  
  
  If SAVEFLAG Then
     MsgBox "TR-6 Challan No. " & Trim(LBLCHLN) & " Successfiully Saved."
  Else
     MsgBox "TR-6 Challan No. " & Trim(LBLCHLN) & " Successfiully Edited."
  End If
  
  Call UPDATESTATUS
  CN.CommitTrans
  Call cmdCancel_Click
  If CENVAT.Enabled Then CENVAT.SetFocus
  Exit Sub
LAST:
 MsgBox ERR.Description
 Resume
 CN.RollbackTrans
 If SAVDAT.State = 1 Then
   SAVDAT.CancelUpdate
 End If
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
   SendKeys "{TAB}"
 End If
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  Me.KeyPreview = True
  TDAT = Now
  TDAT.MaxDate = FEDT
  TDAT.MinDate = FSDT
  Call btn_sts(True)
  LBLCHLN = GenVNO("TR6", "000001")
  If UNT_ISPAPPER = "Y" Then
    Label15.Visible = True
    TXTCESS.Visible = True
   Else
    Label15.Visible = False
    TXTCESS.Visible = False
  End If
End Sub

Private Sub TDAT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TR6N_LostFocus()
  TR6N.BackColor = vbWhite
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM PLATRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RCOD='PLAREG' AND TR6N='" & TR6N & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
    TDAT = SAVDAT!TDAT
    CENVAT = SAVDAT!CENVAT
    NCCD = SAVDAT!NCCD
    ADUTY = SAVDAT!ADUTY
    EDUCESS = SAVDAT!EDUCESS
    HEDCESS = SAVDAT!HEDCESS
    TXTCESS = SAVDAT!CESS
    BSRCODE = Trim(SAVDAT!BSRCODE)
   Else
    CENVAT = 0
    NCCD = 0
    ADUTY = 0
    EDUCESS = 0
    HEDCESS = 0
    TXTCESS = 0
    BSRCODE = Empty
  End If
End Sub

Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "TR6"
  DLYSTA!PCOD = "TR6 ENTRY"
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = 0
  DLYSTA!VBNO = TR6N & ""
  DLYSTA!AMNT = Val(TOTAMT)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "E"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "TR6"
  DLYSTA!PCOD = "TR6 ENTRY"
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = 0
  DLYSTA!VBNO = LBLCHLN & ""
  DLYSTA!AMNT = Val(TOTAMT)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    CENVAT.Enabled = Not bool
    NCCD.Enabled = Not bool
    TXTCESS.Enabled = Not bool
    EDUCESS.Enabled = Not bool
    ADUTY.Enabled = Not bool
    HEDCESS.Enabled = Not bool
End Sub

Private Sub FILLDET()
Dim TMPRS As ADODB.Recordset
Set TMPRS = New ADODB.Recordset
If TMPRS.State = 1 Then TMPRS.Close
TMPRS.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='TR6' AND CHLN='" & MCHLN & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not TMPRS.EOF Then
   TDAT = Format(TMPRS!Date, "DD/MM/YYYY")
   CHDT = Format(TMPRS!Date, "DD/MM/YYYY")
   LBLCHLN.Caption = TMPRS!chln & ""
   CENVAT = TMPRS!CENVAT
   ADUTY = TMPRS!A_DUTY
   NCCD = TMPRS!NCCD
   TXTCESS = TMPRS!CESS
   EDUCESS = TMPRS!EDUCESS
   HEDCESS = TMPRS!H_ED_CESS
   BSRCODE = TMPRS!BRMK
   txtcrac = GetCode("ACCMST", TMPRS!EXTRA3 & "", "CODE", "NAME")
   txtdbac = GetCode("ACCMST", TMPRS!EXTRA4 & "", "CODE", "NAME")
   txtchqno = TMPRS!EXTRA5 & ""
   txtvbno = TMPRS!EXTRA2 & ""
End If
End Sub

Private Sub TimerBillNo1_Timer()
    Static ctr As Integer
    If ctr Mod 45 = 0 And ctr <= 45 Then
       lblAlert.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
       BORDER.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
       LBLCHLN.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
       LBLCHHEAD.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
    ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
       lblAlert.ForeColor = vbRed
       BORDER.BorderColor = vbRed
       LBLCHLN.ForeColor = vbRed
       LBLCHHEAD.ForeColor = vbRed
    ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
       lblAlert.ForeColor = vbBlue
       BORDER.BorderColor = vbBlue
       LBLCHLN.ForeColor = vbBlue
       LBLCHHEAD.ForeColor = vbBlue
       ctr = 0
    End If
    ctr = ctr + 15
End Sub

Private Sub TXTcrAC_GotFocus()
 txtcrac.BackColor = RGB(BRED, BGREEN, BBLUE)
  If FIXNAM <> Empty Then
    txtcrac.Text = FIXNAM
  End If
End Sub

Private Sub TXTcrAC_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(txtcrac.Text) = Empty Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        FIXNAM = Empty
        If M_CUNT = "Y" Then
          txtcrac.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM BANKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, txtcrac, "List of Receipt Day Book")
         Else
          txtcrac.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM BANKMST WHERE COMP='" & compPth & "'", 0, txtcrac, "List of Receipt Day Book")
        End If
    ElseIf KeyCode = vbKeyDelete Then
        txtcrac = Empty
    End If
    Me.KeyPreview = True
    FIXNAM = txtcrac.Text
End Sub
Private Sub TXTcrAC_LostFocus()
  txtcrac.BackColor = vbWhite
  If txtcrac = Empty Then Exit Sub
  Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "select CODE,PVOU from BANKMST where comp='" & compPth & "'  AND NAME='" & txtcrac & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    Unload Me
  End If
  bankcod = RS!CODE
  Dim BOK_GEN As Boolean
  If SAVEFLAG = True Then
    If SAVEFLAG = True Then
      If M_CUNT = "Y" Then
        txtvbno = genBANKVBNo("Select SRNO AS PVOU from SERIALMASTER where COMP='" & compPth & "' AND [CODE]='" & bankcod & "' AND UNIT='" & UNCD & "' AND FYCD='" & FYCD & "' AND VTYP='PAY'")
       Else
        txtvbno = genBANKVBNo("Select SRNO AS PVOU from SERIALMASTER where COMP='" & compPth & "' AND [CODE]='" & bankcod & "' AND FYCD='" & FYCD & "' AND VTYP='PAY'")
      End If
    End If
  End If
  If txtvbno = "XXX" Then
    MsgBox "Invalid Bank"
    cmdCancel_Click
    Exit Sub
  End If
  FIXNAM = txtcrac.Text
End Sub
Private Function genBANKVBNo(ByVal SQL As String) As String
    Dim TEMPRS As New ADODB.Recordset, ctr As Double
    TEMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF Then
      genBANKVBNo = "XXX"
      Exit Function
    End If
    If TEMPRS![PVOU] = "" Or IsNull(TEMPRS![PVOU]) = True Then
        genBANKVBNo = "00001"
        genBANKVBNo = genBANKVBNo + Mid(Year(FSDT), 3, 2) + Mid(Year(FEDT), 3, 2)
        Exit Function
    End If
    ctr = Val(Trim(TEMPRS![PVOU]))
    ctr = ctr + 1
    If ctr < 10 Then
        genBANKVBNo = "0000" + CStr(ctr)
    ElseIf ctr >= 10 And ctr < 100 Then
        genBANKVBNo = "000" + CStr(ctr)
    ElseIf ctr >= 100 And ctr < 1000 Then
        genBANKVBNo = "00" + CStr(ctr)
    ElseIf ctr >= 1000 And ctr < 10000 Then
        genBANKVBNo = "0" + CStr(ctr)
    Else
        genBANKVBNo = CStr(ctr)
    End If
    genBANKVBNo = genBANKVBNo + Mid(Year(FSDT), 3, 2) + Mid(Year(FEDT), 3, 2)
End Function
Private Sub TXTDBAC_GotFocus()
 txtdbac.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDBAC_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or Trim(txtdbac.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtdbac.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST", 0, txtdbac, "Select P.L.A A/c")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            txtdbac.Text = ""
            frm_Acc.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        txtdbac = Empty
    End If
    Me.KeyPreview = True
End Sub
Private Sub TXTCESS_Change()
  TOTAMT = Val(CENVAT) + Val(NCCD) + Val(ADUTY) + Val(EDUCESS) + Val(HEDCESS) + Val(TXTCESS)
End Sub
Private Sub TXTCESS_GotFocus()
  TXTCESS.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTCESS.SelStart = 0
  TXTCESS.SelLength = Len(TXTCESS)
End Sub

Private Sub TXTCESS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And InStr(1, TXTCESS, ".", vbTextCompare) > 0 Then KeyAscii = 0
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub TXTCESS_LostFocus()
    TXTCESS.BackColor = vbWhite
End Sub

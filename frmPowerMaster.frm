VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmPowerMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POWER MASTER"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   21
      Top             =   5640
      Width           =   5775
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   960
         TabIndex        =   16
         Top             =   240
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
         Image           =   "frmPowerMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdClear 
         Height          =   495
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Clear"
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
         Image           =   "frmPowerMaster.frx":059A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   3840
         TabIndex        =   18
         Top             =   240
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
         Image           =   "frmPowerMaster.frx":09EC
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Width           =   5775
      Begin VB.TextBox TXTOPNPRMUNIT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox TXTPRMRATE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox TXTOPNSEBUNIT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   15
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox TXTMINGASRATE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TXTMINGASUNIT 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TXTMINGASAMT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TXTOPNGGUNIT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   13
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox TXTSEBRATE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   9
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "OPENING OF PREMIUM UNIT :"
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
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label Label7 
         Caption         =   "PREMIUM GAS RATE :"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "OPENING OF STATE ELEC. BOARD (SEB) UNIT :"
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
         Left            =   120
         TabIndex        =   14
         Top             =   4200
         Width           =   4455
      End
      Begin VB.Label Label6 
         Caption         =   "GAS RATE  :"
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
         Left            =   120
         TabIndex        =   2
         Tag             =   "S"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "MINIMUM GAS UNIT  :"
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
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "MINIMUM GAS AMOUNT"
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
         TabIndex        =   4
         Tag             =   "S"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "STATE ELECTRICITY BOARD RATE :"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "OPENING OF GAS GENERATOR (GG) UNIT :"
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
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   4215
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   6120
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   6615
      Left            =   120
      Top             =   120
      Width           =   6015
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   6120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "POWER MASTER"
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
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmPowerMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************
'Module Name : POWER MASTER
'
'Develope By :
'
'Develope Date : 06 JULY 2011
'
'Change Date :
'
'Change By :
'
'Remark :
'*******************************************

Option Explicit
Dim SWITCH As Boolean
Dim M_DVCD As String
Dim M_BXMC As String
Dim ROWNO As Long
Dim SAVEFLAG As Boolean
Dim M_CODE As String

Private Sub cmdClear_Click()
 TXTMINGASAMT = Empty
 TXTMINGASUNIT = Empty
 TXTMINGASRATE = Empty
 TXTSEBRATE = Empty
 TXTOPNGGUNIT = Empty
 TXTOPNSEBUNIT = Empty
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  LBLHEAD.BackColor = &H80&
  LBLHEAD.ForeColor = &HFFFFFF
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  cmdExit.Cancel = True
  Me.KeyPreview = True
  Call FindDetail
  
  Exit Sub
errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdSave_Click()
   On Error GoTo LAST
             
   If Val(TXTMINGASAMT) = 0 And Val(TXTMINGASUNIT) = 0 And Val(TXTMINGASRATE) = 0 And Val(TXTPRMRATE) = 0 And Val(TXTSEBRATE) = 0 And Val(TXTOPNGGUNIT) = 0 And Val(TXTOPNSEBUNIT) = 0 Then
      MsgBox "INVALID DETAILS !!", vbInformation
      Exit Sub
   End If
            
   CN.BeginTrans
   CN.Execute "DELETE FROM POWERMAST"
    
   CN.Execute "INSERT INTO POWERMAST(COMP,UNIT,MINGASUNIT,MINGASRATE,MINGASAMT,PRMRATE,SEBRATE,OPNPRMUNIT,OPNGGUNIT,OPNSEBUNIT) " & _
               "VALUES('" & compPth & "','" & UNCD & "','" & Val(TXTMINGASUNIT) & "','" & Val(TXTMINGASRATE) & _
               "','" & Val(TXTMINGASAMT) & "','" & Val(TXTPRMRATE) & "','" & Val(TXTSEBRATE) & "','" & Val(TXTOPNPRMUNIT) & "','" & Val(TXTOPNGGUNIT) & _
               "','" & Val(TXTOPNSEBUNIT) & "')"
    
   CN.CommitTrans
   MsgBox "Data Successfully Saved"
    
Exit Sub
LAST:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub

Private Sub FindDetail()
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT * FROM POWERMAST WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & "' ", CN, adOpenDynamic, adLockOptimistic
If TEMPRS.EOF = False Then
 TXTMINGASUNIT = Trim(nstr(Val(TEMPRS!MINGASUNIT), 15, 2))
 TXTMINGASRATE = Val(TEMPRS!MINGASRATE)
 TXTMINGASAMT = Trim(nstr(Val(TEMPRS!MINGASAMT), 15, 2))
 TXTPRMRATE = Trim(nstr(Val(TEMPRS!PRMRATE), 15, 2))
 TXTSEBRATE = Trim(nstr(Val(TEMPRS!SEBRATE), 15, 2))
 TXTOPNPRMUNIT = Val(TEMPRS!OPNPRMUNIT)
 TXTOPNGGUNIT = Val(TEMPRS!OPNGGUNIT)
 TXTOPNSEBUNIT = Val(TEMPRS!OPNSEBUNIT)
End If
TEMPRS.Close

End Sub

Private Sub TXTMINGASRATE_Change()
  TXTMINGASAMT = Trim(nstr(Val(TXTMINGASRATE) * Val(TXTMINGASUNIT), 15, 2))
  TXTMINGASAMT = Round(Val(TXTMINGASAMT), 0)
End Sub

Private Sub TXTMINGASUNIT_Change()
   TXTMINGASAMT = Trim(nstr(Val(TXTMINGASRATE) * Val(TXTMINGASUNIT), 15, 2))
   TXTMINGASAMT = Round(Val(TXTMINGASAMT), 0)
End Sub

Private Sub TXTMINGASUNIT_GotFocus(): TXTMINGASUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTMINGASUNIT_LostFocus(): TXTMINGASUNIT.BackColor = vbWhite: End Sub

Private Sub TXTMINGASUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTMINGASUNIT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTMINGASRATE_GotFocus(): TXTMINGASRATE.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTMINGASRATE_LostFocus(): TXTMINGASRATE.BackColor = vbWhite: End Sub

Private Sub TXTMINGASRATE_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTMINGASRATE, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTMINGASAMT_GotFocus(): TXTMINGASAMT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTMINGASAMT_LostFocus(): TXTMINGASAMT.BackColor = vbWhite: End Sub

Private Sub TXTMINGASAMT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTMINGASAMT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTSEBRATE_GotFocus(): TXTSEBRATE.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTSEBRATE_LostFocus(): TXTSEBRATE.BackColor = vbWhite: End Sub

Private Sub TXTSEBRATE_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTSEBRATE, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTPRMRATE_GotFocus(): TXTPRMRATE.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTPRMRATE_LostFocus(): TXTPRMRATE.BackColor = vbWhite: End Sub

Private Sub TXTPRMRATE_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTPRMRATE, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTOPNGGUNIT_GotFocus(): TXTOPNGGUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTOPNGGUNIT_LostFocus(): TXTOPNGGUNIT.BackColor = vbWhite: End Sub

Private Sub TXTOPNGGUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTOPNGGUNIT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTOPNPRMUNIT_GotFocus(): TXTOPNPRMUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTOPNPRMUNIT_LostFocus(): TXTOPNPRMUNIT.BackColor = vbWhite: End Sub

Private Sub TXTOPNPRMUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTOPNPRMUNIT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTOPNSEBUNIT_GotFocus(): TXTOPNSEBUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTOPNSEBUNIT_LostFocus(): TXTOPNSEBUNIT.BackColor = vbWhite: End Sub

Private Sub TXTOPNSEBUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTOPNSEBUNIT, Me) = 0 Then KeyAscii = 0
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmPowerTran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAILY POWER TRANSACTION ENTRY"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8265
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   7815
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1800
         TabIndex        =   7
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
         Image           =   "frmPowerTran.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdClear 
         Height          =   495
         Left            =   3480
         TabIndex        =   8
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
         Image           =   "frmPowerTran.frx":059A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5160
         TabIndex        =   9
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
         Image           =   "frmPowerTran.frx":09EC
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   7815
      Begin VB.TextBox TXTOPNPRMUNIT 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3600
         MaxLength       =   40
         TabIndex        =   15
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox TXTOPNSEBUNIT 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   5
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox TXTCLOPRMUNIT 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   5760
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox TXTCLOSEBUNIT 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   5760
         MaxLength       =   40
         TabIndex        =   6
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox TXTOPNGGUNIT 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   3
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox TXTCLOGGUNIT 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   5760
         MaxLength       =   40
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker TXTDT 
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
         Format          =   50069505
         CurrentDate     =   39347
      End
      Begin VB.Label Label4 
         Caption         =   "STATE ELECTRICITY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Tag             =   "S"
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "OPENING      UNIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   3840
         TabIndex        =   17
         Tag             =   "S"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "GAS GENERATOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Tag             =   "S"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   240
         X2              =   7680
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   240
         X2              =   7680
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2535
         Left            =   240
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "CLOSING    UNIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   6000
         TabIndex        =   14
         Tag             =   "S"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "PREMIUM GAS "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Tag             =   "S"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   5640
         X2              =   5640
         Y1              =   840
         Y2              =   4200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   3480
         X2              =   7680
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   3375
         Left            =   3480
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "DATE :"
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
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   6015
      Left            =   120
      Top             =   120
      Width           =   8055
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   8160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "DAILY POWER TRANSACTION ENTRY"
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
      Left            =   1440
      TabIndex        =   11
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmPowerTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************
'Module Name : DAILY POWER TRANSACTION
'
'Develope By :
'
'Develope Date : 07 JULY 2011
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

Private Sub cmdCLEAR_Click()
 TXTCLOPRMUNIT = Empty
 TXTOPNGGUNIT = Empty
 TXTCLOGGUNIT = Empty
 TXTOPNSEBUNIT = Empty
 TXTCLOSEBUNIT = Empty
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
  Call SetLastDateForEntry
  TXTDT = Now
  TXTDT.MinDate = FSDT
  TXTDT.MaxDate = FEDT
    
  Exit Sub
errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdSave_Click()
   On Error GoTo LAST
             
   If Val(TXTCLOPRMUNIT) = 0 And Val(TXTCLOSEBUNIT) = 0 And Val(TXTCLOGGUNIT) = 0 And Val(TXTOPNGGUNIT) = 0 And Val(TXTOPNSEBUNIT) = 0 Then
      MsgBox "INVALID DETAILS !!", vbInformation
      Exit Sub
   End If
            
   CN.BeginTrans
   
   If RS.State = 1 Then RS.Close
   RS.Open "SELECT * FROM POWERTRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DATE='" & Format(TXTDT, "MM/DD/YYYY") & "'", CN, adOpenDynamic, adLockOptimistic
   If RS.EOF Then
      RS.AddNew
   End If
   RS!COMP = compPth
   RS!unit = UNCD
   RS!Date = Format(TXTDT.Value, "YYYY/MM/DD")
   RS!PRMGAS = Val(TXTCLOPRMUNIT)
   RS!GGUNIT = Val(TXTCLOGGUNIT)
   RS!SEBUNIT = Val(TXTCLOSEBUNIT)
   RS.Update
   RS.Close
      
   CN.Execute "UPDATE POWERMAST SET [OPNPRMUNIT]='" & Val(TXTCLOPRMUNIT) & "',[OPNGGUNIT]='" & Val(TXTCLOGGUNIT) & _
              "',[OPNSEBUNIT]='" & Val(TXTCLOSEBUNIT) & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' "
   
   CN.Execute "UPDATE POWERMAST SET [LSTENTRYDT]='" & Format(TXTDT, "MM/DD/YYYY") & "' WHERE COMP='" & compPth & _
              "' AND UNIT='" & UNCD & "' "
   
   CN.CommitTrans
   MsgBox "Data Successfully Saved"
    
Exit Sub
LAST:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub

Private Sub TXTCLOSEBUNIT_GotFocus(): TXTCLOSEBUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTCLOSEBUNIT_LostFocus(): TXTCLOSEBUNIT.BackColor = vbWhite: End Sub

Private Sub TXTCLOSEBUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTCLOSEBUNIT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTCLOGGUNIT_GotFocus(): TXTCLOGGUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTCLOGGUNIT_LostFocus(): TXTCLOGGUNIT.BackColor = vbWhite: End Sub

Private Sub TXTCLOGGUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTCLOGGUNIT, Me) = 0 Then KeyAscii = 0
End Sub


Private Sub TXTDT_Change()
   Call SETREADING
End Sub

Private Sub TXTDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub TXTDT_LostFocus()
Call SETREADING
End Sub

Private Sub TXTCLOPRMUNIT_GotFocus(): TXTCLOPRMUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTCLOPRMUNIT_LostFocus(): TXTCLOPRMUNIT.BackColor = vbWhite: End Sub

Private Sub TXTCLOPRMUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTCLOPRMUNIT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTOPNGGUNIT_GotFocus(): TXTOPNGGUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTOPNGGUNIT_LostFocus(): TXTOPNGGUNIT.BackColor = vbWhite: End Sub

Private Sub TXTOPNGGUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTOPNGGUNIT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTOPNSEBUNIT_GotFocus(): TXTOPNSEBUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTOPNSEBUNIT_LostFocus(): TXTOPNSEBUNIT.BackColor = vbWhite: End Sub

Private Sub TXTOPNSEBUNIT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTOPNSEBUNIT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub SetLastDateForEntry()
Dim DTRS As ADODB.Recordset
Set DTRS = New ADODB.Recordset
    
    If DTRS.State = 1 Then DTRS.Close
    DTRS.Open "SELECT OPNGGUNIT,OPNSEBUNIT FROM POWERMAST " & _
              "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not DTRS.EOF Then
       TXTOPNGGUNIT = Val(DTRS!OPNGGUNIT)
       TXTOPNSEBUNIT = Val(DTRS!OPNSEBUNIT)
    End If
    DTRS.Close
    
    If DTRS.State = 1 Then DTRS.Close
    DTRS.Open "SELECT IsNull(LSTENTRYDT,'" & FSDT & "') AS LSTENTRYDT FROM POWERMAST " & _
              "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not DTRS.EOF Then
       TXTDT.MinDate = Format(DTRS!LSTENTRYDT, "DD/MM/YYYY")
       TXTDT = Format(DTRS!LSTENTRYDT, "DD/MM/YYYY")
       TXTDT = TXTDT + 1
    End If
    DTRS.Close
   
End Sub

Private Sub SETREADING()
   Dim DTRS As ADODB.Recordset
   Set DTRS = New ADODB.Recordset
   
   Exit Sub
   
   If DTRS.State = 1 Then DTRS.Close
   DTRS.Open "SELECT OPNGGUNIT,OPNSEBUNIT FROM POWERMAST " & _
             "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
   If Not DTRS.EOF Then
      TXTOPNGGUNIT = Val(DTRS!OPNGGUNIT)
      TXTOPNSEBUNIT = Val(DTRS!OPNSEBUNIT)
   End If
   DTRS.Close
   
   If RS.State = 1 Then RS.Close
   RS.Open "SELECT * FROM POWERTRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND DATE='" & Format(TXTDT, "MM/DD/YYYY") & "'", CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      
      TXTCLOPRMUNIT = Val(RS!PRMGAS)
      TXTCLOGGUNIT = Val(RS!GGUNIT)
      TXTCLOSEBUNIT = Val(RS!SEBUNIT)
      
      TXTOPNGGUNIT = Val(TXTOPNGGUNIT) - Val(TXTCLOGGUNIT)
      TXTOPNSEBUNIT = Val(TXTOPNSEBUNIT) - Val(TXTCLOSEBUNIT)
   Else
      TXTCLOPRMUNIT = Empty
      TXTCLOGGUNIT = Empty
      TXTCLOSEBUNIT = Empty
   End If
End Sub

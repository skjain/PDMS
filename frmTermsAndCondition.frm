VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmTermsAndCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TERMS AND CONDITION MASTER"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10635
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   240
      TabIndex        =   26
      Top             =   720
      Width           =   10215
      Begin VB.TextBox TXTDET5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   70
         TabIndex        =   25
         Top             =   3480
         Width           =   5535
      End
      Begin VB.TextBox TXTDET4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   70
         TabIndex        =   23
         Top             =   2760
         Width           =   5535
      End
      Begin VB.TextBox TXTDET3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   70
         TabIndex        =   21
         Top             =   2040
         Width           =   5535
      End
      Begin VB.TextBox TXTDET2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   70
         TabIndex        =   19
         Top             =   1320
         Width           =   5535
      End
      Begin VB.ComboBox CMBTYP 
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
         Left            =   120
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtname 
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
         TabIndex        =   7
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox TXTDET1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   70
         TabIndex        =   17
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox TXTJURI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   11
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox TXTEXP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   15
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox TXTNOTI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   13
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Label Label10 
         Caption         =   "Printing detail - 5 "
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
         Left            =   4440
         TabIndex        =   24
         Tag             =   "S"
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Printing detail - 4"
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
         Left            =   4440
         TabIndex        =   22
         Tag             =   "S"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Printing detail - 3"
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
         Left            =   4440
         TabIndex        =   20
         Tag             =   "S"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Printing detail - 2"
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
         Left            =   4440
         TabIndex        =   18
         Tag             =   "S"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "T&&C TYPE  :"
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
         Tag             =   "S"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "NAME    :"
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
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Printing detail - 1"
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
         Left            =   4440
         TabIndex        =   16
         Tag             =   "S"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "JURISDICTION :"
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
         Tag             =   "S"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "NOTIFICATION NO."
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
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "EXEMPTED NOTIFICATION NO."
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
         Top             =   3240
         Width           =   3015
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   4920
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
      Image           =   "frmTermsAndCondition.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4920
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
      Image           =   "frmTermsAndCondition.frx":039A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   4920
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
      Image           =   "frmTermsAndCondition.frx":0734
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   4920
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
      Image           =   "frmTermsAndCondition.frx":0ACE
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   4920
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
      Image           =   "frmTermsAndCondition.frx":0F20
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   4920
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
      Image           =   "frmTermsAndCondition.frx":1372
      cBack           =   -2147483633
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   5415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   240
      X2              =   10560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "TERMS AND CONDITION MASTER"
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
      Left            =   3120
      TabIndex        =   27
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmTermsAndCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************
'Module Name : TERMS AND CONDITION MASTER
'
'Develope By :
'
'Develope Date : 02 FEBRUARY 2011
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

Private Sub CMBTYP_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TXTJURI_GotFocus()
  TXTJURI.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTJURI.SelStart = 0
  TXTJURI.SelLength = Len(TXTJURI)
  TXTJURI.ToolTipText = "Enter Jurisdiction Details"
End Sub

Private Sub TXTJURI_LostFocus()
  TXTJURI.BackColor = vbWhite
End Sub

Private Sub TXTNOTI_GotFocus()
  TXTNOTI.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTNOTI.SelStart = 0
  TXTNOTI.SelLength = Len(TXTNOTI)
  TXTNOTI.ToolTipText = "Enter Party Address2"
End Sub

Private Sub TXTNOTI_LostFocus()
  TXTNOTI.BackColor = vbWhite
End Sub

Private Sub cmdAdd_Click()
If M_USRSECLEVL = "1" Then
   If ReadConfigMaster("0008", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
Else
   'Call chk_ldt
End If
    
Call ClsData(Me)
CMBTYP.ListIndex = 0
Call btn_sts(False)
Me.Frame1.Enabled = True
txtName.SetFocus
SAVEFLAG = False
cmdCancel.Cancel = True
cmdDelETE.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000005", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  On Error GoTo ErrorHandler
  If (MsgBox("Are You sure you want to delete :" & txtName, vbYesNo) = vbYes) Then
     CN.BeginTrans
     CN.Execute "UPDATE BILLTERMS SET STATUS = 'D',RECSTAT='D' WHERE CODE = '" & Trim(M_CODE) & "'"
     CN.CommitTrans
     MsgBox "Successfully Deleted"
  End If
  cmdCancel_Click
  Exit Sub
ErrorHandler:
    If (Err.Number = -2147168237) Then
       CN.RollbackTrans
    Else
       MsgBox Err.Description
    End If
End Sub

Private Sub cmdEdit_Click()
If M_USRSECLEVL = 1 Then
    If ReadConfigMaster("000005", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
End If
SAVEFLAG = True
M_DESC = Empty
Key = Empty
Frm_List.cmdAddNew.Visible = False
M_CODE = SearchList("Select DISTINCT(CODE) ,[NAME] from BILLTERMS where RECSTAT='A'")
If Trim(M_CODE) = Empty Then MsgBox "No Record Found ": cmdCancel_Click: Exit Sub
Call entry(M_CODE)
Call btn_sts(False)
Me.Frame1.Enabled = True
txtName.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  lblHead.BackColor = &H80&
  lblHead.ForeColor = &HFFFFFF
  CMBTYP.ListIndex = 0
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
  Call btn_sts(True)
  cmdExit.Cancel = True
  Me.Frame1.Enabled = False
  Me.KeyPreview = True
  M_CODE = Empty
  CMBTYP.AddItem "SALE"
  CMBTYP.AddItem "PURCHASE"
  CMBTYP.AddItem "WORKORDER"
  CMBTYP.AddItem "PURCHASEORDER"
  CMBTYP.AddItem "SALEORDER"
  Exit Sub

errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdSave_Click()
Dim INDEX As Long
   On Error GoTo LAST
             
   If Trim(txtName) = Empty Then
        MsgBox "Please Enter Party !!", vbInformation
        txtName.SetFocus
        Exit Sub
   End If
            
   If SAVEFLAG = False Then
        Dim CHKRS As ADODB.Recordset
        Set CHKRS = New ADODB.Recordset
        
        If CHKRS.State = 1 Then CHKRS.Close
        CHKRS.Open "Select * from BILLTERMS where Upper([NAME])='" & UCase(txtName.Text) & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
        If CHKRS.EOF = False Then MsgBox "This Name Is Already In Use !!!", vbInformation, App.Title: CHKRS.Close: txtName.SetFocus: Exit Sub
        M_CODE = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From BILLTERMS")
   End If
         
    CN.BeginTrans
    CN.Execute "DELETE FROM BILLTERMS WHERE CODE='" & M_CODE & "'"
    
    CN.Execute "INSERT INTO BILLTERMS(CODE,NAME,TYPE,SRNO,JURI,NOTI,EXP,DET1,DET2,DET3,DET4,DET5,STATUS,RECSTAT) " & _
               "VALUES('" & M_CODE & "','" & txtName & "','" & Trim(CMBTYP.Text) & "','" & INDEX & "','" & TXTJURI & _
                "','" & TXTNOTI & "','" & TXTEXP & "','" & Trim(TXTDET1) & "','" & Trim(TXTDET2) & "','" & Trim(TXTDET3) & _
                "','" & Trim(TXTDET4) & "','" & Trim(TXTDET5) & "','A','A')"
    
    CN.CommitTrans
      
If SAVEFLAG = False Then
   MsgBox "Data Successfully Saved"
Else
   MsgBox "Data Successfully Edited"
End If

Call RESETALL
Call cmdCancel_Click

Exit Sub
LAST:
  ErrNumber = Err.Number
  ErrMessage = Err.Description
  frm_ErrorHandler.Show vbModal
End Sub

Private Sub RESETALL()
 txtName = Empty
 txtName.SetFocus
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long

    If Trim(TXTDET1) = Empty Then
        CheckData = True
        TXTDET1.SetFocus
        Exit Function
    End If
End Function

Private Sub TXTEXP_GotFocus()
    TXTEXP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTEXP_LostFocus()
   TXTEXP.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
  txtName.BackColor = RGB(BRED, BGREEN, BBLUE)
  txtName.SelStart = 0
  txtName.SelLength = Len(txtName)
  txtName.ToolTipText = "Enter Party Name"
End Sub

Private Sub TXTNAME_LostFocus()
   txtName.BackColor = vbWhite
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelETE.Enabled = Not bool
    
    txtName.Enabled = Not bool
    TXTJURI.Enabled = Not bool
    TXTNOTI.Enabled = Not bool
    TXTEXP.Enabled = Not bool
    CMBTYP.Enabled = Not bool
End Sub

Private Function entry(CODE As String)
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
Dim ROW As Long

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "Select * from BILLTERMS where CODE ='" & CODE & "' AND STATUS='A' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If TEMPRS.EOF = False Then
 SWITCH = True
 txtName = Trim(TEMPRS!NAME & "")
 TXTJURI = Trim(TEMPRS!JURI & "")
 TXTNOTI = Trim(TEMPRS!NOTI & "")
 TXTEXP = Trim(TEMPRS!Exp & "")
 
 TXTDET1 = Trim(TEMPRS!DET1 & "")
 TXTDET2 = Trim(TEMPRS!DET2 & "")
 TXTDET3 = Trim(TEMPRS!DET3 & "")
 TXTDET4 = Trim(TEMPRS!DET4 & "")
 TXTDET5 = Trim(TEMPRS!DET5 & "")
 CMBTYP.Text = Trim(TEMPRS!Type & "")
End If
TEMPRS.Close
  
End Function

Private Sub TXTDET1_GotFocus(): TXTDET1.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTDET1_LostFocus(): TXTDET1.BackColor = vbWhite: End Sub

Private Sub TXTDET2_GotFocus(): TXTDET2.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTDET2_LostFocus(): TXTDET2.BackColor = vbWhite: End Sub

Private Sub TXTDET3_GotFocus(): TXTDET3.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTDET3_LostFocus(): TXTDET3.BackColor = vbWhite: End Sub

Private Sub TXTDET4_GotFocus(): TXTDET4.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTDET4_LostFocus(): TXTDET4.BackColor = vbWhite: End Sub

Private Sub TXTDET5_GotFocus(): TXTDET5.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTDET5_LostFocus(): TXTDET5.BackColor = vbWhite: End Sub

VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_schedulemaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule Master"
   ClientHeight    =   4215
   ClientLeft      =   3360
   ClientTop       =   3930
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7005
   Begin VB.Frame FramHead 
      Height          =   720
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6825
      Begin VB.Label lblHead 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SCHEDULE MASTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame FramCont 
      Height          =   2355
      Left            =   120
      TabIndex        =   1
      Top             =   855
      Width           =   6825
      Begin VB.TextBox txtseqc 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtschno 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   5
         ToolTipText     =   "Type Group Description"
         Top             =   600
         Width           =   4395
      End
      Begin VB.ComboBox schgrp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_schedulemaster.frx":0000
         Left            =   1680
         List            =   "frm_schedulemaster.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select Category of Group"
         Top             =   960
         Width           =   4410
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   3
         Top             =   180
         Width           =   1215
      End
      Begin VB.ComboBox cmbSCAT 
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3720
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1845
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule No."
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
         Left            =   120
         TabIndex        =   8
         Top             =   1485
         Width           =   1170
      End
      Begin VB.Label lblDESc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Name"
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lblCat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Gr"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblSALCAT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Type of Group:"
         Height          =   195
         Left            =   3480
         TabIndex        =   19
         Top             =   3720
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code :"
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
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   570
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   3255
      Width           =   6855
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_schedulemaster.frx":003A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelETE 
         Height          =   495
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_schedulemaster.frx":03D4
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_schedulemaster.frx":076E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_schedulemaster.frx":14F8
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5520
         TabIndex        =   16
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_schedulemaster.frx":194A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_schedulemaster.frx":1D9C
         cBack           =   -2147483633
      End
   End
End
Attribute VB_Name = "frm_schedulemaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean, Category As String
Dim RECST As New ADODB.Recordset
Private Sub cmdAdd_Click()
    cmdCancel.Cancel = True
    Call btn_sts(False)
    Call ClsData
    txtCode.Text = GenHCODE1("Select Max(HCOD) From HEDMST", "HCOD")
    txtCode.SetFocus
    SAVEFLAG = True
End Sub

Private Sub cmdCancel_Click()
    Call btn_sts(True)
    Call ClsData
    cmdAdd.SetFocus
    cmdExit.Cancel = True
End Sub

Private Sub cmdDelete_Click()
    Dim ANS As String, TEMPRS As New ADODB.Recordset

    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("000010", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    Set TEMPRS = New ADODB.Recordset
    If TEMPRS.State = 1 Then TEMPRS.Close
     
    TEMPRS.Open "SElect CODE from GRPMST where HCOD ='" & Trim(txtCode.Text) & "'", CN, adOpenDynamic, adLockOptimistic
    
    If TEMPRS.EOF = False Then
        MsgBox "Further Entry Exist Can not delete record.", vbCritical, App.Title
        Exit Sub
    Else
        ANS = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, App.Title)
        If ANS = vbYes Then
            CN.Execute ("Delete from HEDMST where HCOD ='" & Trim(txtCode.Text) & "'")
        End If
    End If
    Call btn_sts(True)
    Call ClsData
End Sub

Private Sub cmdEdit_Click()
    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("000010", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    txtCode = Empty
    SAVEFLAG = False
    M_DESC = Empty
    Key = Empty
    txtCode.Text = SearchList("Select HCOD AS CODE ,[NAME] from HEDMST")
    
    
    Call btn_sts(False)
    txtCode.Enabled = False
    
    Call FILLDATA
    
    
    txtDesc.SetFocus
End Sub

Private Sub cmdExit_Click()
    Msg ""
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdList_Click()
    If SAVEFLAG Then Exit Sub
    txtCode = Empty
    SAVEFLAG = False
    M_DESC = Empty
    Key = Empty
    txtCode.Text = SearchList("Select CODE ,[NAME] from IGMMST")
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveRec
    Dim SQL As String
    Dim cat As String, TEMPRS As New ADODB.Recordset
    Dim cprq, mrrq, ltrq, adrq, dgrq As String
    Dim Ctrl As Control

    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
        
    If Trim(txtCode.Text) = "" Then
        MsgBox "Code Can not be empty.", vbCritical, App.Title
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtDesc.Text = "" Then
        MsgBox "Description Can not be empty.", vbCritical, App.Title
        txtDesc.SetFocus
        Exit Sub
    End If
    
    If schgrp.ListIndex = -1 Then
        schgrp.SetFocus
        MsgBox "Schedule Group Should Not Be Empty.....", vbInformation, App.Title
        Exit Sub
    End If
    
    
    If Trim(txtschno) = Empty Then
        txtschno.SetFocus
        MsgBox "Schedule Group can not be empty.", vbCritical, App.Title
        Exit Sub
    End If
    
    If Not IsNumeric(txtseqc) Then
        txtseqc.SetFocus
        MsgBox "Sequence No Should be numeric", vbCritical, App.Title
        Exit Sub
    End If
    
    Dim M_DRCR As String
    Dim SCH6 As String
    M_DRCR = "*"
    
    Select Case schgrp.Text
     Case "Shareholder fund"
      M_DRCR = "C"
      SCH6 = "1"
     Case "Loan Funds"
      M_DRCR = "C"
      SCH6 = "1"
     Case "Deferred Tax Liabilites"
      M_DRCR = "C"
      SCH6 = "1"
     Case "Fixed Assets"
      M_DRCR = "D"
      SCH6 = "2"
     Case "Investments"
      M_DRCR = "D"
      SCH6 = "2"
     Case "Current Assets,Loans and advances"
      SCH6 = "2"
      M_DRCR = "D"
     Case "Current Liabilites and provision"
      M_DRCR = "C"
      SCH6 = "2"
     Case "Miscellaneous expenditure"
      M_DRCR = "C"
      SCH6 = "2"
     Case "Income"
      M_DRCR = "C"
      SCH6 = "3"
     Case "Expenditure"
      M_DRCR = "D"
      SCH6 = "4"
    End Select
    
    
    
     
     
    
    
            
    If SAVEFLAG = True Then
        TEMPRS.Open "Select NAME from HEDMST where [NAME] ='" & Trim(txtDesc.Text) & "'", CN, adOpenDynamic, adLockOptimistic
        If TEMPRS.EOF = False Then
            MsgBox "Can not insert Duplicate Name.", vbCritical, App.Title
            TEMPRS.Close
            txtDesc.SetFocus
            Exit Sub
        End If
        TEMPRS.Close
        
        TEMPRS.Open "Select HCOD from HEDMST where HCOD ='" & Trim(txtCode.Text) & "'", CN, adOpenDynamic, adLockOptimistic
        If TEMPRS.EOF = False Then
            MsgBox "Can not insert Duplicate Code.", vbCritical, App.Title
            TEMPRS.Close
            txtCode.SetFocus
            Exit Sub
        End If
        TEMPRS.Close
        
        SQL = "insert into HEDMST (HCOD,[NAME],DRCR,SCH6,SCH_GRP,SCH_NOS,SCH_SEQ) " _
        & " values('" & Trim(txtCode.Text) & "','" & Trim(txtDesc.Text) & "','" & M_DRCR & "','" & SCH6 & "' " & _
        ",'" & schgrp.Text & "','" & txtschno & "','" & txtseqc & "')"

               
                  
        CN.BeginTrans
            CN.Execute SQL
        CN.CommitTrans
        
    Else
        CN.BeginTrans
            CN.Execute "UPDATE HEDMST SET [NAME] ='" & txtDesc.Text & "',DRCR='" & M_DRCR & "',SCH6='" & SCH6 & "',SCH_GRP='" & schgrp & "',SCH_NOS='" & txtschno & "',SCH_SEQ='" & txtseqc & "' WHERE HCOD = '" & txtCode & "'"
            
            'DAILYSTAT
        CN.CommitTrans
    End If
    Call btn_sts(True)
    Call ClsData
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub
    
errSaveRec:
Resume
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
On Error GoTo errLoad
    Call CenterChild(frm_Main, Me)
    Me.KeyPreview = True
    schgrp.Clear
    schgrp.AddItem "Shareholder fund"
    schgrp.AddItem "Loan Funds"
    schgrp.AddItem "Deferred Tax Liabilites"
    schgrp.AddItem "Fixed Assets"
    schgrp.AddItem "Investments"
    schgrp.AddItem "Current Assets,Loans and advances"
    schgrp.AddItem "Current Liabilites and provision"
    schgrp.AddItem "Miscellaneous expenditure"
    schgrp.AddItem "Income"
    schgrp.AddItem "Expenditure"
    schgrp.ListIndex = 0
    
    
    txtDesc.Enabled = False
    
        
    
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelETE.Enabled = False
    cmdExit.Cancel = True
    
    Exit Sub

errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(boo As Boolean)
    txtCode.Enabled = Not boo
    txtDesc.Enabled = Not boo
    
        
    cmbSCAT.Enabled = Not boo
    
    
    cmdSave.Enabled = Not boo
    cmdCancel.Enabled = Not boo
    cmdDelETE.Enabled = Not boo
    cmdAdd.Enabled = boo
    cmdEdit.Enabled = boo
End Sub

Private Sub ClsData()
    txtCode.Text = ""
    txtDesc.Text = ""
    txtschno.Text = ""
    txtseqc.Text = ""
End Sub
Private Sub txtCode_GotFocus()
txtCode.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtCode_LostFocus()
txtCode.BackColor = vbWhite
End Sub
Private Sub txtDesc_GotFocus()
    txtDesc.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Please Enter Item Group Name"
End Sub
Private Sub txtDesc_LostFocus()
 txtDesc.BackColor = vbWhite
End Sub
Private Sub txtseqc_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, txtseqc, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub FILLDATA()
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open "SELECT NAME,SCH_GRP,SCH_NOS,SCH_SEQ  FROM HEDMST WHERE HCOD='" & txtCode & "'", CN, adOpenDynamic, adLockOptimistic
  If Not EDTDAT.EOF Then
    txtDesc = EDTDAT!NAME & ""
    txtschno = EDTDAT!SCH_NOS & ""

    txtseqc = EDTDAT!SCH_SEQ & ""
     Select Case Trim(EDTDAT!SCH_GRP)
     Case "Shareholder fund"
      schgrp.ListIndex = 0
     Case "Loan Funds"
      schgrp.ListIndex = 1
     Case "Deferred Tax Liabilites"
      schgrp.ListIndex = 2
     Case "Fixed Assets"
      schgrp.ListIndex = 3
     Case "Investments"
      schgrp.ListIndex = 4
     Case "Current Assets,Loans and advances"
      schgrp.ListIndex = 5
     Case "Current Liabilites and provision"
      schgrp.ListIndex = 6
     Case "Miscellaneous expenditure"
      schgrp.ListIndex = 7
     Case "Income"
      schgrp.ListIndex = 8
     Case "Expenditure"
      schgrp.ListIndex = 9
     Case Else
      schgrp.ListIndex = 0
    End Select
   Else
    txtschno = Empty
    
    txtseqc = Empty
    schgrp.ListIndex = 0

  End If
  EDTDAT.Close
End Sub

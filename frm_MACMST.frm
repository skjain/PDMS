VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_MACMST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Master"
   ClientHeight    =   3225
   ClientLeft      =   3585
   ClientTop       =   2625
   ClientWidth     =   8100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8100
   Begin VB.Frame FramCmd 
      Height          =   900
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   7935
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
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
         Image           =   "frm_MACMST.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4152
         TabIndex        =   12
         Top             =   240
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
         Image           =   "frm_MACMST.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   5496
         TabIndex        =   13
         Top             =   240
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
         Image           =   "frm_MACMST.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1464
         TabIndex        =   10
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
         Image           =   "frm_MACMST.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2808
         TabIndex        =   11
         Top             =   240
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
         Image           =   "frm_MACMST.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6840
         TabIndex        =   14
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
         Image           =   "frm_MACMST.frx":1CAA
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      Begin VB.TextBox M_SPDL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox M_DVNM 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox M_NAME 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin MSMask.MaskEdBox M_BOXN 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "A999999/99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Last Box No."
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
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "No. of Spindles"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Division"
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
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Machine Name"
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
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_MACMST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim M_DVCD As String
Dim M_CODE As String
Private Sub cmdAdd_Click()
    Call ClsData
    Call btn_sts(False)
    M_NAME.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData
End Sub

Private Sub cmdDelete_Click()
    'Check for Delete
    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("000016", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    If RS.State = 1 Then RS.Close
    
    RS.Open "Select * From BOXREG Where COMP='" & compPth & "' and MCCD='" & M_CODE & "'", CN
    
    If Not RS.EOF Then
        MsgBox "Record With This Machine is already Exists !! Can't Delete !!", vbInformation, "Access Denied !!"
        RS.Close
        Exit Sub
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "Select * From SPTRAN Where COMP='" & compPth & "' and PCOD='" & M_CODE & "'", CN
    If Not RS.EOF Then
        MsgBox "Record With This Machine is already Exists !! Can't Delete !!", vbInformation, "Access Denied !!"
        RS.Close
        Exit Sub
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "Select * From PURTRAN Where COMP='" & compPth & "' and PCOD='" & M_CODE & "'", CN
    If Not RS.EOF Then
        MsgBox "Record With This Machine is already Exists !! Can't Delete !!", vbInformation, "Access Denied !!"
        RS.Close
        Exit Sub
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "Select * From STORETRAN Where COMP='" & compPth & "' and PCOD='" & M_CODE & "'", CN
    If Not RS.EOF Then
        MsgBox "Record With This Machine is already Exists !! Can't Delete !!", vbInformation, "Access Denied !!"
        RS.Close
        Exit Sub
    End If
    If M_CODE = "" Then
      Exit Sub
    End If
    
    Dim ays
    
    ays = MsgBox("Are You Sure ? Want to Delete This Machine ?", vbYesNo + vbQuestion, "Are You Sure ?")
    
    If ays = vbYes Then
      CN.BeginTrans
           CN.Execute "delete from macmst where code='" & M_CODE & "' AND COMP='" & compPth & "' AND DVCD='" & M_DVCD & "' AND UNIT='" & UNCD & "'"
           
            Call DAILYSTATUS("TPT", M_CODE, "", 0, "", 0, cUName, "D", Now, Now)
       CN.CommitTrans
    End If
    
    Call cmdCancel_Click
    
    cmdAdd.SetFocus
End Sub

Private Sub cmdEdit_Click()
On Error GoTo errLoadData
  If M_USRSECLEVL = 1 Then
      If ReadConfigMaster("000016", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  M_DESC = Empty
  Key = Empty
  'M_NAME.Text = SearchList1("SELECT TOP 20 Code,[NAME] from MACMST", 0, "", "Machine Master List")
  'M_CODE = Key
  LOAD frm_MacMstLst
  
  frm_MacMstLst.Show 1
  
  If M_NAME = Empty Then Exit Sub
  
  If RS.State = 1 Then RS.Close
  
  RS.Open "select MACMST.* from MACMST INNER JOIN DIVMST ON DIVMST.CODE=MACMST.DVCD where MACMST.Code='" & M_NAME.Tag & "' AND DIVMST.NAME='" & M_DVNM & "' AND MACMST.UNIT='" & UNCD & "'", CN, adOpenKeyset, adLockPessimistic
  
  If RS.EOF Then
    MsgBox "Machine Does Not Exist ", vbInformation
    Exit Sub
  End If
  M_CODE = RS!CODE
  M_DVCD = RS!DVCD
  M_NAME.Text = Trim(RS!NAME)
  M_SPDL.Text = Trim(RS!spdl)
  If Len(Trim(RS!BOXN)) = 10 Then M_BOXN.Text = Trim(RS!BOXN)
  
  M_DVCD = RS!DVCD & ""
  
  btn_sts (False)
  M_NAME.SetFocus
  
  Exit Sub
  
errLoadData:
Resume
  ErrNumber = Err.Number
  ErrMessage = Err.Description
  frm_ErrorHandler.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveRec

    If RS.State = 1 Then RS.Close
    
    If Trim(M_NAME) = Empty Then
        MsgBox "Please enter valid Machine Name !!", vbInformation
        M_NAME.SetFocus
        Exit Sub
    End If
    
    If M_DVNM = Empty Then
        MsgBox "Please Select Division From List !!", vbInformation
        M_DVNM.SetFocus
        Call M_DVNM_KeyDown(13, 0)
        Exit Sub
    End If
    
    RS.Open "SELECT * FROM MACMST WHERE NAME='" & M_NAME.Text & "' AND DVCD='" & M_DVCD & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockPessimistic
    
    If Not RS.EOF Then
        If RS!CODE = M_CODE Then
            'Nothing To Do
        Else
            MsgBox "Duplicate Name Not Allowed", vbInformation
            M_NAME.SetFocus
            Exit Sub
        End If
    End If
    
    If RS.State = 1 Then RS.Close
    
    RS.Open "SELECT * FROM DIVMST WHERE NAME='" & M_DVNM.Text & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
    
    If RS.EOF Then
        MsgBox "Division not exist", vbInformation
        M_DVNM.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(M_SPDL) Then
        MsgBox "Spindles should be number", vbInformation
        M_SPDL.SetFocus
        Exit Sub
    End If
    
    If M_CODE = Empty Or M_CODE = "" Then
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT ISNULL(MAX(CODE),000000) AS COD1 FROM MACMST WHERE DVCD='" & M_DVCD & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
    Dim COD1
        COD1 = Val(RS!COD1) + 1
        If COD1 <= 9 Then
            M_CODE = "00000" + Trim(STR(COD1))
        End If
        
        If COD1 > 9 And COD1 <= 99 Then
            M_CODE = "0000" + Trim(STR(COD1))
        End If
        
        If COD1 > 99 And COD1 <= 999 Then
            M_CODE = "000" + Trim(STR(COD1))
        End If
        
        If COD1 > 999 And COD1 <= 9999 Then
            M_CODE = "00" + Trim(STR(COD1))
        End If
        
        If COD1 > 9999 And COD1 <= 99999 Then
            M_CODE = "0" + Trim(STR(COD1))
        End If
        
        If COD1 > 99999 Then
            M_CODE = Trim(STR(COD1))
        End If
    End If
    
    If RS.State = 1 Then RS.Close
    
    RS.Open "SELECT * FROM MACMST WHERE CODE='" & M_CODE & "' AND DVCD='" & M_DVCD & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenKeyset, adLockPessimistic
    
    CN.BeginTrans
        If RS.EOF Then
            RS.AddNew
            
             Call DAILYSTATUS("MAC", M_CODE, "", 0, "", 0, cUName, "N", Now, Now)
        Else
            
             Call DAILYSTATUS("MAC", M_CODE, "", 0, "", 0, cUName, "M", Now, Now)
        End If
        
        If Trim(M_BOXN) = "_______/__" Then M_BOXN.Mask = "": M_BOXN.Text = Empty: M_BOXN = Empty
        RS!COMP = compPth
        RS!CODE = M_CODE
        RS!NAME = Trim(M_NAME)
        RS!DVCD = Trim(M_DVCD)
        RS!spdl = Val(M_SPDL)
        RS!BOXN = Trim(M_BOXN)
        RS!unit = Trim(UNCD)
        RS.Update
    CN.CommitTrans
    
    Call cmdCancel_Click
    
    cmdAdd.SetFocus

    Exit Sub
    
errSaveRec:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
    On Error Resume Next
    CN.RollbackTrans
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If ActiveControl.NAME = "M_DVNM" Then Exit Sub
   If KeyAscii = vbKeyReturn Then
     SendKeys "{TAB}"
   End If
End Sub
Private Sub Form_Load()
Call ColorComponent(Me)
On Error GoTo errLoad
    CMDSAVE.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    M_NAME.Enabled = False
    M_DVNM.Enabled = False
    M_SPDL.Enabled = False
    M_BOXN.Enabled = False
    Call CenterChild(frm_Main, Me)
    cmdExit.Cancel = True
    Me.KeyPreview = True
  Exit Sub

errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
    
End Sub
Private Sub btn_sts(bool As Boolean)
    CMDSAVE.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    M_NAME.Enabled = Not bool
    M_DVNM.Enabled = Not bool
    M_SPDL.Enabled = Not bool
End Sub
Private Sub ClsData()
    M_NAME.Text = ""
    M_DVNM.Text = ""
    M_SPDL.Text = ""
    M_BOXN.Mask = ""
    M_BOXN.Text = ""
    M_BOXN.Mask = "A999999/99"
    M_DVCD = Empty
    M_CODE = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    For Each LastFrm In Forms
        If LastFrm.NAME = "frmRefStatus" Then
            frmRefStatus.ZOrder
            Exit For
        End If
    Next

End Sub

Private Sub M_BOXN_GotFocus()
M_BOXN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_BOXN_LostFocus()
M_BOXN.BackColor = vbWhite
End Sub

Private Sub M_BOXN_Validate(Cancel As Boolean)
    If M_BOXN = "_______/__" Then Exit Sub
    If Len(M_BOXN) <> M_BOXN.MaxLength Or InStr(1, M_BOXN, "_") > 0 Then
        MsgBox "Please Enter " & M_BOXN.MaxLength & " Code Value ", vbInformation
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub M_DVNM_GotFocus()
M_DVNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_DVNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 And M_DVNM = Empty) Or KeyCode = vbKeyF2 Then
        M_DESC = Empty
        Key = Empty
        M_DVNM.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A' AND CODE<>'000001'", 0, "", "DIVISION MASTER HELP")
        M_DVCD = Key
        If M_DVNM.Text <> "" Then
            M_SPDL.SetFocus
        Else
            M_DVNM.SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn Then
            SendKeys "{TAB}"
    End If

    If KeyCode = vbKeyDelete Then
        M_DVNM.Text = ""
    End If
End Sub

Private Sub M_DVNM_LostFocus()
M_DVNM.BackColor = vbWhite
End Sub

Private Sub M_NAME_GotFocus()
M_NAME.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_NAME_LostFocus()
M_NAME.BackColor = vbWhite
End Sub

Private Sub M_SPDL_GotFocus()
M_SPDL.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_SPDL_KeyPress(KeyAscii As Integer)
    Call CheckNumericKey(KeyAscii, M_SPDL, Me)
    If KeyAscii = 46 Then KeyAscii = 0
End Sub

Private Sub M_SPDL_LostFocus()
M_SPDL.BackColor = vbWhite
End Sub

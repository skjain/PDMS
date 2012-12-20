VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmPlyMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ply Master"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6960
   Begin VB.TextBox M_NAME 
      Height          =   285
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   7
      Top             =   480
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   6735
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3360
         TabIndex        =   3
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
         Image           =   "frmPlyMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4440
         TabIndex        =   4
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
         Image           =   "frmPlyMaster.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1200
         TabIndex        =   1
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
         Image           =   "frmPlyMaster.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2280
         TabIndex        =   2
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
         Image           =   "frmPlyMaster.frx":14BE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5520
         TabIndex        =   5
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
         Image           =   "frmPlyMaster.frx":1910
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
         Image           =   "frmPlyMaster.frx":1D62
         cBack           =   -2147483633
      End
   End
   Begin VB.Label Label1 
      Caption         =   "PLY NAME :"
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
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmPlyMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Public M_CODE As String

Private Sub cmdAdd_Click()
Me.Caption = "HELLO"
Me.Caption = "99999999999999O"
Me.Caption = "HELLO121332212223"

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
If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000022", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  If M_CODE = "" Then
     Exit Sub
  End If
    
    Dim AYS
    
    AYS = MsgBox("Are You Sure ? Want to Delete This Ply Master ?", vbYesNo + vbQuestion, "Are You Sure ?")
    
    If AYS = vbYes Then
      CN.BeginTrans
      CN.Execute "DELETE FROM PLYMST WHERE CODE='" & M_CODE & "' AND RECSTAT='A'"
      'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','EXP','XXXXXXXXXXXXX','" & M_NAME & "',NULL,'" & M_CODE & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
      '-------------------------------
      'DAILYSTATUS
      Call DAILYSTATUS("PLM", M_CODE, "", 0, "", 0, cUName, "D", Now, Now)
      '-------------------------------
      CN.CommitTrans
    End If
    
    Call cmdCancel_Click
    cmdAdd.SetFocus
End Sub

Private Sub cmdEdit_Click()

  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000022", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
On Error GoTo errLoadData
  SAVEFLAG = False
  NEW_VISIBLE = False
  Key = Empty
  M_DESC = Empty
  
  M_NAME = SearchList1("SELECT DISTINCT CODE, NAME FROM PLYMST WHERE RECSTAT='A'", 0, "", "List of Ply from Master")
  M_CODE = Key
  
  If M_CODE <> Empty Then
     'Call FILLFLEX
  End If
  
  If M_NAME.Enabled = True Then M_NAME.SetFocus

  If M_CODE = Empty Then Exit Sub
  
  btn_sts (False)
  M_NAME.SetFocus
  Exit Sub
  
errLoadData:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveRec

    If RS.State = 1 Then RS.Close
    
    If Trim(M_NAME) = Empty Or InStr(1, M_NAME, " ") <> 0 Then
        MsgBox "Please Enter valid Ply Name !!", vbInformation
        M_NAME.SetFocus
        Exit Sub
    ElseIf UCase(Trim(M_NAME)) = "TOP" Or UCase(Trim(M_NAME)) = "BOTTOM" Then
        MsgBox "ENTER OTHER THAN 'TOP' AND 'BOTTOM' NAME", vbInformation
        M_NAME.SetFocus
        Exit Sub
    End If
   
    Call AddColumn
   
    If SAVEFLAG Then
       M_CODE = GENSIXCOD("SELECT ISNULL(MAX(CODE),'000000') AS CODE FROM PLYMST WHERE CODE LIKE '0%'")
    End If
      
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM PLYMST WHERE NAME='" & M_NAME & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockPessimistic
    If Not RS.EOF Then
        If RS!CODE = M_CODE Then
            'Nothing To Do
        Else
            MsgBox "Duplicate Name For Export Type Master.", vbInformation
            M_NAME.SetFocus
            Exit Sub
        End If
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM PLYMST WHERE CODE='" & M_CODE & "'", CN, adOpenKeyset, adLockPessimistic
    CN.BeginTrans
    If RS.EOF Then
       RS.AddNew
       
       Call DAILYSTATUS("PLM", M_CODE, "", 0, "", 0, cUName, "N", Now, Now)
    Else
       Call DAILYSTATUS("PLM", M_CODE, "", 0, "", 0, cUName, "M", Now, Now)
       
    End If
        
        RS!CODE = M_CODE
        RS!NAME = M_NAME
        RS!RECSTAT = "A"
        RS.Update
        
    CN.CommitTrans
    
    Call cmdCancel_Click
    
    cmdAdd.SetFocus

    Exit Sub
    
errSaveRec:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
    On Error Resume Next
    CN.RollbackTrans
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
On Error GoTo errLoad
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    M_NAME.Enabled = False
    Call CenterChild(frm_Main, Me)
    cmdExit.Cancel = True
    Me.KeyPreview = True
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    M_NAME.Enabled = Not bool
End Sub

Private Sub ClsData()
    M_NAME.Text = ""
End Sub

Private Sub M_NAME_GotFocus()
M_NAME.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_NAME_KeyPress(KeyAscii As Integer)

If Trim(M_NAME) = Empty Then
Select Case KeyAscii
  Case vbKey0 To vbKey9
       KeyAscii = 0
End Select
End If

End Sub

Private Sub M_NAME_LostFocus()
M_NAME.BackColor = vbWhite
End Sub

Private Sub AddColumn()
On Error Resume Next
  CN.Execute "ALTER TABLE BOXREGISTER ADD " & M_NAME & " DECIMAL(9) NOT NULL DEFAULT 0"
  CN.Execute "ALTER TABLE PKGMAN ADD " & M_NAME & " DECIMAL(9) NOT NULL DEFAULT 0"
  CN.Execute "ALTER TABLE PKGSTK ADD " & M_NAME & " DECIMAL(9) NOT NULL DEFAULT 0"
End Sub

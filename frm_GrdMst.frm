VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FRM_GRDMST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grade Master (Finish Item)"
   ClientHeight    =   3345
   ClientLeft      =   1935
   ClientTop       =   3195
   ClientWidth     =   8265
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8265
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   7815
      Begin VB.CheckBox chkdwn 
         Caption         =   "Check if Down Grade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox M_NAME 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox M_SEQC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   " (Reporting Purpose) "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Grade Name     :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Sequence No.   :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2640
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
      Image           =   "frm_GrdMst.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   2640
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
      Image           =   "frm_GrdMst.frx":039A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   2640
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
      Image           =   "frm_GrdMst.frx":0734
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2640
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
      Image           =   "frm_GrdMst.frx":0ACE
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
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
      Image           =   "frm_GrdMst.frx":1858
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   2640
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
      Image           =   "frm_GrdMst.frx":1CAA
      cBack           =   -2147483633
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "GRADE MASTER "
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
      TabIndex        =   13
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   8160
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   8160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   3135
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "FRM_GRDMST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
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
    cmdAdd.SetFocus
End Sub
Private Sub cmdDelete_Click()
 If M_USRSECLEVL = 1 Then
   If ReadConfigMaster("000008", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
 End If
 
 If isFurtherEntryExist("GRADE", M_CODE) Then
     MsgBox "Further Entry Exist"
     Call cmdCancel_Click
     cmdAdd.SetFocus
     Exit Sub
 End If
 
 'Check for Delete
 If M_CODE = "" Then Exit Sub
 Dim AYS
 AYS = MsgBox("Are You Sure to Delete ", vbYesNo)
 If AYS = vbYes Then
   CN.BeginTrans
   CN.Execute "delete from grdmst where code='" & M_CODE & "'"
   
   '------------------------------------------------------------
   'DAILYSTAT
    Call DAILYSTATUS("GRD", M_CODE, "", 0, "", 0, cUName, "D", Now, Now)
   '-----------------------------------------------------------
   
   CN.CommitTrans
 End If
 Call cmdCancel_Click
 cmdAdd.SetFocus
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = 1 Then
     If ReadConfigMaster("000008", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  btn_sts (False)
  
  LOAD frm_GRDList
  frm_GRDList.Show 1
  
  If M_SEQC = Empty Then Call cmdCancel_Click: Exit Sub
  
  RS.Open "select * from grdmst where SEQC='" & M_SEQC & "'", CN, adOpenKeyset, adLockPessimistic
  
  If RS.EOF Then
    MsgBox "Error Occured While Loading Grade Detail", vbInformation
    Exit Sub
  End If
  
  M_NAME.Text = Trim(RS!grad)
  M_SEQC.Text = Trim(RS!SEQC)
  If RS!EXTRA1 & "" = "Yes" Then
    chkdwn.Value = 1
   Else
    chkdwn.Value = 0
  End If
  M_CODE = RS!CODE
  
  If RS.State = adStateOpen Then RS.Close
  
  M_NAME.SetFocus
  
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveRec
Dim Ctrl As Control

Dim m_chkdwngrd As String

  For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
  Next
   
  If chkdwn.Value = 1 Then
     m_chkdwngrd = "Yes"
  Else
     m_chkdwngrd = "No"
  End If
    
   If Trim(Me.M_NAME) = Empty Then
        MsgBox "Please Enter Valid Grade Name !!", vbInformation
        M_NAME.SetFocus
        Exit Sub
   End If
    
   If RS.State = 1 Then RS.Close
   RS.Open "SELECT * FROM GRDMST WHERE GRAD='" & M_NAME.Text & "'", CN, adOpenKeyset, adLockPessimistic
        
    'For Duplicate Name check
   If Not RS.EOF Then
        If RS!CODE = M_CODE Then
            'Nothing To Do
        Else
            MsgBox "Duplicate Grade Not Allowed", vbInformation
            M_NAME.SetFocus
            Exit Sub
        End If
    End If
   
    'For Sequence Duplicate Check
    If RS.State = 1 Then RS.Close
   
    RS.Open "SELECT * FROM GRDMST WHERE SEQC='" & Val(M_SEQC.Text) & "'", CN, adOpenKeyset, adLockPessimistic
    If Not RS.EOF Then
        If RS!CODE = M_CODE Then
            'Nothing To Do
        Else
            MsgBox "Duplicate Sequence Not Allowed", vbInformation
            M_SEQC.SetFocus
            Exit Sub
        End If
    End If
   
    If Not IsNumeric(M_SEQC) Then
        MsgBox "Sequence should be number", vbInformation
        M_SEQC.SetFocus
        Exit Sub
    End If
   
    If M_CODE = Empty Or M_CODE = "" Then
        If RS.State = 1 Then RS.Close
        
        RS.Open "SELECT ISNULL(MAX(CODE),000000) AS COD1 FROM GRDMST", CN, adOpenKeyset, adLockPessimistic
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
    
    Call SaveShadeInfo
   
   If RS.State = 1 Then RS.Close
   RS.Open "SELECT * FROM GRDMST WHERE CODE='" & M_CODE & "'", CN, adOpenKeyset, adLockPessimistic
   
   '----------------------------
   'DAILYSTATUS ENTRY
   If RS.EOF Then
     RS.AddNew
     Call DAILYSTATUS("GRD", M_CODE, "", 0, "", 0, cUName, "N", Now, Now)
    Else
      Call DAILYSTATUS("GRD", M_CODE, "", 0, "", 0, cUName, "M", Now, Now)
   End If
   
   RS!CODE = M_CODE
   RS!grad = Trim(M_NAME.Text)
   RS!SEQC = Val(M_SEQC)
   RS!EXTRA1 = m_chkdwngrd
   RS.Update
   Call cmdCancel_Click
   cmdAdd.SetFocus
   Exit Sub
errSaveRec:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
    
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
Call ColorComponent(Me)
On Error GoTo errLoad

    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    M_NAME.Enabled = False
    M_SEQC.Enabled = False
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
    M_SEQC.Enabled = Not bool
End Sub
Private Sub ClsData()
    M_NAME.Text = ""
    M_SEQC.Text = ""
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

Private Sub M_NAME_GotFocus()
M_NAME.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_NAME_LostFocus()
M_NAME.BackColor = vbWhite
End Sub

Private Sub M_SEQC_GotFocus()
M_SEQC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_SEQC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 8
            'Valid Keys
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub M_SEQC_LostFocus()
M_SEQC.BackColor = vbWhite
End Sub

Private Sub SaveShadeInfo()
On Error GoTo LAST
Dim TMPRS As ADODB.Recordset
Set TMPRS = New ADODB.Recordset

Dim INFORS As ADODB.Recordset
Set INFORS = New ADODB.Recordset

If TMPRS.State = 1 Then TMPRS.Close
TMPRS.Open "SELECT CODE,CFGTYP FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND RECSTAT='A' AND CFGTYP='GS'", CN, adOpenDynamic, adLockOptimistic
Do While Not TMPRS.EOF

   CN.Execute "DELETE FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND DVCD='" & TMPRS!CODE & "' AND GRAD= '" & M_CODE & "' "
           
   If INFORS.State = 1 Then INFORS.Close
   INFORS.Open "SELECT DISTINCT SUBGRD,NAME,SWGT,EWGT,RDIFF,SEQNO FROM SUBGRDMST " & _
               "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND  DVCD='" & TMPRS!CODE & "' AND RECSTAT='A'", CN, adOpenKeyset, adLockPessimistic
   Do While Not INFORS.EOF
           
            CN.Execute "INSERT INTO SUBGRDMST(COMP,UNIT,DVCD,GRAD,SUBGRD,NAME,SWGT,EWGT,RDIFF,SEQNO," & _
            "STATUS,RECSTAT,SUBPKGCODE) VALUES('" & compPth & "','" & UNCD & "','" & TMPRS!CODE & _
            "','" & M_CODE & "','" & INFORS!SUBGRD & "','" & INFORS!NAME & _
            "','" & INFORS!SWGT & "','" & INFORS!EWGT & "','" & INFORS!RDIFF & _
            "','" & INFORS!SEQNO & "','A','A','')"
        
   INFORS.MoveNext
   Loop
   INFORS.Close
   
TMPRS.MoveNext
Loop
TMPRS.Close

   
Exit Sub
LAST:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub


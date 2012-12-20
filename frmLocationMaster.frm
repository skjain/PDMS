VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmLocationMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2490
   ClientLeft      =   3015
   ClientTop       =   3630
   ClientWidth     =   7065
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   6735
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3360
         TabIndex        =   8
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
         Image           =   "frmLocationMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4440
         TabIndex        =   9
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
         Image           =   "frmLocationMaster.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1200
         TabIndex        =   6
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
         Image           =   "frmLocationMaster.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2280
         TabIndex        =   7
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
         Image           =   "frmLocationMaster.frx":14BE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5520
         TabIndex        =   10
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
         Image           =   "frmLocationMaster.frx":1910
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdNew 
         Height          =   495
         Left            =   120
         TabIndex        =   5
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
         Image           =   "frmLocationMaster.frx":1D62
         cBack           =   -2147483633
      End
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1395
      TabIndex        =   2
      Top             =   555
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1395
      TabIndex        =   4
      ToolTipText     =   "Enter Group Description"
      Top             =   975
      Width           =   4575
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   570
   End
   Begin VB.Label LBLTITLE 
      Alignment       =   2  'Center
      Caption         =   "Location Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6090
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1035
      Width           =   1095
   End
End
Attribute VB_Name = "frmLocationMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MB_ISEDITINGRECORD As Boolean
Dim MB_ISCREATINGNEW As Boolean
Dim M_ITEM As String
Dim rsLocation As Recordset

Private Sub cmdCancel_Click()
    Call ClsData(Me)
    Call SetControlEnabled(NO)
    MB_ISEDITINGRECORD = False
    MB_ISCREATINGNEW = False
    cmdExit.Cancel = True
    If cmdNew.Enabled = True Then
      cmdNew.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
   If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("000011", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
        
    Call cmdEdit_Click
    If txtDesc = Empty Then Exit Sub
    Call SetControlEnabled(NO)
    
    If MsgBox("Are You Sure ? Want To Delete Record ?", vbQuestion + vbYesNo, "Delete Location") = vbYes Then
        Set RS = New Recordset
        RS.Open "Select * From LOCATION Where LOCNAME='" & txtDesc & "' AND LOCID IN (SELECT DISTINCT LOCID FROM ITMMST)", CN, adOpenDynamic
        
        CN.BeginTrans
            If RS.EOF Then
                CN.Execute "DELETE FROM LOCATION WHERE LOCID='" & txtCode & "'"
               
               '-------------------------------------
               'DAILYSTATUS
                Call DAILYSTATUS("LOC", txtCode, "", 0, "", 0, cUName, "D", Now, Now)
               '--------------------------------------
            Else
               
                MsgBox "You Can Not Remove This Record !! Child Record Found !!", vbInformation, "Delete Failed"
            End If
        CN.CommitTrans
        
        RS.Close
    End If
    
    Call cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
    Dim I As Integer
    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("000011", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    MB_ISEDITINGRECORD = True
    
    Call ClsData(Me)
    
    M_DESC = Empty
    NEW_VISIBLE = False
    txtCode = Empty
    txtDesc = SearchList1("SELECT TOP 20 LOCID,LOCNAME FROM LOCATION", 0, "", "SELECT LOCATION FROM MASTER")
    txtCode = Key
    
    If txtDesc = Empty Then
        Call cmdCancel_Click
        Exit Sub
    Else
        SetControlEnabled (Yes)
        txtDesc.SetFocus
        txtCode.Enabled = False
    End If
    cmdCancel.Cancel = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    If M_USRSECLEVL <> 0 Then
        If ReadConfigMaster("0005", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
     
    Call ClsData(Me)
    
    MB_ISCREATINGNEW = True
    MB_ISEDITINGRECORD = False
    
    Call SetControlEnabled(Yes)
    
    txtCode.MaxLength = 6
    txtCode = GENCode()
    
    'txtCode.Enabled = True
    txtDesc.SetFocus
    cmdCancel.Cancel = True
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSave
    Dim MCOD As String
    Dim ACIT As String
    Dim Ctrl As Control

    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next

    If txtDesc = Empty Then
        txtDesc.SetFocus
        MsgBox "Please Enter Location Name", vbInformation, "Location Name Missing"
        Exit Sub
    End If
    
    If MB_ISCREATINGNEW Then
        txtCode = GENCode()
    End If
    
    If MB_ISCREATINGNEW Then txtCode_Validate (False)
    
    CN.BeginTrans
        If MB_ISCREATINGNEW Then
            CN.Execute "INSERT INTO LOCATION(LOCID,LOCNAME) VALUES('" & txtCode & "','" & txtDesc & "')"
            
            '--------------------------------------------------------------------
            'DAILYSTATUS
             Call DAILYSTATUS("LOC", txtCode, "", 0, "", 0, cUName, "N", Now, Now)
            '--------------------------------------------------------------------
            
        Else
            CN.Execute "UPDATE LOCATION SET LOCNAME='" & txtDesc & "' WHERE LOCID='" & txtCode & "'"
            
            '--------------------------------------------------------------------
            'DAILYSTATUS
             Call DAILYSTATUS("LOC", txtCode, "", 0, "", 0, cUName, "M", Now, Now)
            '--------------------------------------------------------------------
        End If
    CN.CommitTrans
    
    If MB_ISCREATINGNEW Then MsgBox "Code Generated : " & txtCode
    
    Call cmdCancel_Click
    
    cmdExit.Cancel = True
    MB_ISCREATINGNEW = False
    MB_ISEDITINGRECORD = False
    Exit Sub

errSave:
    If InStr(1, Err.Description, "COLUMN FOREIGN KEY constraint") > 0 Then
        CN.RollbackTrans
        ErrNumber = "000343"
        ErrMessage = txtDesc & " Location is missing in database." & vbCrLf & "Please Create / Edit Location then Try Again !!"
    Else
        ErrNumber = Err.Number
        ErrMessage = Err.Description
        CN.RollbackTrans
    End If
    
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call CenterChild(frm_Main, Me)
    Call SetControlEnabled(NO)
    
    Exit Sub

errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub SetControlEnabled(Enable As Status)
'*******************************************************************
' PROCEDURE NAME : SetControlEnabled
'
' PROCEDURE PARAMETER : -
'
' PROCEDURE RETURN VALUE :  None
'
' PROCEDURE PURPOSE :   To Enable Or Disable Content
'
' PROCEDURE LAST DATE : -
'
' PROCEDURE WRITTEN BY :Ankur
'*******************************************************************
    txtCode.Enabled = False
    txtCode.Visible = True
    txtDesc.Enabled = Enable
    cmdNew.Enabled = Not Enable
    CMDSAVE.Enabled = Enable
    cmdCancel.Enabled = Enable
    cmdEdit.Enabled = Not Enable
    cmdExit.Enabled = Not Enable
    cmdDelete.Enabled = Not Enable
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If key_PressNew Then key_PressNew = False
End Sub

Private Sub txtCode_GotFocus()
txtCode.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtCode_LostFocus()
txtCode.BackColor = vbWhite
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    Dim rsCodes As Recordset
    Dim I As Byte

    Set rsCodes = New Recordset
    rsCodes.Open "Select * From LOCATION Where LOCID='" & txtCode & "'", CN
    
    If rsCodes.EOF = False Then
        MsgBox "Code (" & txtCode & ") is in use by " & rsCodes!LOCNAME & " !! Please Check Your Code", vbInformation
        Cancel = True
    End If
    
    rsCodes.Close
End Sub

Private Sub txtDesc_GotFocus()
txtDesc.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.BackColor = vbWhite
    Call txtDesc_Validate(False)
End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
    If MB_ISEDITINGRECORD Then Exit Sub
    
    Set rsLocation = New Recordset
    txtDesc = Replace(txtDesc, "'", "", 1)
    rsLocation.Open "Select LOCNAME From LOCATION Where LOCNAME='" & txtDesc & "'", CN, adOpenDynamic, adLockOptimistic
    
    If rsLocation.EOF = False Then
        Cancel = True
        MsgBox "Location Already Exist....", vbInformation, "Location Master"
        Exit Sub
    End If
    
    rsLocation.Close
End Sub

Private Function GENCode() As String
    Dim ctr As Long

    Set RS = New Recordset
    RS.Open "Select ISNULL(MAX(LOCID),0) AS CODE From LOCATION", CN, adOpenDynamic
    
    If RS.EOF Then
        GENCode = "000001"
    Else
        ctr = Val(RS!CODE) + 1
        
        If ctr < 10 Then
            GENCode = "00000" + CStr(ctr)
        ElseIf ctr >= 10 And ctr <= 99 Then
            GENCode = "0000" + CStr(ctr)
        ElseIf ctr >= 100 And ctr <= 999 Then
            GENCode = "000" + CStr(ctr)
        ElseIf ctr >= 1000 And ctr <= 9999 Then
            GENCode = "00" + CStr(ctr)
        ElseIf ctr >= 10000 And ctr <= 99999 Then
            GENCode = "0" + CStr(ctr)
        Else
            GENCode = CStr(ctr)
        End If
    End If
    
    RS.Close
End Function

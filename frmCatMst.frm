VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmCatMst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3150
   ClientLeft      =   3015
   ClientTop       =   3630
   ClientWidth     =   6945
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   6735
      Begin WelchButton.lvButtons_H cmdNew 
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
         Image           =   "frmCatMst.frx":0000
         cBack           =   -2147483633
      End
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
         Image           =   "frmCatMst.frx":039A
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
         Image           =   "frmCatMst.frx":0734
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
         Image           =   "frmCatMst.frx":0CCE
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
         Image           =   "frmCatMst.frx":1268
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
         Image           =   "frmCatMst.frx":16BA
         cBack           =   -2147483633
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   1965
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3466
      BorderStyle     =   4
      BackColorGradient=   -2147483633
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
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         ItemData        =   "frmCatMst.frx":1C54
         Left            =   1440
         List            =   "frmCatMst.frx":1C56
         TabIndex        =   13
         Text            =   "cmbItemType"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TXTCODEGEN 
         Height          =   285
         Left            =   4995
         MaxLength       =   1
         TabIndex        =   9
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         ToolTipText     =   "Enter Group Description"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Item Type :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code Generation :"
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
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label LBLTITLE 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Item Category Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   30
         TabIndex        =   15
         Top             =   15
         Width           =   6690
      End
   End
End
Attribute VB_Name = "frmCatMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MB_ISEDITINGRECORD As Boolean
Dim MB_ISCREATINGNEW As Boolean
Dim M_ITEM As String
Option Explicit
Dim rsIGMMST As Recordset

Private Sub cmdCancel_Click()
    Call ClsData(Me)
    Call SetControlEnabled(NO)
    MB_ISEDITINGRECORD = False
    MB_ISCREATINGNEW = False
    cmdExit.Cancel = True
    cmdNew.SetFocus
End Sub

Private Sub cmddelete_Click()
    
    If M_USRSECLEVL <> 0 Then
        If ReadConfigMaster("0005", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If

    Call cmdEdit_Click
    
    If isFurtherEntryExist("ICATEGORY", txtCode) Then
       MsgBox "Further Entry Exist"
       Call cmdCancel_Click
       Exit Sub
    End If
    
    If txtDesc = Empty Then Exit Sub
    Call SetControlEnabled(NO)
    
    If MsgBox("Are You Sure ? Want To Delete Record ?", vbQuestion + vbYesNo, "Delete Item Group") = vbYes Then
        Set RS = New Recordset
        
        RS.Open "Select * From SCAT_MST Where NAME='" & txtDesc & "' AND CODE IN (SELECT DISTINCT IHCD FROM IGMMST)", CN, adOpenDynamic
        
        CN.BeginTrans
            If RS.EOF Then
                CN.Execute "DELETE FROM SCAT_MST WHERE CODE='" & txtCode & "'"
                CN.Execute "IF EXISTS(SELECT * FROM SYSOBJECTS WHERE NAME='DAILYSTAT') INSERT INTO DAILYSTAT(COMP,VTYP,PCOD,SRNO,VBNO,AMNT,CUSR,ACTN) VALUES('" & compPth & "','CAT','" & txtDesc & "','XXXXXXXXXX','XXXXXXXXXX',0,'" & cUName & "','D')"
                Call DAILYSTATUS("ICA", txtCode, "", 0, "", 0, cUName, "D", Now, Now)
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
    If M_USRSECLEVL <> 0 Then
        If ReadConfigMaster("0005", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    MB_ISEDITINGRECORD = True
    
    Call ClsData(frm_ITMIGMList)
    
    frm_ITMIGMList.FillList ("SCAT_MST")
    
    frm_ITMIGMList.Show 1
    
    If frm_ITMIGMList.Tag <> "Cancel" Then
        Me.txtCode = frmCatMst.txtCode
        Me.txtDesc = frmCatMst.txtDesc
        Me.TXTCODEGEN = frmCatMst.TXTCODEGEN
        Me.cmbItemType.ListIndex = frmCatMst.cmbItemType.ListIndex
    Else
        
    End If
    
    Unload frm_ITMIGMList
    
    If txtDesc = Empty Then
        Call cmdCancel_Click
        Exit Sub
    Else
        SetControlEnabled (Yes)
        TXTCODEGEN.SetFocus
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
    
    txtCode = "XX"
    txtCode.MaxLength = 2
    
    txtCode.Enabled = True
    TXTCODEGEN.SetFocus
    cmdCancel.Cancel = True
    cmbItemType.ListIndex = 1
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
        MsgBox "Please Enter Item Group/Sub Group Name", vbInformation, "Item Group Name Missing"
        Exit Sub
    End If
    

    If MB_ISCREATINGNEW Then
        txtCode = GENCode()
    End If
    
    If MB_ISCREATINGNEW Then txtCode_Validate (False)
    
    CN.BeginTrans
        If MB_ISCREATINGNEW Then
            CN.Execute "INSERT INTO SCAT_MST(CODE,NAME,EXTRA1,extra2) VALUES('" & txtCode & "','" & txtDesc & "','" & TXTCODEGEN.Text & "','" & IIf(cmbItemType.ListIndex = 0, "RM", "SM") & "')"
            'CN.Execute "IF EXISTS(SELECT * FROM SYSOBJECTS WHERE NAME='DAILYSTAT') INSERT INTO DAILYSTAT(COMP,VTYP,PCOD,SRNO,VBNO,AMNT,CUSR,ACTN) VALUES('" & compPth & "','CAT','" & txtDesc & "','XXXXXXXXXX','XXXXXXXXXX',0,'" & cUName & "','N')"
            
            Call DAILYSTATUS("ICA", txtCode, "", 0, "", 0, cUName, "N", Now, Now)
        Else
            CN.Execute "UPDATE SCAT_MST SET NAME='" & txtDesc & "',EXTRA1='" & TXTCODEGEN.Text & "',EXTRA2 ='" & IIf(cmbItemType.ListIndex = 0, "RM", "SM") & "' WHERE CODE='" & txtCode & "'"
            'CN.Execute "IF EXISTS(SELECT * FROM SYSOBJECTS WHERE NAME='DAILYSTAT') INSERT INTO DAILYSTAT(COMP,VTYP,PCOD,SRNO,VBNO,AMNT,CUSR,ACTN) VALUES('" & compPth & "','CAT','" & txtDesc & "','XXXXXXXXXX','XXXXXXXXXX',0,'" & cUName & "','M')"
            Call DAILYSTATUS("ICA", txtCode, "", 0, "", 0, cUName, "M", Now, Now)
        End If
    CN.CommitTrans
    
    If MB_ISCREATINGNEW Then MsgBox "Code Generated : " & txtCode
    
    Call cmdCancel_Click
    
    cmdExit.Cancel = True
    MB_ISCREATINGNEW = False
    MB_ISEDITINGRECORD = False
    Exit Sub

errSave:
    If InStr(1, ERR.Description, "COLUMN FOREIGN KEY constraint") > 0 Then
        CN.RollbackTrans
        ErrNumber = "000343"
        ErrMessage = txtDesc & " Item is missing in PDMS database." & vbCrLf & "Please Create / Edit Item Group then Try Again !!"
    Else
        ErrNumber = ERR.Number
        ErrMessage = ERR.Description
        CN.RollbackTrans
    End If
    
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmbItemType_GotFocus()
cmbItemType.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cmbItemType_LostFocus()
cmbItemType.BackColor = vbWhite
End Sub



Private Sub Form_Activate()
   Call ColorComponent(Me): Me.BackColor = RGB(RED, GREEN, BLUE)
   Me.BackColor = RGB(RED, GREEN, BLUE)
Dim I As Integer

'''    If Me.Tag = "GROUP" Then
'''        LBLTITLE.Caption = "Group Master"
'''        Set rsIGMMST = New Recordset
'''        rsIGMMST.Open "SELECT * FROM SCAT_MST ORDER BY CODE", CN
'''        cboCategory.Clear
'''
'''        Do While rsIGMMST.EOF = False
'''            cboCategory.AddItem rsIGMMST!Name
'''            rsIGMMST.MoveNext
'''        Loop
'''
'''        For I = 0 To cboCategory.ListCount - 1
'''            If cboCategory.list(I) = M_TMPPUBLIC And cboCategory.ListCount > 0 Then
'''                cboCategory.ListIndex = I
'''                Exit For
'''            End If
'''        Next
'''
'''        cboCategory.AddItem "<ADD>"
'''        If M_ITEM <> "" Then Me.cboCategory.Text = M_ITEM
'''        M_ITEM = Empty
'''        rsIGMMST.Close
'''    ElseIf Me.Tag = "SUBGRP" Then
'''        LBLTITLE.Caption = "Category Master"
'''        'cboCategory.Visible = False
'''        'lblSubCat.Visible = False
'''
'''    Else
'''        LBLTITLE.Caption = "Color Master"
'''        cboCategory.Visible = False
'''        lblSubCat.Visible = False
'''    End If
    
'''    If Me.Tag = "SUBGRP" Then
'''        With cboCategory
'''            .AddItem "Raw Item"
'''            .AddItem "Finish Item"
'''            .AddItem "Chemical & Oil"
'''            .AddItem "Others"
'''        End With
'''    End If
'''
'''    If key_PressNew Then Call cmdNew_Click
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
On Error GoTo errLoad
    Call CenterChild(frm_Main, Me)
    Call SetControlEnabled(NO)
    cmbItemType.AddItem "Raw Material"
    cmbItemType.AddItem "Store Material"
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
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
' PROCEDURE WRITTEN BY :
'*******************************************************************
    txtCode.Enabled = False
    txtCode.Visible = True
    txtDesc.Enabled = Enable
    cmdNew.Enabled = Not Enable
    cmdSave.Enabled = Enable
    cmdCancel.Enabled = Enable
    cmdEdit.Enabled = Not Enable
    cmdExit.Enabled = Not Enable
    cmdDelete.Enabled = Not Enable
    cmbItemType.Enabled = Enable
    TXTCODEGEN.Enabled = Enable
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
    
    'For I = 1 To 3 - Len(txtCode)
    '    txtCode = "0" & txtCode
    'Next
    
    rsCodes.Open "Select * From SCAT_MST Where CODE='" & txtCode & "'", CN
    
    If rsCodes.EOF = False Then
        MsgBox "Code (" & txtCode & ") is in use by " & rsCodes!NAME & " !! Please Check Your Code", vbInformation
        Cancel = True
    End If
    
    rsCodes.Close

End Sub

Private Sub TXTCODEGEN_Validate(Cancel As Boolean)
    If TXTCODEGEN.Enabled = False Then Exit Sub
    If Len(Trim(TXTCODEGEN.Text)) < 1 Then
        TXTCODEGEN.SetFocus
        MsgBox "Code Generation must be of 1 digit"
        Cancel = True
    End If
    If Not IsNumeric(TXTCODEGEN.Text) Then
        TXTCODEGEN.SetFocus
        MsgBox "Code Generation must be number only"
        Cancel = True
    End If
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
    
    Set rsIGMMST = New Recordset
    txtDesc = Replace(txtDesc, "'", "", 1)
    rsIGMMST.Open "Select Name From SCAT_MST Where Name='" & txtDesc & "'", CN, adOpenDynamic, adLockOptimistic
    
    If rsIGMMST.EOF = False Then
        Cancel = True
        MsgBox "Item Category Already Exitsts....", vbInformation, "Group Exists"
        Exit Sub
    End If
    
    rsIGMMST.Close
    
End Sub

Private Function GENCode() As String
Dim ctr As Long

    Set RS = New Recordset
    
    RS.Open "Select ISNULL(MAX(Code),0) AS CODE From SCAT_MST", CN, adOpenDynamic
    
    If RS.EOF Then
        GENCode = "01"
    Else
        ctr = Val(RS!CODE) + 1
        
        If ctr < 10 Then
            GENCode = "0" & ctr
        Else
            GENCode = ctr
        End If
    End If
    
    RS.Close
    
End Function


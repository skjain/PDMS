VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form FRM_IGRP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Group"
   ClientHeight    =   3195
   ClientLeft      =   2325
   ClientTop       =   1605
   ClientWidth     =   7080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7080
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   6855
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3360
         TabIndex        =   2
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
         Image           =   "Frm_IGrp.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelETE 
         Height          =   495
         Left            =   4440
         TabIndex        =   3
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
         Image           =   "Frm_IGrp.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1200
         TabIndex        =   9
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
         Image           =   "Frm_IGrp.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2280
         TabIndex        =   1
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
         Image           =   "Frm_IGrp.frx":14BE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5520
         TabIndex        =   4
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
         Image           =   "Frm_IGrp.frx":1910
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
         Image           =   "Frm_IGrp.frx":1D62
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame FramCont 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   6825
      Begin VB.TextBox txtItemCodeGen 
         Height          =   285
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
      Begin VB.ComboBox cmbSCAT 
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3720
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "..."
         Height          =   300
         Left            =   5985
         TabIndex        =   17
         Top             =   255
         Width           =   510
      End
      Begin VB.ComboBox cmbCat 
         Height          =   315
         ItemData        =   "Frm_IGrp.frx":20FC
         Left            =   1485
         List            =   "Frm_IGrp.frx":210C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select Category of Group"
         Top             =   1040
         Width           =   4410
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1485
         MaxLength       =   25
         TabIndex        =   6
         ToolTipText     =   "Type Group Description"
         Top             =   600
         Width           =   4395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code Generation:"
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
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   1920
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
         Left            =   375
         TabIndex        =   19
         Top             =   225
         Width           =   570
      End
      Begin VB.Label lblSALCAT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Type of Group:"
         Height          =   195
         Left            =   3480
         TabIndex        =   18
         Top             =   3720
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblCat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category: "
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
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblDESc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
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
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame FramHead 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   10
      Top             =   -15
      Width           =   6825
      Begin VB.Label lblHead 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Group"
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
         TabIndex        =   11
         Top             =   240
         Width           =   6615
      End
   End
End
Attribute VB_Name = "FRM_IGRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean, Category As String
Dim RECST As New ADODB.Recordset

Private Sub cmbCat_Change()
    Category = GetCodeERP("SCAT_MST", cmbCat, "Name", "Code")
End Sub

Private Sub cmbCat_Click()
    Category = GetCodeERP("SCAT_MST", cmbCat, "Name", "Code")
End Sub

Private Sub cmbCat_GotFocus()
    cmbCat.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Please Select Relevant Category From List"
End Sub

Private Sub cmbCat_LostFocus()
   cmbCat.BackColor = vbWhite
End Sub

Private Sub cmbSCAT_GotFocus()
    Msg "Select Relevant Group"
End Sub

Private Sub cmdAdd_Click()
    cmdCancel.Cancel = True
    Call btn_sts(False)
    Call ClsData
    txtCode.Text = GenCode1("Select Max(Code) From IGMMST", "GROUP")
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

    TEMPRS.Open "SElect * from ITMMST where IGCD ='" & Trim(txtCode.Text) & "'", CN, adOpenDynamic, adLockOptimistic
    
    If TEMPRS.EOF = False Then
        MsgBox "Can not delete record.", vbCritical, App.Title
        Exit Sub
    Else
        ANS = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, App.Title)
        If ANS = vbYes Then
            CN.Execute ("Delete from IGMMST where CODE ='" & Trim(txtCode.Text) & "'")
            'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','IGP','XXXXXXXXXXXXX','" & txtDesc & "',NULL,'" & txtCode & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
            '-------------------------------------------------------------------------
            'DAILYSTAT
            Call DAILYSTATUS("IGR", txtCode.Text, "", 0, "", 0, cUName, "D", Now, Now)
            '-------------------------------------------------------------------------
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
    txtCode.Text = SearchList("Select CODE ,[NAME] from IGMMST")
    Call btn_sts(False)
    txtCode.Enabled = False
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
    
    If cmbCat.ListIndex = -1 Then
        cmbCat.SetFocus
        MsgBox "Category Should Not Be Empty.....", vbInformation, App.Title
        Exit Sub
    End If
    
    If Len(Trim(txtItemCodeGen.Text)) < 4 Then
        txtItemCodeGen.SetFocus
        MsgBox "Item Code Generation must be of 4 digit"
        Exit Sub
    End If
    
    If Not IsNumeric(txtItemCodeGen.Text) Then
        If Mid(M_COMPBILL, 1, 3) = "CIL" Then
         Else
          txtItemCodeGen.SetFocus
          MsgBox "Item Code Generation must be number only"
          Exit Sub
        End If
    End If
            
    If SAVEFLAG = True Then
        TEMPRS.Open "Select * from IGMMST where [NAME] ='" & Trim(txtDesc.Text) & "'", CN, adOpenDynamic, adLockOptimistic
        If TEMPRS.EOF = False Then
            MsgBox "Can not insert Duplicate Name.", vbCritical, App.Title
            TEMPRS.Close
            txtDesc.SetFocus
            Exit Sub
        End If
        TEMPRS.Close
        
        TEMPRS.Open "Select * from IGMMST where CODE ='" & Trim(txtCode.Text) & "'", CN, adOpenDynamic, adLockOptimistic
        If TEMPRS.EOF = False Then
            MsgBox "Can not insert Duplicate Code.", vbCritical, App.Title
            TEMPRS.Close
            txtCode.SetFocus
            Exit Sub
        End If
        TEMPRS.Close
        
       
        SQL = "insert into IGMMST ([COMP],CODE,[NAME],CATA,SCAT,IHCD,EXTRA1) " _
        & " values('" & compPth & "','" & Trim(txtCode.Text) & "','" & Trim(txtDesc.Text) & _
        "','" & Left(cmbCat, 1) & "','" & Left(cmbSCAT.Text, 1) & "','" & Trim(Category) & _
        "','" & Trim(txtItemCodeGen.Text) & "')"
               
                  
        CN.BeginTrans
            CN.Execute SQL
            '--------------------------------------------------------------------
            'DAILYSTAT
            Call DAILYSTATUS("IGR", txtCode.Text, "", 0, "", 0, cUName, "N", Now, Now)
            '--------------------------------------------------------------------
        CN.CommitTrans
        
    Else
        CN.BeginTrans
            CN.Execute "UPDATE IGMMST SET [NAME] ='" & txtDesc.Text & "' WHERE CODE = '" & Key & "'"
            CN.Execute "UPDATE IGMMST SET SCAT ='" & Left(cmbSCAT.Text, 1) & "' WHERE CODE = '" & Key & "'"
            CN.Execute "UPDATE IGMMST SET IHCD ='" & Category & "' WHERE CODE = '" & Key & "'"
            CN.Execute "UPDATE IGMMST SET EXTRA1 ='" & Trim(txtItemCodeGen.Text) & "' WHERE CODE = '" & Key & "'"
            '-------------------------------------------------------------------------
            'DAILYSTAT
            Call DAILYSTATUS("IGR", txtCode.Text, "", 0, "", 0, cUName, "M", Now, Now)
            '-------------------------------------------------------------------------
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
    Call FillCmb("Select Name From SCAT_MST", cmbCat)
                
    txtDesc.Enabled = False
    cmbCat.Enabled = False
        
    txtItemCodeGen.Enabled = False
    CMDSAVE.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    cmbSCAT.Enabled = False
        
    cmbSCAT.AddItem "Sale"
    cmbSCAT.AddItem "Purchase"
    cmdExit.Cancel = True
    cmbSCAT.ListIndex = 0
    Exit Sub

errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(boo As Boolean)
    txtCode.Enabled = Not boo
    txtDesc.Enabled = Not boo
    cmbCat.Enabled = Not boo
        
    cmbSCAT.Enabled = Not boo
    
    txtItemCodeGen.Enabled = Not boo
    CMDSAVE.Enabled = Not boo
    cmdCancel.Enabled = Not boo
    cmdDelete.Enabled = Not boo
    cmdAdd.Enabled = boo
    cmdEdit.Enabled = boo
End Sub

Private Sub ClsData()
    txtCode.Text = ""
    txtDesc.Text = ""
    cmbCat.ListIndex = -1
    txtItemCodeGen.Text = ""
End Sub

Private Sub txtCode_Change()
    Dim TEMPRS As New ADODB.Recordset
On Error Resume Next
    If txtCode.Text = "" Then Exit Sub
    If SAVEFLAG = False Then
        TEMPRS.Open "Select * from IGMMST where CODE ='" & Trim(txtCode.Text) & "'", CN, adOpenDynamic, adLockOptimistic
        With TEMPRS
            txtDesc.Text = ![NAME]
            txtItemCodeGen.Text = Trim(!extra1 & "")
            If Not IsNull(!SCAT) Then If !SCAT = "S" Then cmbSCAT.ListIndex = 0 Else cmbSCAT.ListIndex = 1
            Select Case !CATA
                Case "R"
                    cmbCat.Text = "Raw"
                Case "F"
                    cmbCat.Text = "Finish"
                Case "C"
                    cmbCat.Text = "Chemicals & oil"
                Case "O"
                    cmbCat.Text = "Others"
            End Select
            cmbCat.Text = GetName(!ihcd, "SCAT_MST")
                        
        End With
    End If
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
    If M_COMPBILL = "SHB" Then
        If txtCode.Text = Empty Then
            txtDesc.Locked = False
        Else
            txtDesc.Locked = True
        End If
    End If
End Sub

Private Sub txtDesc_LostFocus()
 txtDesc.BackColor = vbWhite
End Sub

Private Sub txtItemCodeGen_GotFocus()
txtItemCodeGen.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtItemCodeGen_LostFocus()
 txtItemCodeGen.BackColor = vbWhite
End Sub

Private Sub txtItemCodeGen_Validate(Cancel As Boolean)
    
    If Len(Trim(txtItemCodeGen.Text)) < 4 Then
        txtItemCodeGen.SetFocus
        MsgBox "Item Code Generation must be of 4 digit"
        Cancel = True
    End If
    
    If Not IsNumeric(txtItemCodeGen.Text) Then
        If Not Mid(M_COMPBILL, 1, 3) = "CIL" Then
          txtItemCodeGen.SetFocus
          MsgBox "Item Code Generation must be number only"
          Cancel = True
       End If
    End If
End Sub

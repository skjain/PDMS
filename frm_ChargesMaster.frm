VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_ChargesMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Charges Master"
   ClientHeight    =   3855
   ClientLeft      =   3480
   ClientTop       =   3525
   ClientWidth     =   7035
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramCmd 
      Height          =   900
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   6855
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
         Image           =   "frm_ChargesMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4080
         TabIndex        =   18
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
         Image           =   "frm_ChargesMaster.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1464
         TabIndex        =   12
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
         Image           =   "frm_ChargesMaster.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2808
         TabIndex        =   13
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
         Image           =   "frm_ChargesMaster.frx":14BE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5400
         TabIndex        =   19
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
         Image           =   "frm_ChargesMaster.frx":1910
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6825
      Begin VB.CheckBox ChkSale 
         Alignment       =   1  'Right Justify
         Caption         =   "Consider In Sale "
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
         Left            =   4440
         TabIndex        =   16
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Edityn 
         Height          =   300
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   7
         Top             =   600
         Width           =   225
      End
      Begin VB.TextBox COSTEFFECT 
         Height          =   300
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1200
         Width           =   225
      End
      Begin VB.TextBox txtPERC 
         Height          =   300
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1560
         Width           =   1425
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1170
         TabIndex        =   5
         Top             =   960
         Width           =   2760
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1065
      End
      Begin WelchButton.lvButtons_H cmdView 
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Help/Edit"
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
         Image           =   "frm_ChargesMaster.frx":1D62
         cBack           =   -2147483633
      End
      Begin VB.Label Label6 
         Caption         =   "Edit % while Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4440
         TabIndex        =   6
         Top             =   600
         Width           =   1785
      End
      Begin VB.Label Label5 
         Caption         =   "Effect In Cost [Y/N] "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         TabIndex        =   10
         Top             =   1200
         Width           =   1785
      End
      Begin VB.Label Label4 
         Caption         =   "Perc (%) "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Code "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   225
         TabIndex        =   2
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Charges Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   15
      Top             =   45
      Width           =   6870
   End
End
Attribute VB_Name = "frm_ChargesMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************
'Module Name : frm_BillEntrySetup
'
'Develope By : Mr. Ankur
'
'Develope Date : -
'
'Change Date : -
'
'Change By : -
'
'Changes : -
'*******************************************

Dim SAVEFLAG As Boolean

Private Sub cmdAdd_Click()
    SAVEFLAG = True
    Call SetControlEnabled(Yes)
    txtCode = GENCode
    TXTNAME.SetFocus
End Sub

Private Sub cmdCancel_Click()
    SAVEFLAG = False
    Call SetControlEnabled(NO)
    cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim rsValidate As Recordset
    Dim M_OBJDEP As String
    Dim P1 As Integer
    Dim P2 As Integer

On Error GoTo ERRDELETE
    If txtCode = Empty Then Call cmdEdit_Click
    If txtCode = Empty Then Exit Sub
    
    If txtCode < 16 Then
        MsgBox "Charges column is not allowed to delete !! Access Denied !!", vbInformation
        Call cmdCancel_Click
        Exit Sub
    End If
    
    Dim PERC_COL As String
    Set rsValidate = New Recordset
    PERC_COL = "PER" + Trim(TXTNAME)
    SQL = "SELECT * FROM CONFIG WHERE NICK='" & TXTNAME & "'"
    
    rsValidate.Open SQL, CN
    
    If rsValidate.EOF Then
        CN.BeginTrans
            If MsgBox("Are You Sure ? Want to Remove this Column ??", vbQuestion + vbYesNo) = vbYes Then
                
                
                
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='EGPMAN')) ALTER TABLE EGPMAN DROP COLUMN " & TXTNAME
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & PERC_COL & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='EGPMAN')) ALTER TABLE EGPMAN DROP COLUMN " & PERC_COL
                
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='PURMAN')) ALTER TABLE PURMAN DROP COLUMN " & TXTNAME
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & PERC_COL & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='PURMAN')) ALTER TABLE PURMAN DROP COLUMN " & PERC_COL
                
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='GRN')) ALTER TABLE GRN DROP COLUMN " & TXTNAME
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & PERC_COL & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='GRN')) ALTER TABLE GRN DROP COLUMN " & PERC_COL
                
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='WOMST')) ALTER TABLE WOMST DROP COLUMN " & TXTNAME
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & PERC_COL & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='WOMST')) ALTER TABLE WOMST DROP COLUMN " & PERC_COL
                
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='BILLMAIN')) ALTER TABLE BILLMAIN DROP COLUMN " & TXTNAME
                CN.Execute "IF EXISTS(SELECT * FROM SYSCOLUMNS WHERE NAME='" & PERC_COL & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='BILLMAIN')) ALTER TABLE BILLMAIN DROP COLUMN " & PERC_COL
                          
                
                CN.Execute "DELETE FROM CHRGMST WHERE CODE='" & txtCode & "'"
                'BEC = BILL ENTRY CHARGES
                'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','BEC','XXXXXXXXXXXXX','" & TXTNAME & "',NULL,'" & txtCode & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
            End If
        CN.CommitTrans
      Else
        MsgBox "Charges column is not allowed to delete !! Access Denied !!", vbInformation
    End If
    
    Call cmdCancel_Click
    
    Exit Sub
    
ERRDELETE:
    If InStr(1, Err.Description, "is dependent", vbTextCompare) > 0 Then
        If InStr(1, Err.Description, "DF_", vbTextCompare) > 0 Then
            P1 = InStr(1, Err.Description, "'", vbTextCompare) + 1
            P2 = InStr(P1 + 1, Err.Description, "'", vbTextCompare)
            P2 = P2 - P1
            M_OBJDEP = Mid(Err.Description, P1, P2)
            
            
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdEdit_Click()
    Call SetControlEnabled(Yes)
    M_DESC = Empty
    Key = Empty
    TXTNAME = SearchList1("Select TOP 20 Code,Name From CHRGMST", 0, Empty, "Select Charge Name")
    txtCode = Key
    If txtCode = Empty Then Call cmdCancel_Click: Exit Sub
    TXTNAME.Enabled = False
    txtPerc.SetFocus
End Sub

Private Sub cmdExit_Click()
    CN.Execute "UPDATE CHRGMST SET EXTRA1='N' WHERE (EXTRA1 IS NULL)"
    CN.Execute "UPDATE CHRGMST SET EXTRA2='Y' WHERE (EXTRA2 IS NULL)"
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveClick
    Dim Ctrl As Control
    
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
 
    If TXTNAME = Empty Then
        MsgBox "Please Enter Valid Name !!", vbInformation, "Charges Name is MisMatch"
        TXTNAME.SetFocus
        Exit Sub
    End If
    If Edityn.Text = "Y" Or Edityn.Text = "N" Then
      'O.k
     Else
      Edityn.Text = "N"
    End If
    If COSTEFFECT = "Y" Or COSTEFFECT = "N" Then
      'O.k
     Else
      COSTEFFECT = "N"
    End If
    If SAVEFLAG Then txtCode = GENCode
    
    Call AddColumn
    
    Dim PERC_COL As String
    Set rsValidate = New Recordset
    PERC_COL = "PER" + Trim(TXTNAME)
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM CHRGMST WHERE CODE='" & txtCode & "'", CN, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
      CN.Execute "UPDATE CHRGMST SET EXTRA1='" & COSTEFFECT & "',EXTRA2='" & Edityn.Text & "',EXTRA3='" & ChkSale.Value & "' WHERE CODE='" & txtCode & "'"
      Call cmdCancel_Click
      ChkSale.Value = 0
      Exit Sub
    End If
    If SAVEFLAG Then
        If ColumnExists Then
            MsgBox "Column is already Exists !!", vbInformation
            TXTNAME.SetFocus
            SendKeys "{HOME}+{END}"
            Exit Sub
        End If
        CN.BeginTrans
            CN.Execute "INSERT INTO CHRGMST(CODE,NAME,PERC,EXTRA1,EXTRA2,EXTRA3) VALUES('" & txtCode & "','" & TXTNAME & "'," & Val(txtPerc) & ",'" & COSTEFFECT & "','" & Edityn & "','" & ChkSale.Value & "')"
            
            
            
            
            CN.Execute "ALTER TABLE EGPMAN ADD " & TXTNAME & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            CN.Execute "ALTER TABLE EGPMAN ADD " & PERC_COL & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            
            'CN.Execute "ALTER TABLE PURMAN ADD " & txtName & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            'CN.Execute "ALTER TABLE PURMAN ADD " & PERC_COL & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            
            CN.Execute "ALTER TABLE GRN ADD " & TXTNAME & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            CN.Execute "ALTER TABLE GRN ADD " & PERC_COL & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            
            CN.Execute "ALTER TABLE JOBGRN ADD " & TXTNAME & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            CN.Execute "ALTER TABLE JOBGRN ADD " & PERC_COL & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            
            'CN.Execute "ALTER TABLE WOMST ADD " & txtName & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            'CN.Execute "ALTER TABLE WOMST ADD " & PERC_COL & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            
            CN.Execute "ALTER TABLE BILLMAIN ADD " & TXTNAME & " DECIMAL(18,2) NOT NULL DEFAULT 0"
            CN.Execute "ALTER TABLE BILLMAIN ADD " & PERC_COL & " DECIMAL(18,2) NOT NULL DEFAULT 0"
                    
            'CN.Execute "ALTER TABLE QUOTATION ADD " & txtName & " DECIMAL(18,3) NOT NULL DEFAULT 0"
            'CN.Execute "ALTER TABLE QUOTATION ADD " & PERC_COL & " DECIMAL(18,3) NOT NULL DEFAULT 0"
            'CN.Execute "ALTER TABLE QUOTATION ADD " & (txtName & "RMRK") & " CHAR(50) NULL"
                        
                                    
            'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','BCH','XXXXXXXXXXXXX','" & TXTNAME & "',NULL,'" & txtCode & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','N')"
            CN.Execute "INSERT INTO REPFLD(RPCD,POSN,FLDN,[DESC],LNTH,TYPE) VALUES('REG'," & GETLASTPOS & ",'" & TXTNAME & "','" & TXTNAME & "',12.2,'DEC')"
        CN.CommitTrans
    End If
    
    Call cmdCancel_Click
    Exit Sub
    
errSaveClick:
    CN.RollbackTrans
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdView_Click()
    TXTNAME = Empty
    Key = Empty
    CANCEL_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    TXTNAME = SearchList1("Select TOP 20 Code,Name From CHRGMST ", 0, Empty, "Select Charge To View")
    txtCode = Key
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM CHRGMST WHERE CODE='" & txtCode & "'", CN, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
      COSTEFFECT = Trim(RS!extra1 & "")
      Edityn = Trim(RS!EXTRA2 & "")
      txtPerc = Format(RS!PERC, "####.00")
      If Trim(RS!EXTRA3 & "") = Empty Then ChkSale.Value = 0
      If Trim(RS!EXTRA3 & "") = "0" Then ChkSale.Value = 0
      If Trim(RS!EXTRA3 & "") = "1" Then ChkSale.Value = 1
      CMDSAVE.Enabled = True
    End If
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
    Call SetControlEnabled(NO)
    
    Exit Sub
errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Function SetControlEnabled(YesNo As Status)
    cmdView.Enabled = Not YesNo
    txtCode.Enabled = False
    TXTNAME.Enabled = YesNo
    CMDSAVE.Enabled = YesNo
    cmdCancel.Enabled = YesNo
    cmdAdd.Enabled = Not YesNo
    txtPerc.Enabled = YesNo
    cmdDelete.Enabled = Not YesNo
    Call ClsData(Me)
End Function

Private Function GENCode() As String
    Dim rsCode As Recordset
    Dim ctr As Byte
    Set rsCode = New Recordset
    
    rsCode.Open "Select Isnull(Max(Code),0) As Code From CHRGMST", CN
    
    If rsCode.EOF Then
        GENCode = "01"
    Else
        If IsNull(rsCode!CODE) Then
            GENCode = "01"
        Else
            ctr = rsCode!CODE
            
            ctr = ctr + 1
            If ctr < 10 Then
                GENCode = "0" & ctr
            Else
                GENCode = ctr
            End If
        End If
    End If
    
    rsCode.Close
End Function

Private Sub txtCode_Change()
    If SAVEFLAG Then Exit Sub
    Dim rsViewData As Recordset
    
    Set rsViewData = New Recordset
    rsViewData.Open "Select code,name,isnull(perc,0) as perc From CHRGMST Where Code='" & txtCode & "'", CN
    
    If rsViewData.EOF = False Then
        txtPerc = rsViewData!PERC
    End If
End Sub

Private Sub txtCode_GotFocus()
txtCode.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtCode_LostFocus()
txtCode.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 97 To 121
            KeyAscii = KeyAscii - 32
        Case 65 To 90
        Case 8  'Back Space
        
        Case 38
        Case 95 'Under Score
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTNAME_LostFocus()
 TXTNAME.BackColor = vbWhite
    If InStr(1, TXTNAME, "&", vbTextCompare) Then
        TXTNAME = "[" & TXTNAME & "]"
    End If
End Sub

Private Sub txtPERC_GotFocus()
txtPerc.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtPERC_KeyPress(KeyAscii As Integer)
    Call CheckNumericKey(KeyAscii, txtPerc, Me)
End Sub

Private Function ColumnExists() As Boolean
    Dim rsValidateCol As Recordset
    Set rsValidateCol = New Recordset

    If rsValidateCol.State = 1 Then rsValidateCol.Close
    rsValidateCol.Open "Select * from syscolumns where name='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='EGPMAN')", CN
    If rsValidateCol.State = 1 Then rsValidateCol.Close
    rsValidateCol.Open "Select * from syscolumns where name='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='PURMAN')", CN
    If rsValidateCol.State = 1 Then rsValidateCol.Close
    rsValidateCol.Open "Select * from syscolumns where name='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='GRN')", CN
    If rsValidateCol.State = 1 Then rsValidateCol.Close
    rsValidateCol.Open "Select * from syscolumns where name='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='WOMST')", CN
    If rsValidateCol.State = 1 Then rsValidateCol.Close
    rsValidateCol.Open "Select * from syscolumns where name='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='BILLMAIN')", CN
    If rsValidateCol.State = 1 Then rsValidateCol.Close
    rsValidateCol.Open "Select * from syscolumns where name='" & TXTNAME & "' AND ID=(SELECT ID FROM SYSOBJECTS WHERE NAME='QUOTATION')", CN
        
    If Not rsValidateCol.EOF Then
        ColumnExists = True
    End If
    rsValidateCol.Close
End Function

Private Function GETLASTPOS() As Integer
On Error GoTo ERRPOS
    Set RS = New Recordset
    RS.Open "Select MAX(POSN) AS POS From REPFLD Where RPCD='REG'", CN
    If RS.EOF = False Then
        GETLASTPOS = RS!POS + 1
    Else
        GETLASTPOS = 1
    End If
    
    RS.Close
    Exit Function
    
ERRPOS:
    MsgBox Err.Description
End Function

Private Sub txtPerc_LostFocus()
txtPerc.BackColor = vbWhite
End Sub

Private Sub AddColumn()
On Error Resume Next
  If ChkSale.Value = 1 Then
     CN.Execute "ALTER TABLE TAXMST ADD " & TXTNAME & " DECIMAL(18,2) NOT NULL DEFAULT 0"
     CN.Execute "ALTER TABLE RATEMST ADD " & TXTNAME & " DECIMAL(18,2) NOT NULL DEFAULT 0"
  End If
End Sub

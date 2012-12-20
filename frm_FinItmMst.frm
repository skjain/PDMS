VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frm_FinItmMst 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finish Item Master"
   ClientHeight    =   4305
   ClientLeft      =   3360
   ClientTop       =   2235
   ClientWidth     =   8865
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
   ScaleHeight     =   4305
   ScaleWidth      =   8865
   Begin VB.Frame FramCont 
      Height          =   2625
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   8385
      Begin VB.TextBox TXTCONVERSION 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "1"
         ToolTipText     =   "Meter Per KG"
         Top             =   2160
         Width           =   675
      End
      Begin VB.TextBox TXTDENI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1920
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1080
         Width           =   6225
      End
      Begin VB.CheckBox chkCopsReturnable 
         Caption         =   "Cops Returnable "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtQty 
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
         Left            =   6720
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "Q"
         Top             =   1800
         Width           =   450
      End
      Begin VB.TextBox txtUOM 
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1800
         Width           =   1275
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Enter the Description of Item."
         Top             =   720
         Width           =   5475
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conversion :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " [Q / P / X ]"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7320
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblQORP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity/Pieces :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   1800
         Width           =   1680
      End
      Begin VB.Label ICOD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Item Code."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Item        :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit of Measure :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Description :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1710
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3600
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
      Image           =   "frm_FinItmMst.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   3600
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
      Image           =   "frm_FinItmMst.frx":039A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   5880
      TabIndex        =   11
      Top             =   3600
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
      Image           =   "frm_FinItmMst.frx":0734
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3600
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
      Image           =   "frm_FinItmMst.frx":0ACE
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   3600
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
      Image           =   "frm_FinItmMst.frx":1858
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   3600
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
      Image           =   "frm_FinItmMst.frx":1CAA
      cBack           =   -2147483633
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "FINISH ITEM MASTER "
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
      Left            =   5640
      TabIndex        =   21
      Top             =   240
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   8760
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   8760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   4095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label lblDivName 
      BackStyle       =   0  'Transparent
      Caption         =   "Division Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   20
      Top             =   240
      Width           =   5505
   End
End
Attribute VB_Name = "frm_FinItmMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_IGCD As String
Dim SAVEFLAG As Boolean, iGrpFlag As Boolean
Dim SQL As String
Public ONLINEITEM As Boolean
Public ISRETURNABLE As String
Dim i As Long

Private Sub chkCopsReturnable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub cmdAdd_Click()
    Call btn_sts(False)
    txtQty.Text = "Q"
    txtDesc.SetFocus
    ICOD.Caption = GENICODE
    SAVEFLAG = True
    cmdCancel.Cancel = True
End Sub

Private Sub cmdAdd_GotFocus()
    Msg cmdAdd.ToolTipText
End Sub

Private Sub cmdCancel_Click()
    ICOD.Caption = GENICODE
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    chkCopsReturnable.Value = 0
End Sub

Private Sub cmdCancel_GotFocus()
    Msg cmdCancel.ToolTipText
End Sub

Private Sub cmdDelete_Click()
If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("000007", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    
On Error GoTo MsgFlash
    Dim ANS As String
    Dim TEMPRS As New ADODB.Recordset
    
    Call cmdEdit_Click
    
    If isFurtherEntryExist("FITEM", Trim(ICOD.Caption)) Then
       MsgBox "Further Entry Exist"
       ICOD.Caption = GENICODE
       Call btn_sts(True)
       Call ClsData(Me)
       Exit Sub
    End If
  
    
    If txtDesc.Text = "" Or ICOD.Caption = Empty Then
        MsgBox "There is no Record to delete.", vbCritical, App.TITLE
        Exit Sub
    End If
                
    Dim STR(4) As String
                   
    STR(0) = "SPTRAN"
    STR(1) = "PURTRAN"
    STR(2) = "STORETRAN"
    
        
    For i = 0 To 2
     If TEMPRS.State = 1 Then TEMPRS.Close
     TEMPRS.Open "Select * from " & STR(i) & " where COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND ICOD ='" & ICOD.Caption & "'", CN, adOpenDynamic, adLockOptimistic
     If Not TEMPRS.EOF Then MsgBox "Item Entry Exist in Database : " & STR(i), vbCritical, App.TITLE: TEMPRS.Close: Exit Sub
    Next i
    
    If TEMPRS.State = 1 Then TEMPRS.Close
    ANS = MsgBox("Are You Sure To delete this record ? ", vbYesNo)
    If ANS = vbYes Then
        CN.Execute "DELETE FROM FINITMMST where COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND CODE ='" & Trim(ICOD.Caption) & "'"
        CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','FTM','XXXXXXXXXXXXX','" & txtDesc & "',NULL,'" & ICOD.Caption & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
        '----------------------------------
        'DAILYSTATUS
        Call DAILYSTATUS("ITM", ICOD, "", 0, "", 0, cUName, "D", Now, Now)
        '----------------------------------
    End If
    
    ICOD.Caption = GENICODE
    Call btn_sts(True)
    Call ClsData(Me)
    Exit Sub
    
MsgFlash:
   MsgBox "Error Number : " & ERR.Description & ". Error Description " & ERR.Description
End Sub

Private Sub cmdDelete_GotFocus()
    Msg cmdDelete.ToolTipText
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("000007", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    
  SAVEFLAG = False
  NEW_VISIBLE = False
  Key = Empty
  
  txtDesc = SearchList1("SELECT  TOP 20 CODE,NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "'", 0, Empty, "Select Item From List")
  If Not txtDesc = Empty Then
     Call SetData(Key)
     txtDesc.Tag = Key
     Call btn_sts(False)
     txtDesc.SetFocus
  Else
    Call cmdCancel_Click
  End If
End Sub

Private Sub cmdEdit_GotFocus()
    Msg cmdEdit.ToolTipText
End Sub

Private Sub CMDEXIT_Click()
    Msg Empty
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdExit_GotFocus()
    Msg cmdExit.ToolTipText
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveClick
    Dim TEMPRS As New ADODB.Recordset
    Dim Ctrl As Control
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
    
    If txtDesc.Text = "" Then
        MsgBox "Description can not be empty.", vbCritical, App.TITLE
        txtDesc.Enabled = True: txtDesc.SetFocus
        Exit Sub
    End If
           
    If txtUOM.Text = "" Then
        MsgBox "Unit Can not be empty.", vbCritical, App.TITLE
        txtUOM.Enabled = True: txtUOM.SetFocus
        Exit Sub
    End If
    
    If txtdeni.Text = "" Then
        MsgBox "Denior can not be empty.", vbCritical, App.TITLE
        txtdeni.Enabled = True: txtdeni.SetFocus
        Exit Sub
    End If
      
    If txtQty.Text = "" Or UCase(txtQty.Text) <> "Q" And UCase(txtQty.Text) <> "P" And UCase(txtQty.Text) <> "X" Then
        MsgBox "Please Enter Q / P / X.", vbInformation, App.TITLE
        txtQty.SetFocus
        Exit Sub
    End If
    
    If chkCopsReturnable.Value = 1 Then
       ISRETURNABLE = "Y"
    Else
       ISRETURNABLE = "N"
    End If
          
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from FINITMMST where [NAME] = '" & Trim(txtDesc.Text) & _
    "' and COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD = '" & DIVCOD & "'", CN, adOpenDynamic, adLockOptimistic
    
    If TEMPRS.EOF = False And SAVEFLAG Then
       MsgBox "Can not insert Duplicate Record.", vbCritical, App.TITLE
       txtDesc.SetFocus
       TEMPRS.Close
       Exit Sub
    End If
    
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "SELECT  CODE FROM REFMST WHERE CATA='U' AND NAME='" & Trim(txtUOM) & "'", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = False Then
       txtUOM.Tag = TEMPRS!CODE & ""
    Else
       MsgBox "Unit of Mearsurement missing", vbCritical, App.TITLE
    End If
        
    CN.BeginTrans
    
    If SAVEFLAG = True Then
        ICOD.Caption = GENICODE
        
    SQL = "INSERT INTO FINITMMST(COMP,UNIT,DVCD,CODE,NAME,DENI,UOM,QORP,ISRETURNABLE,CONVERSION)" _
    & " VALUES('" & compPth & "','" & UNCD & "','" & DIVCOD & "','" & ICOD & "','" & txtDesc & _
    "','" & Replace(txtdeni.Text, vbCrLf, "") & "','" & txtUOM.Tag & "','" & txtQty & _
    "','" & ISRETURNABLE & "','" & Val(TXTCONVERSION) & "')"
   
    Else
    
    SQL = "UPDATE FINITMMST SET ISRETURNABLE='" & ISRETURNABLE & "',NAME='" & txtDesc & "',DENI='" & Replace(txtdeni.Text, vbCrLf, "") & _
    "',UOM='" & txtUOM.Tag & "',QORP='" & txtQty & "',CONVERSION='" & Val(TXTCONVERSION) & "' WHERE COMP='" & compPth & _
    "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND CODE='" & ICOD & "'"
        
    End If
    
    CN.Execute SQL
    
    '--------------------------
    'DAILYSTATUS ENTRY
    If SAVEFLAG Then
     Call DAILYSTATUS("ITM", ICOD, "", 0, "", 0, cUName, "N", Now, Now)
     Else
     Call DAILYSTATUS("ITM", ICOD, "", 0, "", 0, cUName, "M", Now, Now)
    End If
    '--------------------------
    
    CN.CommitTrans
    
    If SAVEFLAG Then MsgBox "Item Saved Successfully"
    If Not SAVEFLAG Then MsgBox "Item Details Update Successfully"
    
    ICOD.Caption = GENICODE
    sTxt = txtDesc.Text
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub
   
errSaveClick:
    If InStr(1, ERR.Description, "more transaction", vbTextCompare) > 0 Then
        CN.RollbackTrans
    ElseIf ERR.Number = -2147217873 Then
        MsgBox "Item Name Already Exists....", vbInformation, App.TITLE
        Exit Sub
    Else
        LOAD frm_ErrorHandler
        ErrNumber = ERR.Number
        ErrMessage = ERR.Description
        frm_ErrorHandler.Show vbModal
        ERR.Clear
    End If
End Sub

Private Sub cmdSave_GotFocus()
    Msg cmdSave.ToolTipText
End Sub

Private Sub Form_Activate()
If DIVNAM = Empty Or DIVCOD = Empty Then
   MsgBox "SELECT DIVISION"
   Unload Me
   Exit Sub
End If

  Call ColorComponent(Me)
  LBLHEAD.BackColor = &H80&
  LBLHEAD.ForeColor = &HFFFFFF
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  ICOD.Caption = GENICODE
  If Not ONLINEITEM Then
  DIVCOD = Empty: DIVNAM = Empty: Key = Empty
  DIVNAM = SearchList1("SELECT TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION FROM MASTER FOR ITEM")
  DIVCOD = Key:  Me.Tag = DIVCOD
  End If
  ONLINEITEM = False
  lblDivName.Caption = DIVNAM
  Me.Caption = "Finish Item Master For " + UCase(DIVNAM)
  
  Call btn_sts(True)
  Me.KeyPreview = True
  Exit Sub
errLoad:
    MsgBox "Error No. " & ERR.Description & "  Error Description : " & ERR.Description
End Sub

Private Sub btn_sts(Yes As Boolean)
    txtDesc.Enabled = Not Yes
    txtUOM.Enabled = Not Yes
    txtQty.Enabled = Not Yes
    txtdeni.Enabled = Not Yes
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
End Sub


Private Sub TXTCONVERSION_GotFocus()
  SendKeys "{HOME}+{END}"
  TXTCONVERSION.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCONVERSION_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If cmdSave.Enabled Then cmdSave.SetFocus: Exit Sub
  End If
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub TXTCONVERSION_LostFocus()
  TXTCONVERSION.BackColor = vbWhite
End Sub

Private Sub txtDENI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtdeni <> Empty Then SendKeys "{TAB}"
End Sub

Private Sub txtDENI_LostFocus()
txtdeni.BackColor = vbWhite
End Sub

Private Sub txtDesc_LostFocus()
 txtDesc.BackColor = vbWhite
End Sub

Private Sub txtDENI_GotFocus()
     txtdeni.BackColor = RGB(BRED, BGREEN, BBLUE)
     Msg "Enter Denier Name"
     SendKeys "{END}"
End Sub

Private Sub txtDesc_GotFocus()
 txtDesc.BackColor = RGB(BRED, BGREEN, BBLUE)
 txtDesc.SelStart = 0
 txtDesc.SelLength = Len(txtDesc)
 Msg "Enter Item Name"
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtDesc <> Empty Then
       If txtdeni = Empty Then txtdeni = txtDesc
       txtdeni.SetFocus
    End If
    
    Select Case KeyAscii
        Case 34, 39
            KeyAscii = 0
    End Select
End Sub

Private Sub txtQty_Change()
    If cmdSave.Enabled = False Then Exit Sub
End Sub

Private Sub TXTQTY_GotFocus()
    txtQty.BackColor = RGB(BRED, BGREEN, BBLUE)
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty)
    Msg "Enter Q => Quantity / P => Pieces"
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtQty <> Empty Then SendKeys "{TAB}"
    If Not (UCase(Chr(KeyAscii)) = "Q" Or UCase(Chr(KeyAscii)) = "P" Or UCase(Chr(KeyAscii)) = "X") And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Function GENICODE() As String
  
  Dim ITMRS As New ADODB.Recordset
  Set ITMRS = New ADODB.Recordset
           
  If ITMRS.State = 1 Then ITMRS.Close
  ITMRS.Open "Select IsNull(Max(CODE),0) AS CODE From FINITMMST WHERE COMP='" & compPth & _
  "' AND UNIT='" & UNCD & "' AND CODE LIKE '00000%'", CN, adOpenDynamic, adLockOptimistic
        
  If Trim(ITMRS!CODE) = "0" Then  'C1
   GENICODE = "0000000001"
  Else
  
   GENICODE = Val(ITMRS!CODE) + 1
   ITMRS.Close
   
   If GENICODE < 10 Then
      GENICODE = "000000000" & GENICODE
   ElseIf GENICODE < 100 Then
      GENICODE = "00000000" & GENICODE
   ElseIf GENICODE < 1000 Then
      GENICODE = "0000000" & GENICODE
   ElseIf GENICODE < 10000 Then
      GENICODE = "000000" & GENICODE
   ElseIf GENICODE < 100000 Then
      GENICODE = "00000" & GENICODE
   ElseIf GENICODE < 1000000 Then
      GENICODE = "0000" & GENICODE
   ElseIf GENICODE < 10000000 Then
      GENICODE = "000" & GENICODE
   ElseIf GENICODE < 100000000 Then
      GENICODE = "00" & GENICODE
   ElseIf GENICODE < 1000000000 Then
     GENICODE = "0" & GENICODE
   Else
      GENICODE = GENICODE
   End If
 End If    'C1
End Function

Private Sub TXTQTY_LostFocus()
 txtQty.BackColor = vbWhite
End Sub

Private Sub TXTUOM_GotFocus()
 txtUOM.BackColor = RGB(BRED, BGREEN, BBLUE)
 txtUOM.SelStart = 0
 txtUOM.SelLength = Len(txtUOM)
 Msg "Enter Unit of Measurement"
End Sub

Private Sub txtUOM_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Or Trim(txtUOM.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtUOM.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM REFMST WHERE CATA='U'", 0, txtUOM, "Select Unit for Finish Item")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            txtUOM.Text = ""
            Ref_Cat = "U"
            Frm_Ref_FAS.Show
        Else
            txtUOM.Tag = Key
        End If
        
    ElseIf KeyCode = vbKeyDelete Then
        txtUOM = Empty
    End If
    If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
End Sub

Private Sub TXTUOM_LostFocus()
 txtUOM.BackColor = vbWhite
End Sub

Private Sub SetData(CODE As String)
Dim ITMRS As New ADODB.Recordset
Set ITMRS = New ADODB.Recordset
  
If ITMRS.State = 1 Then ITMRS.Close
ITMRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND CODE='" & CODE & "'", CN, adOpenDynamic, adLockOptimistic
If Not ITMRS.EOF Then
   ICOD.Caption = CODE
   txtDesc = ITMRS!NAME
   txtdeni = ITMRS!DENI
   TXTCONVERSION = ITMRS!CONVERSION
   txtQty = ITMRS!QORP
   txtUOM = GetCode("REFMST", ITMRS!UOM, "CODE", "NAME")
   If Trim(ITMRS!ISRETURNABLE & "") = "Y" Then
      chkCopsReturnable.Value = 1
   Else
      chkCopsReturnable.Value = 0
   End If
End If
ITMRS.Close

End Sub

VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmVehicleMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Master"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9120
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   5355
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9446
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   8438015
      BackColor       =   16777215
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
      Begin VB.TextBox txtCode 
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
         Left            =   4560
         MaxLength       =   49
         TabIndex        =   14
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtName 
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
         Left            =   360
         MaxLength       =   49
         TabIndex        =   2
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1320
         Width           =   3195
      End
      Begin VB.TextBox txtTransport 
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
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   140
         TabIndex        =   4
         Top             =   2040
         Width           =   5205
      End
      Begin VB.ListBox lstRef 
         Height          =   3180
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtLimit 
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
         Left            =   360
         MaxLength       =   140
         TabIndex        =   6
         Top             =   2760
         Width           =   3165
      End
      Begin ButtonPlusCtl.ButtonPlus cmdFind 
         Height          =   375
         Left            =   7560
         TabIndex        =   15
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Find"
      End
      Begin ButtonPlusCtl.ButtonPlus cmdClear 
         Height          =   375
         Left            =   6360
         TabIndex        =   16
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C&lear"
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   720
         TabIndex        =   0
         Top             =   4560
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
         Image           =   "frmVehicleMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4680
         TabIndex        =   9
         Top             =   4560
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
         Image           =   "frmVehicleMaster.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6000
         TabIndex        =   10
         Top             =   4560
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
         Image           =   "frmVehicleMaster.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   4560
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
         Image           =   "frmVehicleMaster.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   4560
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
         Image           =   "frmVehicleMaster.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   11
         Top             =   4560
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
         Image           =   "frmVehicleMaster.frx":1CAA
         cBack           =   -2147483633
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3360
         TabIndex        =   17
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle &No.    "
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
         Left            =   405
         TabIndex        =   1
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tra&nsport Name"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   600
         Y2              =   4440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   9000
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   150
         X2              =   8880
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   5175
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity &Limit"
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
         Left            =   360
         TabIndex        =   5
         Top             =   2400
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmVehicleMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean

Private Sub cmdAdd_Click()
    Call ClsData(Me)
    Call btn_sts(False)
    
    TXTNAME.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
End Sub

Private Sub cmdCLEAR_Click()
    Call ClsData(Me)
    lstRef.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000023", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
End If

    Dim ANS As String, TEMPRS As New ADODB.Recordset
    
    If isFurtherEntryExist("VEHICLE", txtCode) Then
       MsgBox "Further Entry Exist"
       Call ClsData(Me)
       lstRef.ListIndex = -1
       Call btn_sts(True)
       Exit Sub
    End If
    
    
    If txtCode.Text = "" Then Exit Sub
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.Title)
    If ANS = vbYes Then
       CN.Execute "Delete from VHCLMST where CODE ='" & Trim(txtCode.Text) & "'"
       '---------------------------------
       'DAILYSTATUS
       Call DAILYSTATUS("VLM", txtCode, "", 0, "", 0, cUName, "D", Now, Now)
       '---------------------------------
       lstRef.RemoveItem lstRef.ListIndex
    End If
                
    Call ClsData(Me)
    lstRef.ListIndex = -1
    Call btn_sts(True)
End Sub

Private Sub cmdEdit_Click()
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000023", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    
End If
    cmdCancel.Cancel = True
    Call btn_sts(False)
    
    If lstRef.ListIndex = -1 Then lstRef.SetFocus Else TXTNAME.SetFocus
    SAVEFLAG = False
    
End Sub

Private Sub cmdExit_Click()
    key_PressNew = False
    Unload Me
End Sub

Private Sub CMDFIND_Click()
    NEW_VISIBLE = False
    If Me.Tag <> Empty Then Ref_Cat = Me.Tag
    M_DESC = Empty
    Key = Empty
    TXTNAME.Text = SearchList1("Select TOP 20 CODE, NAME FROM VHCLMST WHERE RECSTAT<>'D'", 0, "", "List Of " & Me.Caption)
    txtCode.Text = Key
    
    lstRef.Text = TXTNAME.Text
        
    If TXTNAME <> Empty Then
       TXTNAME.Enabled = True
       TXTNAME.SetFocus
    End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo errPRIMARYKEY
    Dim SQL As String
    Dim TEMPRS As New ADODB.Recordset
    Dim Ctrl As Control
    
    TXTNAME.Text = Trim(TXTNAME.Text)
    
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
           
    If Trim(TXTNAME.Text) = "" Then
        MsgBox "Please Enter Transporter Name.", vbInformation, App.Title
        TXTNAME = Trim(TXTNAME)
        TXTNAME.SetFocus
        Exit Sub
    End If
              
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from VHCLMST where Upper([name])='" & UCase(Trim(TXTNAME.Text)) & "' and recstat<>'d'", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = False And SAVEFLAG Then
       MsgBox "This Name Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.Title
       TEMPRS.Close
       Exit Sub
    End If
    
    If SAVEFLAG = True Then
        On Error GoTo errPRIMARYKEY
        
        txtCode.Text = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From VHCLMST ")
          
        SQL = "insert into VHCLMST (CODE,[NAME],TRCD,CAPACITY,RECSTAT)" _
                  & " values('" & Trim(txtCode.Text) & "','" & UCase(Trim(TXTNAME.Text)) & _
                  "','" & Trim(txtTransport.Tag) & "','" & Val(txtLimit) & "','A')"
        
        CN.BeginTrans
        CN.Execute SQL
        '---------------------------
        'DAILYSTATUS
        Call DAILYSTATUS("VLM", txtCode, "", 0, "", 0, cUName, "N", Now, Now)
        '---------------------------
        CN.CommitTrans
        
        lstRef.AddItem UCase(TXTNAME.Text)
    Else
    CN.BeginTrans
    CN.Execute ("Update VHCLMST set NAME = '" & UCase(Trim(TXTNAME.Text)) & "',TRCD = '" & txtTransport.Tag & _
    "',CAPACITY='" & Val(txtLimit) & "' where CODE ='" & Trim(txtCode.Text) & "' AND RECSTAT<>'D'")
    
    
    '-------------------------------------
    'DAILYSTATUS
    Call DAILYSTATUS("VLM", txtCode, "", 0, "", 0, cUName, "M", Now, Now)
    '-------------------------------------
    CN.CommitTrans
    lstRef.Clear
    Call FillList("Select [NAME] from VHCLMST WHERE RECSTAT<>'D' ORDER BY [NAME]", lstRef)
     
    lstRef.ListIndex = -1
    End If
  
    Call btn_sts(True)
    sTxt = TXTNAME.Text
 
    Call ClsData(Me)
    Call cmdCancel_Click
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub

errPRIMARYKEY:
    CN.RollbackTrans
    If Err.Number = -2147217873 Or -2147217900 Then
        TXTNAME.SetFocus
        MsgBox "This Name Already Registered With Other Category!!!", vbInformation, "Already Registered"
    Else
        ErrNumber = Err.Number
        ErrMessage = Err.Description
        frm_ErrorHandler.Show vbModal
    End If
    Err.Clear
End Sub

Private Sub Form_Activate()
    Call ColorComponent(Me)
    Me.BackColor = RGB(RED, GREEN, BLUE)
    If key_PressNew Then cmdAdd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ActiveControl.NAME = "lstRef" Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
On Error GoTo errLoad
  Call btn_sts(True)
    
  Call FillList("Select [NAME] from VHCLMST WHERE RECSTAT<>'D'", lstRef)
    
  cmdExit.Cancel = True
  Me.Show
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
    TXTNAME.Enabled = Not bool
    txtTransport.Enabled = Not bool
    txtLimit.Enabled = Not bool
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Msg Empty
End Sub

Private Sub lstRef_Click()
    SAVEFLAG = False
    Dim TEMPRS As New ADODB.Recordset
    Dim NAME As String
    If lstRef.ListIndex = -1 And TXTNAME <> Empty Then
       NAME = TXTNAME
    ElseIf Trim(lstRef.List(lstRef.ListIndex)) <> Empty Then
       NAME = (lstRef.List(lstRef.ListIndex))
    Else
       Exit Sub
    End If
    
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "SELECT VHCLMST.*,TRANSPORTMST.NAME AS TRANSPORT FROM VHCLMST LEFT JOIN TRANSPORTMST ON VHCLMST.TRCD=TRANSPORTMST.CODE WHERE VHCLMST.NAME = '" & NAME & "' AND VHCLMST.RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
    If Not TEMPRS.EOF Then
    With TEMPRS
        txtCode = Trim(!CODE & "")
        TXTNAME = Trim(![NAME] & "")
        txtTransport = Trim(!transport & "")
        txtTransport.Tag = Trim(!TRCD & "")
        txtLimit = Trim(!CAPACITY & "")
    End With
    End If
    TEMPRS.Close
End Sub

Private Sub lstRef_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 TXTNAME.Enabled = True
 TXTNAME.SetFocus
End If
End Sub

Private Sub lstRef_GotFocus()
    lstRef.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Address"
End Sub

Private Sub lstRef_LostFocus()
lstRef.BackColor = vbWhite
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtLimit, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTransport_GotFocus()
    txtTransport.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtLimit_LostFocus()
 txtLimit.BackColor = vbWhite
End Sub

Private Sub txtLimit_GotFocus()
    txtLimit.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtTransport_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(txtTransport) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtTransport.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM TRANSPORTMST WHERE RECSTAT<>'D'", 0, txtTransport.Text, "SELECT TRANSPORT FROM LIST")
        If key_PressNew = True Then
          M_DESC = ""
          txtTransport = Empty
          frmTransportMaster.Show
        Else
          txtTransport.Tag = Key
        End If
    End If
    
Me.KeyPreview = True
    
End Sub

Private Sub txtTransport_LostFocus()
 txtTransport.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
    TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Vehicle No."
    TXTNAME.SelStart = 0
    TXTNAME.SelLength = Len(TXTNAME)
End Sub

Private Sub TXTNAME_LostFocus()
TXTNAME.BackColor = vbWhite
End Sub

Public Sub FillList(SQL As String, lst As ListBox)
    Dim TEMPRS As New ADODB.Recordset
    TEMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = True Then Exit Sub
    TEMPRS.MoveFirst
    Do While Not TEMPRS.EOF
        lst.AddItem Trim(TEMPRS.Fields(0).Value)
        TEMPRS.MoveNext
    Loop
    TEMPRS.Close
End Sub

VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmTransportMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transport Master"
   ClientHeight    =   6525
   ClientLeft      =   1890
   ClientTop       =   435
   ClientWidth     =   9105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9105
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   6555
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11562
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
      Begin VB.TextBox TXTNOV 
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
         MaxLength       =   4
         TabIndex        =   16
         Top             =   5280
         Width           =   2325
      End
      Begin VB.TextBox txtph2 
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
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   10
         Top             =   3600
         Width           =   2325
      End
      Begin VB.TextBox txtph1 
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
         MaxLength       =   20
         TabIndex        =   8
         Top             =   3600
         Width           =   2325
      End
      Begin VB.TextBox TXTPAN 
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
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   14
         Top             =   4440
         Width           =   2325
      End
      Begin VB.TextBox txtADD2 
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
         Width           =   5205
      End
      Begin VB.ListBox lstRef 
         Height          =   4350
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtADD1 
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
         TabIndex        =   4
         Top             =   2040
         Width           =   5205
      End
      Begin VB.TextBox txtSTAX 
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
         MaxLength       =   20
         TabIndex        =   12
         Top             =   4440
         Width           =   2325
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
         Width           =   5235
      End
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
         TabIndex        =   26
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin ButtonPlusCtl.ButtonPlus cmdFind 
         Height          =   375
         Left            =   7680
         TabIndex        =   24
         Top             =   5280
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
         Left            =   6480
         TabIndex        =   23
         Top             =   5280
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
         Top             =   5880
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
         Image           =   "frmTransportMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4680
         TabIndex        =   19
         Top             =   5880
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
         Image           =   "frmTransportMaster.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6000
         TabIndex        =   20
         Top             =   5880
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
         Image           =   "frmTransportMaster.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2040
         TabIndex        =   17
         Top             =   5880
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
         Image           =   "frmTransportMaster.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3360
         TabIndex        =   18
         Top             =   5880
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
         Image           =   "frmTransportMaster.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   21
         Top             =   5880
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
         Image           =   "frmTransportMaster.frx":1CAA
         cBack           =   -2147483633
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of &Vehicles Owned"
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
         TabIndex        =   15
         Top             =   4920
         Width           =   2235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone N&o. 2"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   3240
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone &No. 1"
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
         TabIndex        =   7
         Top             =   3240
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&PAN No."
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
         Left            =   3120
         TabIndex        =   13
         Top             =   4080
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address-&2"
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
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   6375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8895
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   150
         X2              =   8880
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   9000
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   600
         Y2              =   5760
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address-&1"
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
         Width           =   1020
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service &Tax No."
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
         TabIndex        =   11
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport &Name     "
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
         Width           =   1905
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transport Master"
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
         Left            =   3240
         TabIndex        =   27
         Top             =   120
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmTransportMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim COND As String

Private Sub cmdAdd_Click()
    Call ClsData(Me)
    Call btn_sts(False)
    
    TXTNAME.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdAdd_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
End Sub

Private Sub cmdCancel_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdCLEAR_Click()
    Call ClsData(Me)
    lstRef.ListIndex = -1
End Sub

Private Sub cmdClear_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdDelete_Click()
If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000015", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    Dim ANS As String, TEMPRS As New ADODB.Recordset
    
    If isFurtherEntryExist("TRANSPORT", txtCode) Then
       MsgBox "Further Entry Exist"
       Call ClsData(Me)
       lstRef.ListIndex = -1
       Call btn_sts(True)
       Exit Sub
    End If
    
    If txtCode.Text = "" Then Exit Sub
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.Title)
    If ANS = vbYes Then
       CN.Execute "Delete from transportmst where CODE ='" & Trim(txtCode.Text) & "'"
        Call DAILYSTATUS("TPT", txtCode, "", 0, "", 0, cUName, "D", Now, Now)
       lstRef.RemoveItem lstRef.ListIndex
    End If
                
    Call ClsData(Me)
    lstRef.ListIndex = -1
    Call btn_sts(True)
End Sub

Private Sub cmdDelete_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdEdit_Click()
If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000015", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    cmdCancel.Cancel = True
    Call btn_sts(False)
    
    If lstRef.ListIndex = -1 Then lstRef.SetFocus Else TXTNAME.SetFocus
    SAVEFLAG = False
    
End Sub

Private Sub cmdEdit_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdExit_Click()
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdExit_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub CMDFIND_Click()
    NEW_VISIBLE = False
    If Me.Tag <> Empty Then Ref_Cat = Me.Tag
    M_DESC = Empty
    Key = Empty
    TXTNAME.Text = SearchList1("Select TOP 20 CODE, NAME FROM TRANSPORTMST WHERE RECSTAT<>'D'", 0, "", "List Of " & Me.Caption)
    txtCode.Text = Key
    
    lstRef.Text = TXTNAME.Text
        
    If TXTNAME <> Empty Then
       TXTNAME.Enabled = True
       TXTNAME.SetFocus
    End If
End Sub

Private Sub cmdFind_GotFocus()
    Msg ActiveControl.ToolTipText
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
    TEMPRS.Open "Select * from transportmst where Upper([name])='" & UCase(Trim(TXTNAME.Text)) & "' and recstat<>'d'", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = False And SAVEFLAG Then
       MsgBox "This Name Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.Title
       TEMPRS.Close
       Exit Sub
    End If
    
    If SAVEFLAG = True Then
        On Error GoTo errPRIMARYKEY
        
        txtCode.Text = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From TRANSPORTMST ")
          
        SQL = "insert into TRANSPORTMST (CODE,[NAME],ADD1,ADD2,PH1,PH2,STAX,PANNO,NOV,RECSTAT)" _
                  & " values('" & Trim(txtCode.Text) & _
                  "','" & UCase(Trim(TXTNAME.Text)) & "','" & Trim(txtAdd1.Text) & _
                  "','" & Trim(txtAdd2.Text) & "','" & Trim(txtph1.Text) & "','" & Trim(txtph2.Text) & _
                  "','" & Trim(txtSTAX.Text) & "','" & TXTPAN & "','" & Val(TXTNOV) & "','A')"
        
        CN.BeginTrans
        CN.Execute SQL
        
        '----------------------------------
        'DAILYSTATUS
        Call DAILYSTATUS("TPT", txtCode, "", 0, "", 0, cUName, "N", Now, Now)
        '----------------------------------
        CN.CommitTrans
        
        lstRef.AddItem UCase(TXTNAME.Text)
    Else
    CN.BeginTrans
    CN.Execute ("Update TRANSPORTMST set NAME = '" & UCase(Trim(TXTNAME.Text)) & "',ADD1 = '" & txtAdd1 & _
    "',ADD2='" & txtAdd2 & "',PH1='" & txtph1 & "',PH2='" & txtph2 & "',STAX='" & txtSTAX & _
    "',PANNO='" & TXTPAN & "',NOV='" & TXTNOV & "' where CODE ='" & Trim(txtCode.Text) & "' AND RECSTAT<>'D'")
    
    '-----------------------------------
    'DAILYSTATUS
     Call DAILYSTATUS("TPT", txtCode, "", 0, "", 0, cUName, "M", Now, Now)
    '----------------------------------
     
    CN.CommitTrans
    lstRef.Clear
    Call FillList("Select [NAME] from TRANSPORTMST WHERE RECSTAT<>'D' ORDER BY [NAME]", lstRef)
     
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

Private Sub cmdSave_GotFocus()
    Msg ActiveControl.ToolTipText
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
    
  Call FillList("Select [NAME] from TRANSPORTMST WHERE RECSTAT<>'D'", lstRef)
    
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
    txtAdd1.Enabled = Not bool
    txtAdd2.Enabled = Not bool
    txtph1.Enabled = Not bool
    txtph2.Enabled = Not bool
    txtSTAX.Enabled = Not bool
    TXTPAN.Enabled = Not bool
    TXTNOV.Enabled = Not bool
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
    TEMPRS.Open "Select * from TRANSPORTMST where [NAME] = '" & NAME & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
    If Not TEMPRS.EOF Then
    With TEMPRS
        txtCode = Trim(!CODE & "")
        TXTNAME = Trim(![NAME] & "")
        txtAdd1 = Trim(!ADD1 & "")
        txtAdd2 = Trim(!ADD2 & "")
        txtph1 = Trim(!PH1 & "")
        txtph2 = Trim(!PH2 & "")
        txtSTAX = Trim(!STAX & "")
        TXTPAN = Trim(!PANNO & "")
        TXTNOV = Trim(!Nov & "")
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

Private Sub txtSTAX_GotFocus()
    txtSTAX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtSTAX_LostFocus()
 txtSTAX.BackColor = vbWhite
End Sub

Private Sub txtNOV_GotFocus()
    TXTNOV.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtNOV_LostFocus()
 TXTNOV.BackColor = vbWhite
End Sub

Private Sub txtPAN_GotFocus()
    TXTPAN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtPAN_LostFocus()
 TXTPAN.BackColor = vbWhite
End Sub

Private Sub txtADD1_GotFocus()
    txtAdd1.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtADD1_LostFocus()
 txtAdd1.BackColor = vbWhite
End Sub

Private Sub txtADD2_GotFocus()
    txtAdd2.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtADD2_LostFocus()
 txtAdd2.BackColor = vbWhite
End Sub

Private Sub txtph1_GotFocus()
    txtph1.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtph1_LostFocus()
 txtph1.BackColor = vbWhite
End Sub

Private Sub txtph2_GotFocus()
    txtph2.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtph2_LostFocus()
 txtph2.BackColor = vbWhite
End Sub

Private Sub txtph1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtph2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtNOV_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtName_GotFocus()
    TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Transport Name"
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




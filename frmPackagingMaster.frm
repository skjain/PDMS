VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmPackagingMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packaging Master"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9120
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   6315
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11139
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
      Begin VB.TextBox txtNOP 
         Enabled         =   0   'False
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   14
         Top             =   4920
         Width           =   765
      End
      Begin VB.OptionButton OptOthers 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   4200
         Width           =   1935
      End
      Begin VB.OptionButton OptPallet 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pallet Packing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   1935
      End
      Begin VB.OptionButton OptCarton 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Carton Packing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   3240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ListBox lstRef 
         Height          =   4155
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtTwgt 
         Enabled         =   0   'False
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
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2400
         Width           =   1605
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
         TabIndex        =   9
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
         TabIndex        =   16
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin ButtonPlusCtl.ButtonPlus cmdFind 
         Height          =   375
         Left            =   7680
         TabIndex        =   8
         Top             =   4920
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
         TabIndex        =   7
         Top             =   4920
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
         Top             =   5520
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
         Image           =   "frmPackagingMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   5520
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
         Image           =   "frmPackagingMaster.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6000
         TabIndex        =   4
         Top             =   5520
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
         Image           =   "frmPackagingMaster.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   5520
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
         Image           =   "frmPackagingMaster.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   5520
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
         Image           =   "frmPackagingMaster.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   5520
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
         Image           =   "frmPackagingMaster.frx":1CAA
         cBack           =   -2147483633
      End
      Begin VB.Label lblNOP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.of PLY  "
         Enabled         =   0   'False
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
         Left            =   480
         TabIndex        =   20
         Top             =   4920
         Width           =   1140
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1575
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Packaging  Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   6135
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
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   600
         Y2              =   5400
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tare Weight     "
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
         TabIndex        =   18
         Top             =   2040
         Width           =   1530
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packaging Name     "
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
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000080&
         Height          =   345
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   195
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmPackagingMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean

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
  Call btn_sts(True) 'PACKAGING
  Call FillList("Select [NAME] from PKGNGMST where RECSTAT='A' ORDER BY [NAME]", lstRef)
  cmdExit.Cancel = True
  Me.Show
  Exit Sub

errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

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
    OPTCARTON.Value = True
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
    If ReadConfigMaster("000021", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If

    Dim ANS As String, TEMPRS As New ADODB.Recordset
    
    If isFurtherEntryExist("PACKAGING", txtCode) Then
         MsgBox "Further Entry Exist"
         Call ClsData(Me)
         lstRef.ListIndex = -1
         Call btn_sts(True)
         OPTCARTON.Value = True
         Exit Sub
    End If
    
    
    If txtCode.Text = "" Then Exit Sub
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.Title)
    If ANS = vbYes Then
       CN.Execute "UPDATE PKGNGMST SET RECSTAT='D' where CODE ='" & Trim(txtCode.Text) & "'"
       '------------------------------
       'DAILYSTATUS
       Call DAILYSTATUS("PGM", txtCode, "", 0, "", 0, cUName, "D", Now, Now)
       '------------------------------
       lstRef.RemoveItem lstRef.ListIndex
    End If
                
    Call ClsData(Me)
    lstRef.ListIndex = -1
    Call btn_sts(True)
    OPTCARTON.Value = True
End Sub

Private Sub cmdDelete_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdEdit_Click()
  
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000021", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
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
    TXTNAME.Text = SearchList1("Select TOP 20 CODE, NAME FROM PKGNGMST WHERE RECSTAT='A'", 0, "", "List Of " & Me.Caption)
    txtCode.Text = Key
    
    lstRef.Text = TXTNAME.Text
    'If cmdEdit.Enabled = True Then
    '    cmdEdit.SetFocus
    'End If
    
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
    
    If OptPallet.Value = True Then
       OptPallet.Tag = "Y"
    ElseIf OPTCARTON.Value = True Then
       OptPallet.Tag = "N"
    ElseIf OptOthers.Value = True Then
       OptPallet.Tag = "X"
    End If
        
    TXTNAME.Text = Trim(TXTNAME.Text)
    
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
           
    If Trim(TXTNAME.Text) = "" Then
        MsgBox "Please Enter Packaging Station Name.", vbInformation, App.Title
        TXTNAME = Trim(TXTNAME)
        TXTNAME.SetFocus
        Exit Sub
    End If
              
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from PKGNGMST where RECSTAT='A' AND Upper([name])='" & UCase(Trim(TXTNAME.Text)) & "' ", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = False And SAVEFLAG Then
       MsgBox "This Name Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.Title
       If TXTNAME.Enabled Then TXTNAME.SetFocus
       TEMPRS.Close
       Exit Sub
    End If
    
    If SAVEFLAG = True Then
        On Error GoTo errPRIMARYKEY
                
        txtCode.Text = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From PKGNGMST")
          
        SQL = "insert into PKGNGMST (CODE,[NAME],TWGT,PALLET,NOPLY,STATUS,RECSTAT)" _
                  & " values('" & Trim(txtCode.Text) & "','" & UCase(Trim(TXTNAME.Text)) & _
                  "','" & Val(Trim(txtTwgt.Text)) & "','" & OptPallet.Tag & "','" & Val(txtNOP) & "','A','A')"
        
        CN.BeginTrans
        CN.Execute SQL
        
        
        '------------------
        'DAILYSTAT
        Call DAILYSTATUS("PGM", txtCode, "", 0, "", 0, cUName, "N", Now, Now)
        '------------------
        CN.CommitTrans
        
        lstRef.AddItem UCase(TXTNAME.Text)
    Else
    CN.BeginTrans
    CN.Execute ("Update PKGNGMST set NAME = '" & UCase(Trim(TXTNAME.Text)) & "',NOPLY='" & Val(txtNOP) & "',TWGT = '" & txtTwgt & "',PALLET='" & OptPallet.Tag & "' where CODE ='" & Trim(txtCode.Text) & "'")
    
    '----------------------------
    'DAILYSTATUS
    Call DAILYSTATUS("PGM", txtCode, "", 0, "", 0, cUName, "M", Now, Now)
    '----------------------------
    CN.CommitTrans
    lstRef.Clear
    Call FillList("Select [NAME] from PKGNGMST where RECSTAT='A' ORDER BY [NAME]", lstRef)
     
    lstRef.ListIndex = -1
    End If
  
    Call btn_sts(True)
    sTxt = TXTNAME.Text
 
    Call ClsData(Me)
       
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    OPTCARTON.Value = True
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

Private Sub btn_sts(bool As Boolean)
    CMDSAVE.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    TXTNAME.Enabled = Not bool
    txtTwgt.Enabled = Not bool
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Msg Empty
End Sub



Private Sub lstRef_Click()
    SAVEFLAG = False
    Dim TEMPRS As New ADODB.Recordset
    If lstRef.ListIndex = -1 Then Exit Sub
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from PKGNGMST where [NAME] = '" & (lstRef.List(lstRef.ListIndex)) & "'", CN, adOpenDynamic, adLockOptimistic
    
    With TEMPRS
        txtCode.Text = !CODE & ""
        TXTNAME.Text = ![NAME] & ""
        txtTwgt = !twgt & ""
        txtNOP = Val(!NOPLY & "")
        OptPallet.Tag = !PALLET & ""
        If OptPallet.Tag = "Y" Then
           OptPallet.Value = True: OPTCARTON.Value = False: OptOthers.Value = False
        ElseIf OptPallet.Tag = "N" Then
           OptPallet.Value = False: OPTCARTON.Value = True: OptOthers.Value = False
        Else
           OptOthers.Value = True
        End If
    End With
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

Private Sub OptCarton_Click()
Call FindEnable
End Sub

Private Sub OptOthers_Click()
Call FindEnable
End Sub

Private Sub OptPallet_Click()
 Call FindEnable
End Sub

Private Sub txtName_GotFocus()
    TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Packing Station Name"
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

Private Sub txtNOP_GotFocus()
txtNOP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtNOP_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtTwgt, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNOP_LostFocus()
txtNOP.BackColor = vbWhite
End Sub

Private Sub txtTwgt_GotFocus()
  txtTwgt.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtTwgt_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtTwgt, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub FindEnable()
If OptPallet.Value = True Then
   lblNOP.Enabled = True
   txtNOP.Enabled = True
Else
   txtNOP = Empty
   lblNOP.Enabled = False
   txtNOP.Enabled = False
End If
End Sub

Private Sub txtTwgt_LostFocus()
 txtTwgt.BackColor = vbWhite
End Sub

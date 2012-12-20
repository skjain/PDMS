VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FrmSaleTaxMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAX MASTER"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9105
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   6915
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12197
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sales Tax For "
         Height          =   1095
         Left            =   240
         TabIndex        =   22
         Top             =   4800
         Width           =   5415
         Begin VB.OptionButton optLocal 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Local Sale {With in state}"
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
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   3495
         End
         Begin VB.OptionButton optInter 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Inter State Sale"
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
            TabIndex        =   24
            Top             =   480
            Width           =   3495
         End
         Begin VB.OptionButton optExport 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Export Sale {Out of country}"
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
            TabIndex        =   23
            Top             =   720
            Width           =   3495
         End
      End
      Begin VB.OptionButton optBasic 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Basic Rate"
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
         Left            =   2040
         TabIndex        =   16
         Top             =   4320
         Width           =   1575
      End
      Begin VB.OptionButton optNet 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Net Rate"
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
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txtGroup 
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
         Locked          =   -1  'True
         MaxLength       =   49
         TabIndex        =   12
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2040
         Width           =   5235
      End
      Begin VB.TextBox TXTRATECODE 
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
         Locked          =   -1  'True
         MaxLength       =   49
         TabIndex        =   14
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2880
         Width           =   5235
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
         TabIndex        =   10
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1080
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
         TabIndex        =   18
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1080
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
         Top             =   6120
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
         Image           =   "FrmSaleTaxMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   6120
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
         Image           =   "FrmSaleTaxMaster.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6000
         TabIndex        =   4
         Top             =   6120
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
         Image           =   "FrmSaleTaxMaster.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   6120
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
         Image           =   "FrmSaleTaxMaster.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   6120
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
         Image           =   "FrmSaleTaxMaster.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   6120
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
         Image           =   "FrmSaleTaxMaster.frx":1CAA
         cBack           =   -2147483633
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   5640
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Shape Shape3 
         Height          =   1335
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   5415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note : Only for, Sale Bill from challan without Order ."
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
         TabIndex        =   21
         Top             =   3480
         Width           =   5205
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Calculation Based on :"
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
         TabIndex        =   20
         Top             =   3960
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Group"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Name"
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
         Left            =   2400
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label LBLHEAD1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Master"
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
         Left            =   3600
         TabIndex        =   19
         Tag             =   "0"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   6615
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
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   600
         Y2              =   6000
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Name     "
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
         Left            =   2400
         TabIndex        =   9
         Top             =   720
         Width           =   1305
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
Attribute VB_Name = "FrmSaleTaxMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim COND As String
Dim INDEX As Long

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
    optNet.Value = True
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
    If ReadConfigMaster("000019", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    Dim ANS As String, TEMPRS As New ADODB.Recordset
    
    If isFurtherEntryExist("SALETAX", txtCode) Then
         MsgBox "Further Entry Exist"
         Call ClsData(Me)
         lstRef.ListIndex = -1
         Call btn_sts(True)
         Exit Sub
    End If
    
    If txtCode.Text = "" Then Exit Sub
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.TITLE)
    If ANS = vbYes Then
       CN.Execute "Delete from TAXMST where CODE ='" & Trim(txtCode.Text) & "'"
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
    If ReadConfigMaster("000019", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    cmdCancel.Cancel = True
    Call btn_sts(False)
    
    If lstRef.ListIndex = -1 Then lstRef.SetFocus Else TXTNAME.SetFocus
    SAVEFLAG = False
End Sub

Private Sub cmdEdit_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub CMDEXIT_Click()
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
    TXTNAME.Text = SearchList1("Select TOP 20 CODE, NAME FROM TAXMST", 0, "", "List Of " & Me.Caption)
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
    Dim TAXNO As Long
    
    If optLocal.Value = True Then
       TAXNO = 1
    ElseIf optInter.Value = True Then
       TAXNO = 2
    ElseIf optExport.Value = True Then
       TAXNO = 3
    Else
       TAXNO = 1
    End If
       
    Dim REVERSERATEREQ As String: REVERSERATEREQ = "Y"
    If optBasic.Value = True Then
       REVERSERATEREQ = "N"
    End If
    
    Dim TEMPRS As New ADODB.Recordset
    Dim Ctrl As Control
    
    TXTNAME.Text = Trim(TXTNAME.Text)
    
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
    
    If Trim(txtGroup.Text) = "" Then
        MsgBox "Please Enter Tax Group.", vbInformation, App.TITLE
        txtGroup = Trim(txtGroup)
        txtGroup.SetFocus
        Exit Sub
    End If
           
    If Trim(TXTNAME.Text) = "" Then
        MsgBox "Please Enter Sale Tax Name.", vbInformation, App.TITLE
        TXTNAME = Trim(TXTNAME)
        TXTNAME.SetFocus
        Exit Sub
    End If
    
    If Trim(TXTRATECODE.Text) = "" Then
        MsgBox "Please Enter Rate Factor.", vbInformation, App.TITLE
        TXTRATECODE = Trim(TXTRATECODE)
        TXTRATECODE.SetFocus
        Exit Sub
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM TAXGRPMST WHERE NAME ='" & txtGroup & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       txtGroup.Tag = Trim(RS!CODE & "")
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM RATEMST WHERE NAME='" & Trim(TXTRATECODE) & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       TXTRATECODE.Tag = Trim(RS!CODE & "")
    Else
       TXTRATECODE.Tag = Empty
       TXTRATECODE = Empty
       If TXTRATECODE.Enabled Then TXTRATECODE.SetFocus
    End If
                                             
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from TAXMST where Upper([name])='" & UCase(Trim(TXTNAME.Text)) & "' ", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = False And SAVEFLAG Then
       MsgBox "This Name Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.TITLE
       TEMPRS.Close
       Exit Sub
    End If
    
    If SAVEFLAG = True Then
        On Error GoTo errPRIMARYKEY
        
        txtCode.Text = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From TAXMST ")
          
        SQL = "INSERT INTO TAXMST (CODE,[NAME],GRPCOD,RECSTAT,RATE_CODE,REVERSERATEREQ,TAXNO) VALUES('" & txtCode & _
        "','" & UCase(Trim(TXTNAME.Text)) & "','" & UCase(Trim(txtGroup.Tag)) & _
        "','A','" & TXTRATECODE.Tag & "','" & REVERSERATEREQ & "'," & TAXNO & ")"
        
        CN.BeginTrans
        CN.Execute SQL
        CN.CommitTrans
        lstRef.AddItem UCase(TXTNAME.Text)
        
    Else
    
    CN.BeginTrans
    
    SQL = "Update TAXMST set NAME = '" & UCase(Trim(TXTNAME.Text)) & "',RATE_CODE = '" & TXTRATECODE.Tag & _
    "',GRPCOD = '" & txtGroup.Tag & "',REVERSERATEREQ = '" & REVERSERATEREQ & _
    "',TAXNO = " & TAXNO & " WHERE CODE ='" & Trim(txtCode.Text) & "'"
    
    CN.Execute SQL
    CN.CommitTrans
    lstRef.Clear
    Call FillList("Select [NAME] from TAXMST where RECSTAT='A' ORDER BY [NAME]", lstRef)
     
    lstRef.ListIndex = -1
    End If
    
    optNet.Value = True
    Call btn_sts(True)
    
    sTxt = TXTNAME.Text
 
    Call ClsData(Me)
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub

errPRIMARYKEY:
    CN.RollbackTrans
    If ERR.Number = -2147217873 Or -2147217900 Then
        TXTNAME.SetFocus
        MsgBox "This Name Already Registered With Other Category!!!", vbInformation, "Already Registered"
    Else
        ErrNumber = ERR.Number
        ErrMessage = ERR.Description
        frm_ErrorHandler.Show vbModal
    End If
    ERR.Clear
End Sub

Private Sub cmdSave_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  If cmdAdd.Enabled Then cmdAdd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
On Error GoTo errLoad
  Call btn_sts(True)
  Call FillList("Select [NAME] from TAXMST WHERE RECSTAT='A' ORDER BY [NAME]", lstRef)
  cmdExit.Cancel = True
  Me.Show
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
    TXTNAME.Enabled = Not bool
    cmdFind.Enabled = Not bool
    cmdClear.Enabled = Not bool
    lstRef.Enabled = Not bool
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Msg Empty
End Sub

Private Sub lstRef_Click()
Dim i As Long, J As Long
    SAVEFLAG = False
    Dim SAVDAT As New ADODB.Recordset
    Set SAVDAT = New ADODB.Recordset
    
    If lstRef.ListIndex = -1 Then Exit Sub
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "Select * from TAXMST where [NAME] = '" & (lstRef.List(lstRef.ListIndex)) & "'", CN, adOpenDynamic, adLockOptimistic
    
    With SAVDAT
        txtCode.Text = !CODE & ""
        TXTNAME.Text = ![NAME] & ""
        
        If !REVERSERATEREQ & "" = "Y" Then
          optNet.Value = True
        Else
          optBasic.Value = True
        End If
        
    If Val(!TAXNO) = 1 Then
       optLocal.Value = True
    ElseIf Val(!TAXNO) = 2 Then
       optInter.Value = True
    ElseIf Val(!TAXNO) = 3 Then
       optExport.Value = True
    Else
       optLocal.Value = True
    End If
        
        
        TXTRATECODE.Text = GetCode("RATEMST", Trim(![RATE_CODE] & ""), "CODE", "NAME")
        txtGroup.Text = GetCode("TAXGRPMST", Trim(![GRPCOD] & ""), "CODE", "NAME")
    End With
    SAVDAT.Close
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

Private Sub txtGroup_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(txtGroup) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtGroup.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM TAXGRPMST WHERE RECSTAT='A'", 0, txtGroup.Text, "SELECT TAX GROUP FROM LIST")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            txtGroup.Text = ""
            frmTaxGroupMaster.Show
        Else
            txtGroup.Tag = Key
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTRATECODE_GotFocus()
TXTRATECODE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRATECODE_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(TXTRATECODE) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTRATECODE.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM RATEMST WHERE RECSTAT='A'", 0, TXTRATECODE.Text, "SELECT RATE FROM LIST")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTRATECODE.Text = ""
            frmRatFactMst.Show
        Else
            TXTRATECODE.Tag = Key
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTRATECODE_LostFocus()
TXTRATECODE.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
    TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    TXTNAME.SelStart = 0
    TXTNAME.SelLength = Len(TXTNAME)
End Sub

Private Sub txtName_LostFocus()
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

Private Sub txtGroup_GotFocus()
    txtGroup.BackColor = RGB(BRED, BGREEN, BBLUE)
    txtGroup.SelStart = 0
    txtGroup.SelLength = Len(txtGroup)
End Sub

Private Sub txtGroup_LostFocus()
  txtGroup.BackColor = vbWhite
End Sub


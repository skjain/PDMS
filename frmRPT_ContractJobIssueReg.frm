VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_ContractJobIssueReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Register"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6795
   Begin VB.Frame frameStatus 
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   6615
      Begin VB.OptionButton optPending 
         Caption         =   "Pending"
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
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optClear 
         Caption         =   "Clear"
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
         Left            =   4680
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
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
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "&Jobwork Status   "
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   6615
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2760
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Pre&view"
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
         Image           =   "frmRPT_ContractJobIssueReg.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5040
         TabIndex        =   20
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "frmRPT_ContractJobIssueReg.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label13 
         Caption         =   "R&eport Zoom %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   6615
      Begin VB.TextBox txtITEM 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtParty 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label7 
         Caption         =   "&Item Name       "
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "&Party Name     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   6615
      Begin VB.ComboBox cboReports 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   4815
      End
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/MM/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox dtTo 
         Height          =   330
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/MM/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "&From Date                "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "&To Date       "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "&Report Format "
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
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtUNIT 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "&Unit Name              "
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRPT_ContractJobIssueReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RPTN As String
Dim m_unit As String
Dim L_CUNT As String
Dim ORDBOK As String
Dim ORDDBC As String
Dim sel_untcod As String
Dim SEL_DVCDNAM As String
Dim SEL_DVCDCOD As String
Dim M_DVCD As String

Private Sub cboReports_GotFocus()
  cboReports.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboReports_KeyPress(KeyAscii As Integer)
If KeyAscii <= 90 Or KeyAscii <= 122 Or KeyAscii <= 57 Or KeyAscii <= 46 Or KeyAscii <= 47 Then
  KeyAscii = 0
End If
End Sub

Private Sub cboReports_LostFocus()
 cboReports.BackColor = vbWhite
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
Me.KeyPreview = False
On Error GoTo errPreview
    
    If cboReports.ListIndex = -1 Then
        MsgBox "Please Select Report Format ", vbInformation
        cboReports.SetFocus
        SendKeys "{DOWN}"
        Exit Sub
    End If
    
    If txtUNIT = Empty Then
       MsgBox "Please Select Unit", vbInformation
       txtUNIT.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(dtFrom) Then
       MsgBox "Please Select Correct Starting Date", vbInformation
       dtFrom.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(dtTo) Then
       MsgBox "Please Select Correct Ending Date", vbInformation
       dtTo.SetFocus
       Exit Sub
    End If
               
    CRPT.Reset
    crptConnect CRPT
     
            
    Call SetReportName
    Call SetSQL
              
    Debug.Print ReportName
    
    If ReportName = Empty Then
        MsgBox "No Report Design For Selected Criteria !!", vbInformation, "Under Development"
        Exit Sub
    End If
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
      
     
    PERIOD = dtFrom & " To " & dtTo
    
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
    
    .Formulas(1) = "COMPANY='" & compNm & "'"
    .Formulas(2) = "UNIT='" & txtUNIT & "'"
    .Formulas(3) = "REPORTHEAD='" & RPTN & "'"
    .Formulas(4) = "PERIOD='" & PERIOD & "'"
    If Me.Tag = "JLD" Or Me.Tag = "PLD" Then
        .Formulas(5) = "OPDT=#" & Format(DateAdd("D", -1, dtFrom), "MM/dd/yyyy") & "#"
    End If
               
    
    If ReportName = App.PATH & "\Reports\JOBWORK_DATEWISEREG.RPT" Or ReportName = App.PATH & "\Reports\JOBWORK_PARTYWISEREG.RPT" Then
        If Me.Tag = "XXX" Or Me.Tag = "JLD" Then
           CRPT.Formulas(6) = "TYPE='ANX'"
        Else
           CRPT.Formulas(6) = "TYPE='RGP'"
        End If
    End If
        
    RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        Select Case frmRPT_ContractJobIssueReg.Tag
        
        Case "NGP"
        
        If ReadConfigMaster("000049", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        
        Case "YYY'"
        
        If ReadConfigMaster("000050", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        Case "RGP"
        If ReadConfigMaster("000051", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        
        Case "IVR4"
        If ReadConfigMaster("000052", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        
        Case "PLD"
        
        If ReadConfigMaster("000053", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        
        Case "XXX"
        If ReadConfigMaster("000054", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        
        Case "ANX"
         If ReadConfigMaster("000055", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        Case "IVR3"
        If ReadConfigMaster("000056", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
      Case "JLD"
         If ReadConfigMaster("000057", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        End Select
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
         txtUNIT.SetFocus
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    Exit Sub

errPreview:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub dtFrom_GotFocus()
  dtFrom.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub dtFrom_LostFocus()
   dtFrom.BackColor = vbWhite
End Sub

Private Sub dtTo_GotFocus()
   dtTo.BackColor = RGB(BRED, BGREEN, BBLUE)
   SendKeys "{HOME}+{END}"
End Sub

Private Sub dtTo_LostFocus()
   dtTo.BackColor = vbWhite
End Sub

Private Sub Form_Activate()
 txtUnit_KeyDown vbKeyReturn, 0
 If cboReports.Text = Empty And cboReports.ListCount = 0 Then
    Call SetCombo
 End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtFrom.Text = Format(FSDT, "dd/MM/yyyy")
    dtTo.Text = Format(FEDT, "dd/MM/yyyy")
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub txtParty_GotFocus()
 txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtParty.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM ACCMST", 0, "", "List Of Jober Name")
        txtParty.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtParty = Empty
        txtParty.Tag = Empty
    End If
End Sub

Private Sub txtParty_LostFocus()
   txtParty.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
  txtITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtITEM = SearchList1("Select TOP 20 Code,Name From ITMMST ", 0, Empty, "Select Item")
        txtITEM.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtITEM = Empty
        txtITEM.Tag = Empty
    End If
End Sub

Private Sub txtItem_LostFocus()
 txtITEM.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
  txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    End If
End Sub

Private Sub txtUNIT_LostFocus()
txtUNIT.BackColor = vbWhite
End Sub

Private Sub TXTZOOM_GotFocus()
txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTZOOM_LostFocus()
txtZoom.BackColor = vbWhite
End Sub


Private Sub SetCombo()
With cboReports
Select Case Me.Tag
Case "ANX"
   .AddItem "DATEWISE JOBWORK ISSUE REGISTER"
   .AddItem "PARTYWISE JOBWORK ISSUE REGISTER"
   .AddItem "ITEMWISE JOBWORK ISSUE REGISTER"
   Me.Caption = "Job Issue Register"
Case "IVR3"
   .AddItem "DATEWISE JOBWORK RECEIVE REGISTER"
   .AddItem "PARTYWISE JOBWORK RECEIVE REGISTER"
   .AddItem "ITEMWISE JOBWORK RECEIVE REGISTER"
   Me.Caption = "Job Receive Register"
   frameStatus.Enabled = False
Case "RGP"
   .AddItem "DATEWISE RETURNABLE ISSUE REGISTER"
   .AddItem "PARTYWISE RETURNABLE ISSUE REGISTER"
   .AddItem "ITEMWISE RETURNABLE ISSUE REGISTER"
   Me.Caption = "Returnable Issue Register"
Case "NGP"
   .AddItem "DATEWISE NON-RETURNABLE ISSUE REGISTER"
   .AddItem "PARTYWISE NON-RETURNABLE ISSUE REGISTER"
   .AddItem "ITEMWISE NON-RETURNABLE ISSUE REGISTER"
   Me.Caption = "Non-Returnable Issue Register"
   frameStatus.Enabled = False
Case "IVR4"
   .AddItem "DATEWISE RETURNABLE RECEIVE REGISTER"
   .AddItem "PARTYWISE RETURNABLE RECEIVE REGISTER"
   .AddItem "ITEMWISE RETURNABLE RECEIVE REGISTER"
   Me.Caption = "Returnable Receive Register"
   frameStatus.Enabled = False
Case "XXX"
   .AddItem "DATEWISE JOBWORK REGISTER"
   .AddItem "PARTYWISE JOBWORK REGISTER"
   Me.Caption = "JOBWORK REGISTER "
Case "YYY"
   .AddItem "DATEWISE RETURNABLE REGISTER"
   .AddItem "PARTYWISE RETURNABLE REGISTER"
   Me.Caption = "RETURNABLE REGISTER "
Case "JLD"
   .AddItem "JOBERWISE ITEM DETAIL LEDGER"
   .AddItem "JOBERWISE ITEM SUMMARY LEDGER"
   Me.Caption = "JOBERWISE ITEM LEDGER"
Case "PLD"
   .AddItem "PARTYWISE ITEM DETAIL LEDGER"
   .AddItem "PARTYWISE ITEM SUMMARY LEDGER"
   Me.Caption = "PARTYWISE ITEM LEDGER"
End Select
End With
 cboReports.ListIndex = 0
End Sub

Private Sub SetReportName()
ReportName = Empty
RPTN = Empty

Select Case Me.Tag
Case "ANX", "RGP", "NGP"
     If cboReports.ListIndex = 0 Then
        ReportName = App.PATH & "\Reports\JOB_ISSUE_REGISTER_DATEWISE.RPT"
        If Me.Tag = "ANX" Then RPTN = "DATEWISE JOBWORK ISSUE REGISTER"
        If Me.Tag = "RGP" Then RPTN = "DATEWISE RETURNABLE ISSUE REGISTER"
        If Me.Tag = "NGP" Then RPTN = "DATEWISE NON-RETURNABLE ISSUE REGISTER"
     ElseIf cboReports.ListIndex = 1 Then
        ReportName = App.PATH & "\Reports\JOB_ISSUE_REGISTER_PARTYWISE.RPT"
        If Me.Tag = "ANX" Then RPTN = "PARTYWISE JOBWORK ISSUE REGISTER"
        If Me.Tag = "RGP" Then RPTN = "PARTYWISE RETURNABLE ISSUE REGISTER"
     ElseIf cboReports.ListIndex = 2 Then
        ReportName = App.PATH & "\Reports\JOB_ISSUE_REGISTER_ITEMWISE.RPT"
        If Me.Tag = "ANX" Then RPTN = "ITEMWISE JOBWORK ISSUE REGISTER"
        If Me.Tag = "RGP" Then RPTN = "ITEMWISE RETURNABLE ISSUE REGISTER"
     End If
Case "IVR3", "IVR4"
     If cboReports.ListIndex = 0 Then
        ReportName = App.PATH & "\Reports\JOB_RECEIVE_REGISTER_DATEWISE.RPT"
        If Me.Tag = "IVR3" Then RPTN = "DATEWISE JOBWORK RECEIVE REGISTER"
        If Me.Tag = "IVR4" Then RPTN = "DATEWISE RETURNABLE RECEIVE REGISTER"
     ElseIf cboReports.ListIndex = 1 Then
        ReportName = App.PATH & "\Reports\JOB_RECEIVE_REGISTER_PARTYWISE.RPT"
        If Me.Tag = "IVR3" Then RPTN = "PARTYWISE JOBWORK RECEIVE REGISTER"
        If Me.Tag = "IVR4" Then RPTN = "PARTYWISE RETURNABLE RECEIVE REGISTER"
     ElseIf cboReports.ListIndex = 2 Then
        ReportName = App.PATH & "\Reports\JOB_RECEIVE_REGISTER_ITEMWISE.RPT"
        If Me.Tag = "IVR3" Then RPTN = "ITEMWISE JOBWORK RECEIVE REGISTER"
        If Me.Tag = "IVR4" Then RPTN = "ITEMWISE RETURNABLE RECEIVE REGISTER"
     End If
Case "XXX", "YYY"
     If cboReports.ListIndex = 0 Then
        ReportName = App.PATH & "\Reports\JOBWORK_DATEWISEREG.RPT"
        If Me.Tag = "XXX" Then RPTN = "DATEWISE JOBWORK REGISTER"
        If Me.Tag = "YYY" Then RPTN = "DATEWISE RETURNABLE REGISTER"
     ElseIf cboReports.ListIndex = 1 Then
        ReportName = App.PATH & "\Reports\JOBWORK_PARTYWISEREG.RPT"
        If Me.Tag = "XXX" Then RPTN = "PARTYWISE JOBWORK REGISTER"
        If Me.Tag = "YYY" Then RPTN = "PARTYWISE RETURNABLE REGISTER"
     End If
Case "JLD", "PLD"
     If cboReports.ListIndex = 0 Then
        ReportName = App.PATH & "\Reports\Joberwise_ItemDetail_Ledger.rpt"
        If Me.Tag = "JLD" Then RPTN = "JOBERWISE ITEM DETAIL LEDGER"
        If Me.Tag = "PLD" Then RPTN = "PARTYWISE ITEM DETAIL LEDGER"
     ElseIf cboReports.ListIndex = 1 Then
        ReportName = App.PATH & "\Reports\Joberwise_ItemSummary_Ledger.rpt"
        If Me.Tag = "JLD" Then RPTN = "JOBERWISE ITEM SUMMARY LEDGER"
        If Me.Tag = "PLD" Then RPTN = "PARTYWISE ITEM SUMMARY LEDGER"
     End If
    
End Select
End Sub

Private Sub SetSQL()
rptsql = Empty
Select Case Me.Tag
Case "XXX", "YYY", "JLD", "PLD"
    rptsql = "{JOBTRACK.COMP}='" & compPth & "' AND {JOBTRACK.UNIT} = '" & txtUNIT.Tag & "' "
    
    If Me.Tag = "JLD" Or Me.Tag = "PLD" Then
       rptsql = rptsql & " AND {JOBTRACK.DATE} <= DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
    Else
       rptsql = rptsql & " AND {JOBTRACK.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & _
       ") AND {JOBTRACK.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
    End If
    
    Exit Sub
    
    If Me.Tag = "XXX" Or Me.Tag = "JLD" Then
       rptsql = rptsql & " AND ({JOBTRACK.VTYP}='ANX' AND {JOBTRACK.DBCD}='000001') OR ({JOBTRACK.VTYP}='XXX' AND {JOBTRACK.DBCD}='000003')"
    Else
       rptsql = rptsql & " AND ({JOBTRACK.VTYP}='RGP' AND {JOBTRACK.DBCD}='000001') OR ({JOBTRACK.VTYP}='XXX' AND {JOBTRACK.DBCD}='000004')"
    End If
    
    If txtParty <> Empty Then rptsql = rptsql & " AND {JOBTRACK.PARTY}='" & txtParty & "'"
    If txtITEM <> Empty Then rptsql = rptsql & " AND {JOBOUT.ITEM}='" & txtITEM & "'"
    If optPending.Value = True Then rptsql = rptsql & " AND {JOBTRACK.CLRSTATUS}='N'"
    If optClear.Value = True Then rptsql = rptsql & " AND {JOBTRACK.CLRSTATUS}='Y'"
    
    'Exit Sub
    
End Select

rptsql = "{JOBOUT.COMP}='" & compPth & "' AND {JOBOUT.UNIT} = '" & txtUNIT.Tag & _
"' AND {JOBOUT.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & _
") AND {JOBOUT.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {JOBOUT.RECSTAT}<>'D' "

Select Case Me.Tag
Case "ANX"
  rptsql = rptsql & " AND {JOBOUT.VTYP}='ANX' AND {JOBOUT.DBCD}='000001'"
  If optPending.Value = True Then rptsql = rptsql & " AND {JOBOUT.CLRSTATUS}='N'"
  If optClear.Value = True Then rptsql = rptsql & " AND {JOBOUT.CLRSTATUS}='Y'"
Case "IVR3"
  rptsql = rptsql & " AND {JOBOUT.VTYP}='IVR' AND {JOBOUT.DBCD}='000003' "
Case "RGP"
  rptsql = rptsql & " AND {JOBOUT.VTYP}='RGP' AND {JOBOUT.DBCD}='000001' "
  If optPending.Value = True Then rptsql = rptsql & " AND {JOBOUT.CLRSTATUS}='N'"
  If optClear.Value = True Then rptsql = rptsql & " AND {JOBOUT.CLRSTATUS}='Y'"
Case "NGP"
  rptsql = rptsql & " AND {JOBOUT.VTYP}='NGP' AND {JOBOUT.DBCD}='000001' "
Case "IVR4"
  rptsql = rptsql & " AND {JOBOUT.VTYP}='IVR' AND {JOBOUT.DBCD}='000004' "
End Select

    If txtParty <> Empty Then rptsql = rptsql & " AND {JOBOUT.PCOD}='" & txtParty.Tag & "'"
            
    If txtITEM <> Empty Then rptsql = rptsql & " AND {JOBOUT.ICOD}='" & txtITEM.Tag & "'"
End Sub

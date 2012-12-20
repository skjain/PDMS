VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_OrderReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Register"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6075
   Begin VB.Frame framMachine 
      Height          =   525
      Left            =   120
      TabIndex        =   34
      Top             =   4920
      Width           =   5925
      Begin VB.OptionButton optUnAprv 
         Caption         =   "UnApproved"
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
         Left            =   2880
         TabIndex        =   25
         Top             =   240
         Width           =   1575
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
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optAprv 
         Caption         =   "Approved"
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
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCancel 
         Caption         =   "Cancelled"
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
         Left            =   4560
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   33
      Top             =   5520
      Width           =   5895
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2640
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3240
         TabIndex        =   29
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
         Image           =   "frmRPT_OrderReg.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4560
         TabIndex        =   30
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
         Image           =   "frmRPT_OrderReg.frx":0452
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
         TabIndex        =   27
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   5895
      Begin VB.TextBox txtSM 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox PTYORD 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   22
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtTXCD 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txtITEM 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtBroker 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtPCOD 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label11 
         Caption         =   "&Sales Man       "
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
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Tax Category "
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
         TabIndex        =   19
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Pa&rty Order      "
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
         TabIndex        =   21
         Top             =   2040
         Width           =   1815
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
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "&Agent Name     "
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
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Party &Name       "
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
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   31
      Top             =   1320
      Width           =   5895
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
         TabIndex        =   10
         Top             =   720
         Width           =   4095
      End
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1560
         TabIndex        =   6
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
         Left            =   4320
         TabIndex        =   8
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
         TabIndex        =   5
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
         Left            =   3360
         TabIndex        =   7
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
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtDVCD 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtUNIT 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "&Division        "
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
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&Unit               "
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
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRPT_OrderReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RPTN As String
Dim m_unit As String
Dim ORDBOK As String
Dim ORDDBC As String
Dim sel_untcod As String
Dim SEL_DVCDNAM As String
Dim SEL_DVCDCOD As String
Dim M_DVCD As String

Private Sub cboReports_GotFocus()
cboReports.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboReports_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And cboReports <> Empty Then txtSM.SetFocus
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
On Error GoTo errPreview
    
    If cboReports.ListIndex = -1 Then
        MsgBox "Please Select Report Format ", vbInformation
        cboReports.SetFocus
        SendKeys "{DOWN}"
        Exit Sub
    End If
    If txtUNIT = Empty Then
       MsgBox "Pleas Select Unit", vbInformation
       txtUNIT.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(dtFrom) Then
       MsgBox "Pleas Select Correct Starting Date", vbInformation
       dtFrom.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(dtTo) Then
       MsgBox "Pleas Select Correct Ending Date", vbInformation
       dtTo.SetFocus
       Exit Sub
    End If
                
    rptsql = Empty
    RPTN = Empty
    CRPT.Reset
    crptConnect CRPT
    
 
    rptsql = "{ORDMAN.COMP}='" & compPth & "' AND {ORDMAN.UNIT} = '" & txtUNIT.Tag & "' AND {ORDMAN.ORDT}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {ORDMAN.ORDT}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") And {ORDMAN.RECSTAT}<>'D'"
    
    If txtDVCD <> Empty Then
       rptsql = rptsql & " AND {ORDMAN.DCOD} = '" & txtDVCD.Tag & "'"
    End If
    
    If optAll.Value <> True Then
       If optAprv.Value = True Then
          rptsql = rptsql & " AND {ORDMAN.FIN_APRV}='O' "
       ElseIf optUnAprv.Value = True Then
          rptsql = rptsql & " AND {ORDMAN.FIN_APRV}='N' "
       ElseIf optCancel.Value = True Then
          rptsql = rptsql & " AND {ORDMAN.FIN_APRV}='C' "
       End If
    End If
        
    If txtPCOD <> Empty Then rptsql = rptsql & " AND {ORDMAN.PCOD}='" & txtPCOD.Tag & "'"
    If txtBroker <> Empty Then rptsql = rptsql & " AND {ORDMAN.BRCD}='" & txtBroker.Tag & "'"
    If txtTXCD <> Empty Then rptsql = rptsql & " AND {ORDMAN.txcd}='" & txtTXCD.Tag & "'"
    If txtITEM <> Empty Then rptsql = rptsql & " AND {ORDMAN.ICOD}='" & txtITEM.Tag & "'"
    If PTYORD <> Empty Then rptsql = rptsql & " AND {ORDMAN.ORDN}='" & PTYORD.Text & "'"
    If txtSM <> Empty Then rptsql = rptsql & " AND {ORDMAN.DBCD} = '" & txtSM.Tag & "'"
    
    ReportName = Empty
    
    If cboReports.ListIndex = 0 Then
            ReportName = App.PATH & "\Reports\ORDREG_DATEWISE_ALL.RPT"
            RPTN = "DATEWISE SALES ORDER REGISTER"
    ElseIf cboReports.ListIndex = 1 Then
            ReportName = App.PATH & "\Reports\ORDREG_PARTYWISE_ALL.RPT"
            RPTN = "PARTYWISE SALES ORDER REGISTER"
    ElseIf cboReports.ListIndex = 2 Then
            ReportName = App.PATH & "\Reports\ORDREG_AGENTWISE_ALL.RPT"
            RPTN = "AGENTWISE SALES ORDER REGISTER"
            
    ElseIf cboReports.ListIndex = 3 Then
            ReportName = App.PATH & "\Reports\ORDREG_ITEMWISE_PENDING.RPT"
            RPTN = "ITEMWISE SALES ORDER REGISTER"
    End If
    
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
    
    
     If PTYORD <> Empty Then
      RPTN = RPTN + " Party Order No. : " + PTYORD.Text
     Else
      RPTN = RPTN
     End If
     
     If optAll.Value <> True Then
       If optAprv.Value = True Then
          RPTN = RPTN + " [Approved ]"
       ElseIf optUnAprv.Value = True Then
          RPTN = RPTN + " [UnApproved ]"
       ElseIf optCancel.Value = True Then
          RPTN = RPTN + " [Cancelled ]"
       End If
    End If
          
       PERIOD = dtFrom & " To " & dtTo
    
     CRPT.ReplaceSelectionFormula rptsql
    
     With CRPT
       .Formulas(1) = "COMPANY='" & compNm & "'"
       .Formulas(2) = "UNIT='" & txtUNIT & "'"
       .Formulas(3) = "REPORTHEAD='" & RPTN & "'"
       .Formulas(4) = "PERIOD='" & PERIOD & "'"
       .Formulas(5) = "LBLVIEW='" & LabelDisplay(txtDVCD.Tag, txtUNIT.Tag) & "'"
       
         RPTN = RPTN + Space(5) + ReportName
       .DiscardSavedData = True
       .WindowTitle = RPTN
       .Destination = crptToWindow
       .WindowState = crptMaximized
       .WindowShowProgressCtls = True
        
        If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000030", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        
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

Private Sub dtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then dtTo.SetFocus
End Sub

Private Sub dtFrom_LostFocus()
   dtFrom.BackColor = vbWhite
End Sub

Private Sub dtTo_GotFocus()
   dtTo.BackColor = RGB(BRED, BGREEN, BBLUE)
   SendKeys "{HOME}+{END}"
End Sub

Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboReports.SetFocus
End Sub

Private Sub dtTo_LostFocus()
   dtTo.BackColor = vbWhite
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 txtUnit_KeyDown vbKeyReturn, 0
 cboReports.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
 
    Call CenterChild(frm_Main, Me)
    
    dtFrom.Text = Format(FSDT, "dd/MM/yyyy")
    dtTo.Text = Format(FEDT, "dd/MM/yyyy")
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    
    With cboReports
        .AddItem "DATEWISE ORDER REGISTER"
        .AddItem "PARTYWISE ORDER REGISTER"
        .AddItem "AGENTWISE ORDER REGISTER"
        .AddItem "ITEMWISE ORDER REGISTER"
        .ListIndex = 0
    End With
    cboReports.ListIndex = 0
    
End Sub

Private Sub ptyord_GotFocus()
PTYORD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub PTYORD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtZoom.SetFocus
End Sub

Private Sub ptyord_LostFocus()
 PTYORD.BackColor = vbWhite
End Sub

Private Sub txtBroker_GotFocus()
 txtBroker.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtBroker_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtBroker.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM REFMST WHERE CATA='B'", 0, "", "List Of Agents")
        txtBroker.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtBroker = Empty
        txtBroker.Tag = Empty
    End If
     If KeyCode = vbKeyReturn Then txtITEM.SetFocus
End Sub

Private Sub txtBroker_LostFocus()
   txtBroker.BackColor = vbWhite
End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "'  AND RECSTAT='A' AND CODE<>'000001'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
        
        
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
    If KeyCode = vbKeyReturn And txtDVCD <> Empty Then dtFrom.SetFocus
End Sub

Private Sub txtDVCD_LostFocus()
txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
 txtITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtITEM = SearchList1("Select TOP 20 Code,Name From FINITMMST Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "' AND DVCD = '" & txtDVCD.Tag & "'", 0, Empty, "Select Item")
        txtITEM.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtITEM = Empty
        txtITEM.Tag = Empty
    End If
    If KeyCode = vbKeyReturn Then txtTXCD.SetFocus
End Sub

Private Sub txtItem_LostFocus()
 txtITEM.BackColor = vbWhite
End Sub

Private Sub txtPCOD_GotFocus()
 txtPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtPCOD = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select Party From List")
        txtPCOD.Tag = Key
        If txtPCOD <> Empty Then txtBroker.SetFocus
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtPCOD = Empty
        txtPCOD.Tag = Empty
    End If
    If KeyCode = vbKeyReturn Then txtBroker.SetFocus
End Sub

Private Sub txtPCOD_LostFocus()
 txtPCOD.BackColor = vbWhite
End Sub

Private Sub txtSM_GotFocus()
txtSM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtSM_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtSM = SearchList1("Select  TOP 20 Code,Name From SALMANMST", 0, Empty, "Select Sales Man From List")
        txtSM.Tag = Key
        If txtSM <> Empty Then txtPCOD.SetFocus
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtSM = Empty
        txtSM.Tag = Empty
    End If
    
If KeyCode = vbKeyReturn And cboReports <> Empty Then txtPCOD.SetFocus
End Sub

Private Sub txtSM_LostFocus()
txtSM.BackColor = vbWhite
End Sub

Private Sub txtTXCD_GotFocus()
 txtTXCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txttxcd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtTXCD.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM TAXMST ", 0, "", "List Of Tax Category")
        txtTXCD.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtTXCD = Empty
        txtTXCD.Tag = Empty
    End If
If KeyCode = vbKeyReturn Then PTYORD.SetFocus
End Sub

Private Sub txtTXCD_LostFocus()
 txtTXCD.BackColor = vbWhite
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
    If KeyCode = vbKeyReturn And txtUNIT <> Empty Then txtDVCD.SetFocus
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
'If KeyAscii <= 90 Or KeyAscii <= 122 Or KeyAscii <= 57 Or KeyAscii <= 46 Or KeyAscii <= 47 Then
'  KeyAscii = 0
'End If
End Sub

Private Sub txtUNIT_LostFocus()
txtUNIT.BackColor = vbWhite
End Sub

Private Sub TXTZOOM_GotFocus()
txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtZoom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdpreview.SetFocus
End Sub

Private Sub TXTZOOM_LostFocus()
txtZoom.BackColor = vbWhite
End Sub

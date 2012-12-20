VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form FRMRPT_BOXSTOCK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box Wise Report"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6825
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   6615
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   12
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
         TabIndex        =   13
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
         Image           =   "FRMRPT_BOXSTOCK.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "FRMRPT_BOXSTOCK.frx":0452
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
         TabIndex        =   15
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtUNIT 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker dtopdt 
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   18350081
         CurrentDate     =   40960
      End
      Begin VB.Label Label2 
         Caption         =   "As On Date"
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
         Top             =   840
         Width           =   1095
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
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame framMachine 
      Height          =   1275
      Left            =   120
      TabIndex        =   3
      Top             =   2520
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
         Height          =   195
         Left            =   4920
         TabIndex        =   7
         Top             =   840
         Width           =   1335
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
         Left            =   600
         TabIndex        =   6
         Top             =   840
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optIssue 
         Caption         =   "Issue"
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
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox ChkSummary 
         Caption         =   "Summary "
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
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   120
         Top             =   240
         Width           =   6375
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6480
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   1920
         Y1              =   720
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   4200
         X2              =   4200
         Y1              =   720
         Y2              =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6615
      Begin VB.TextBox TXTPCOD 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Party Name"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRMRPT_BOXSTOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
    On Error GoTo errPreview
    
    If txtUNIT = Empty Then
       MsgBox "Please Select Unit", vbInformation
       txtUNIT.SetFocus
       Exit Sub
    End If
           
                  
    CRPT.Reset
    crptConnect CRPT
     
            
    ReportName = Empty
    RPTN = Empty
    rptsql = Empty
    
    rptsql = "{TRDBOXREGISTER.COMP}='" & compPth & "' AND {TRDBOXREGISTER.UNIT} = '" & txtUNIT.Tag & "' "
    
    'If txtltno <> Empty Then rptsql = rptsql & " AND {JUMBOREG.SETNO}='" & txtltno & "' "
    
    If TXTPCOD <> Empty Then rptsql = rptsql & " AND {TRDBOXREGISTER.PCOD}='" & TXTPCOD.Tag & "'"
    If ChkSummary.Value = False Then
    If optAll = True Then
        ReportName = App.PATH & "\Reports\Itemwise BoxStock.rpt"
        rptsql = rptsql & " AND {TRDBOXREGISTER.VTYP} = 'IVR'"
        RPTN = "ITEM WISE BOX REPORT {All}"
        
    ElseIf optIssue.Value = True Then
        rptsql = rptsql & " AND ISNULL({TRDBOXREGISTER.RVBNO}) = TRUE AND ({TRDBOXREGISTER.VTYP} = 'ISS' OR  {TRDBOXREGISTER.VTYP} = 'SAL')"
        ReportName = App.PATH & "\Reports\Itemwise BoxStock.rpt"
        RPTN = "ITEM WISE BOX ISSUE REPORT {Issue}"
    
    ElseIf optPending.Value = True Then
        rptsql = rptsql & " AND  {TRDBOXREGISTER.OPER} = '+' AND ISNULL({TRDBOXREGISTER.RVTYP}) = TRUE "
        ReportName = App.PATH & "\Reports\Itemwise BoxStock.rpt"
        RPTN = "ITEM WISE BOX PENDING REPORT {Pending}"
    End If
    
Else
        
    If optAll = True Then
        ReportName = App.PATH & "\Reports\Itemwise BoxStockSummary.rpt"
        rptsql = rptsql & " AND {TRDBOXREGISTER.VTYP} = 'IVR'"
        RPTN = "ITEM WISE BOX  STOCK REPORT {All}"
        
    ElseIf optIssue.Value = True Then
        rptsql = rptsql & " AND ISNULL({TRDBOXREGISTER.RVBNO}) = TRUE AND ({TRDBOXREGISTER.VTYP} = 'ISS' OR {TRDBOXREGISTER.VTYP} = 'SAL')"
        ReportName = App.PATH & "\Reports\Itemwise BoxStockSummary.rpt"
        RPTN = "ITEM WISE BOX ISSUE REPORT {Issue}"
    
    ElseIf optPending.Value = True Then
        rptsql = rptsql & " AND ISNULL({TRDBOXREGISTER.RVBNO}) = TRUE AND {TRDBOXREGISTER.OPER} = '+' AND ISNULL({TRDBOXREGISTER.RVTYP}) = TRUE"
        ReportName = App.PATH & "\Reports\Itemwise BoxStockSummary.rpt"
        RPTN = "ITEM WISE BOX PENDING REPORT {Pending}"
    End If
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
           
    PERIOD = dtFrom & " To " & dtTo
    
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        .Formulas(3) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(4) = "PERIOD='" & PERIOD & "'"
        .Formulas(5) = "OPDT=#" & Format(dtopdt, "MM/dd/yyyy") & "#"
         RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
         If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
         ElseIf ReadConfigMaster("000068", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
         Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
         End If
        
        '.WindowShowPrintBtn = True
        '.WindowShowPrintSetupBtn = True
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
    Call ColorComponent(Me)
    txtUnit_KeyDown vbKeyReturn, 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtopdt = Now
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
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
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "'  AND CODE<>'000001' AND RECSTAT='A'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    End If
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
        txtITEM = SearchList1("Select TOP 20 Code,Name From FINITMMST ", 0, Empty, "Select Item")
        txtITEM.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtITEM = Empty
        txtITEM.Tag = Empty
    End If
End Sub

Private Sub txtItem_LostFocus()
 txtITEM.BackColor = vbWhite
End Sub


Private Sub TXTICOD_GotFocus()
TXTICOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTICOD_LostFocus()
TXTICOD.BackColor = vbWhite
End Sub


Private Sub txtPCOD_GotFocus()
TXTPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTPCOD = SearchList1("Select  TOP 20 CODE,NAME From ACCMST ", 0, Empty, "Select PARTY FROM LIST")
        TXTPCOD.Tag = Key
    End If
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
    End If
    If KeyCode = vbKeyDelete Then
       TXTPCOD = Empty
    End If

End Sub

Private Sub txtPCOD_LostFocus()
TXTPCOD.BackColor = vbWhite
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


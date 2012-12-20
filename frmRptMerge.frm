VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRptMerge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge No. Wise Ledger"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6765
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtUNIT 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
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
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   6615
      Begin VB.TextBox TXTMRGN 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   4815
      End
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
         TabIndex        =   4
         Top             =   720
         Width           =   4815
      End
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1560
         TabIndex        =   2
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
         TabIndex        =   3
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
      Begin VB.Label Label2 
         Caption         =   "Merge No."
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
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
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
         TabIndex        =   12
         Top             =   720
         Width           =   1335
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
         TabIndex        =   11
         Top             =   240
         Width           =   735
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
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   6615
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   5
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
         TabIndex        =   6
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
         Image           =   "frmRptMerge.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   7
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
         Image           =   "frmRptMerge.frx":0452
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
         TabIndex        =   8
         Top             =   240
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmRptMerge"
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
Me.KeyPreview = False

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
     
            
    ReportName = Empty
    RPTN = Empty

   
        If cboReports.ListIndex = 0 Then
           ReportName = App.PATH & "\Reports\MergeNoLedger.rpt"
           RPTN = "MERGE NO. WISE LEDGER"
        End If
        
    Call VwMrgStock
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
    
    CRPT.SubreportToChange = ""
    CRPT.SubreportToChange = "RMREP.rpt"
    CRPT.Connect = "DSN=" & ServerName & ";UID=sa;PWD= " & DefaultPassword_live & ";DSQ=" & CN.DefaultDatabase
            
    CRPT.SubreportToChange = ""
    CRPT.SubreportToChange = "FINSBREP.rpt"
    CRPT.Connect = "DSN=" & ServerName & ";UID=sa;PWD= " & DefaultPassword_live & ";DSQ=" & CN.DefaultDatabase
    
    CRPT.Reset
    crptConnect CRPT
    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        .Formulas(3) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(4) = "PERIOD='" & PERIOD & "'"
               
         RPTN = RPTN + Space(5) + ReportName
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        If cUName = "ADMIN" Then
           CRPT.WindowShowPrintBtn = True
           CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000044", 8, "R") Then
           CRPT.WindowShowPrintBtn = True
           CRPT.WindowShowPrintSetupBtn = True
        Else
           CRPT.WindowShowPrintBtn = False
           CRPT.WindowShowPrintSetupBtn = False
        End If
                   
        If cUName = "ADMIN" Then
           CRPT.WindowShowPrintBtn = True
           CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000045", 8, "R") Then
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
 
 With cboReports
       .AddItem "Merge No. Wise Ledger"
       
       cboReports.ListIndex = 0
       Me.Caption = "Merge No. Wise Ledger "
    
 End With
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
    
    Me.Caption = ""
End Sub

Private Sub TXTMRGN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTMRGN = Empty
    ElseIf KeyCode = vbKeyF2 Then
        M_DESC = Empty
        Key = Empty
        NEW_VISIBLE = False
        TXTMRGN = SearchList1("Select DISTINCT MRGN,MRGN  From MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' ", 0, Empty, "Select MERGE FROM MASTER")
        TXTMRGN.Tag = Key
        'MERGE = Key
    End If
    If KeyCode = vbKeyDelete Then
       TXTMRGN = Empty
    End If
    
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

Private Sub SetSQL()
 rptsql = Empty
 
 
 rptsql = "{VW_MRGSTOCK.COMP}='" & compPth & "' AND {VW_MRGSTOCK.UNIT} = '" & txtUNIT.Tag & "' "
 
 'If txtIGRP <> Empty And UCase(txtITEM) <> "RCOPS" Then rptsql = rptsql & " AND {IGMMST.NAME}='" & txtIGRP & "'"
 If TXTMRGN <> Empty Then rptsql = rptsql & " AND {VW_MRGSTOCK.MRGN}='" & Trim(TXTMRGN) & "'"
    
End Sub

Private Sub VwMrgStock()
Dim QRY As String
On Error GoTo VWERR
   
QRY = "CREATE VIEW VW_MRGSTOCK AS SELECT COMP,UNIT,DVCD,PCOD,ICOD,STORETRAN.ICOD AS RICD,STORETRAN.MRGN AS MRGN,STORETRAN.DATE ,OPER,ISNULL(SUM(PCES),0) AS PCS,ISNULL(SUM(QNTY),0) AS QNTY," & _
" STORETRAN.CHLN,0 AS GRAD,STORETRAN.LTNO AS LOTNO  FROM STORETRAN WHERE STORETRAN.OPER = '+' AND STORETRAN.DVCD = '000001' AND STORETRAN.RECSTAT<>'D'  AND DATE >='" & Format(dtFrom, "MM/DD/YYYY") & "' AND DATE <= '" & Format(dtTo, "MM/DD/YYYY") & "'" & _
" GROUP BY STORETRAN.COMP,STORETRAN.UNIT,STORETRAN.DVCD,STORETRAN.PCOD,STORETRAN.ICOD,STORETRAN.MRGN,STORETRAN.OPER,STORETRAN.CHLN,STORETRAN.LTNO,STORETRAN.GRAD,STORETRAN.DATE" & _
" Union " & _
" SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,PCOD,ICOD,TXULOT.RICD,TXULOT.MRGN AS MRGN,SPTRAN.DATE,OPER,ISNULL(SUM(PCES),0) AS PCS,ISNULL(SUM(QNTY),0) AS QNTY, " & _
" SPTRAN.CHLN,SPTRAN.GRAD,SPTRAN.LTNO AS LOTNO FROM SPTRAN INNER JOIN " & _
" TXULOT ON TXULOT.COMP = SPTRAN.COMP AND TXULOT.UNIT = SPTRAN.UNIT AND TXULOT.DVCD = SPTRAN.DVCD AND TXULOT.LTNO = SPTRAN.LTNO " & _
" WHERE  SPTRAN.VTYP = 'DPF' AND  SPTRAN.RECSTAT<>'D'  AND DATE >='" & Format(dtFrom, "MM/DD/YYYY") & "' AND DATE <='" & Format(dtTo, "MM/DD/YYYY") & "'" & _
" GROUP BY SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,SPTRAN.PCOD,SPTRAN.ICOD,TXULOT.RICD,TXULOT.MRGN,SPTRAN.OPER,SPTRAN.CHLN,SPTRAN.LTNO,SPTRAN.GRAD,SPTRAN.DATE "
        
CN.Execute "IF ( OBJECT_ID('VW_MRGSTOCK') IS NOT NULL ) DROP VIEW VW_MRGSTOCK "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub





VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_RG1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   EXCISE : RG1 Register"
   ClientHeight    =   3825
   ClientLeft      =   2940
   ClientTop       =   870
   ClientWidth     =   6660
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   6495
      Begin VB.ComboBox cboFormats 
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
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   180
         Width           =   4785
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Rpt &Format"
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
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   6495
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3720
         TabIndex        =   14
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
         Image           =   "frmRPT_RG1.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   15
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
         Image           =   "frmRPT_RG1.frx":0452
         cBack           =   -2147483633
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2880
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         TabIndex        =   12
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1545
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   6495
      Begin VB.ComboBox txtExcChapter 
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
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtDVCD 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   4845
      End
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
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
         Left            =   4080
         TabIndex        =   9
         Top             =   1080
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
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&From Date"
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
         TabIndex        =   6
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
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
         Left            =   3240
         TabIndex        =   8
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Excise Chapter           "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LBLDIV 
         BackStyle       =   0  'Transparent
         Caption         =   "&Division Name            "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Height          =   585
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtUNIT 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4845
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRPT_RG1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RPTN As String
Dim PERIOD As String

Private Sub Form_Activate()
  Call ColorComponent(Me)
  If cboFormats.ListCount > 0 Then cboFormats.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtDVCD" And txtDVCD = Empty Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call CenterChild(frm_Main, Me)
    dtFrom = FSDT
    dtTo = FEDT
    Call SetReportFormat
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
    
    If cboFormats.ListIndex = -1 Then
        MsgBox "Invalid Report format !! Choose From List !!", vbInformation
        cboFormats.SetFocus
        SendKeys "%{DOWN}"
        Exit Sub
    End If
    
    If txtUNIT = Empty Then
        MsgBox "Please Select Unit !!", vbInformation, "Unit Is Key Field Missing"
        txtUNIT.SetFocus
        Exit Sub
    End If
    
    If txtExcChapter = Empty Then
        MsgBox "Please Select Chapter !!", vbInformation, "Chapter Is Key Field Missing"
        txtExcChapter.SetFocus
        Exit Sub
    End If
           
    If cboFormats.ListIndex = -1 Then cboFormats.ListIndex = 0
    
    If cboFormats.ListIndex = 1 And txtDVCD = Empty Then
        MsgBox "Please Select Division !!", vbInformation, "Key Field Division is Missing"
        txtDVCD.SetFocus
        Exit Sub
    End If
     
     CRPT.Reset
     crptConnect CRPT
    
     rptsql = Empty
    
     If cboFormats.ListIndex = 0 Then
        rptsql = "{VW_RG1.COMP}='" & compPth & "' AND {VW_RG1.UNIT} = '" & txtUNIT.Tag & _
        "' AND {VW_RG1.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & _
        ") AND {VW_RG1.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
     Else
        rptsql = "{VW_RG1.COMP}='" & compPth & "' AND {VW_RG1.UNIT} = '" & txtUNIT.Tag & _
        "' AND {VW_RG1.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
     End If
     
     rptsql = rptsql & " AND {VW_RG1.CHAPTERNO}='" & txtExcChapter & "'"
     ReportName = Empty
    
     Call GenViewForRG1
    
     Select Case cboFormats.ListIndex
       Case 0
           ReportName = App.PATH & "\Reports\ExciseRG1_Summary.rpt"
           RPTN = "RG-1 REGISTER SUMMARY"
       Case 1
           ReportName = App.PATH & "\Reports\ExciseRG1_Detail.rpt"
           RPTN = "RG-1 REGISTER"
     End Select
        
     RPTN = RPTN & " (Excise Chapter No. : " & txtExcChapter & " )"
        
     If ReportName = Empty Then
        ReportErrorMessage 0
        Exit Sub
     End If
    
     CRPT.ReportFileName = ReportName
    
     If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
     End If
          
     PERIOD = dtFrom & " To " & dtTo
    
     CRPT.ReplaceSelectionFormula rptsql
    
     With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        .Formulas(3) = "DIVISION='" & txtDVCD & "'"
        .Formulas(4) = "PERIOD='" & PERIOD & "'"
        .Formulas(5) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(6) = "STDT=#" & Format(dtFrom, "MM/dd/yyyy") & "#"
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000079", 8, "R") Then
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub SetReportFormat()
    With cboFormats
         .Clear
         .AddItem "RG-1 Register Summary"
         .AddItem "RG-1 Register Detail"
    End With
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' and UNIT='" & txtUNIT.Tag & "' AND CODE<>'000001' AND RECSTAT<>'D'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
End Sub

Private Sub txtExcChapter_GotFocus()
  
  If txtExcChapter = Empty Then
     Call FillCombo
  End If
  
  txtExcChapter.Height = 1155
  txtExcChapter.ZOrder
  ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
  
End Sub

Private Sub txtExcChapter_LostFocus()
   txtExcChapter.BackColor = vbWhite
   txtExcChapter.Height = 325
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If
End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
 txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTZOOM_LostFocus()
 txtZoom.BackColor = vbWhite
End Sub

Private Sub TXTZOOM_GotFocus()
 txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboFormats_GotFocus()
    cboFormats.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "%{DOWN}"
End Sub

Private Sub cboFormats_LostFocus()
 cboFormats.BackColor = vbWhite
End Sub

Private Sub GenViewForConsigner()
Dim QRY As String
On Error GoTo VWERR

QRY = "CREATE VIEW VW_PARTY_ITEMWISE_RATE_LIFTING AS " & _
   "SELECT SPTRAN.COMP,SPTRAN.UNIT,ACCMST.NAME AS PARTY,FINITMMST.NAME AS DENIER, " & _
   "ISNULL(SUM(SPTRAN.QNTY * SPTRAN.RATE),0) / ISNULL(SUM(SPTRAN.QNTY),1) AS RATE  FROM SPTRAN " & _
   "INNER JOIN ACCMST ON ACCMST.CODE = SPTRAN.PCOD " & _
   "INNER JOIN FINITMMST ON FINITMMST.COMP = SPTRAN.COMP AND FINITMMST.UNIT = SPTRAN.UNIT " & _
   "AND FINITMMST.DVCD = SPTRAN.DVCD AND FINITMMST.CODE = SPTRAN.ICOD " & _
   "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & txtUNIT.Tag & _
   "' AND SPTRAN.VTYP='DPF' AND SPTRAN.DVCD='" & txtDVCD.Tag & _
   "' AND SPTRAN.DATE >='" & Format(dtFrom.Text, "MM/DD/YYYY") & _
   "' AND SPTRAN.DATE<='" & Format(dtTo.Text, "MM/DD/YYYY") & "' AND SPTRAN.RECSTAT<>'D' "

If TXTPCOD <> Empty Then QRY = QRY & " AND SPTRAN.PCOD='" & txtParty.Tag & "' "
If TXTITEM <> Empty Then QRY = QRY & " AND SPTRAN.ICOD='" & TXTITEM.Tag & "' "

QRY = QRY & " GROUP BY SPTRAN.COMP,SPTRAN.UNIT,ACCMST.NAME,FINITMMST.NAME"
       
CN.Execute "IF ( OBJECT_ID('VW_PARTY_ITEMWISE_RATE_LIFTING') IS NOT NULL ) DROP VIEW VW_PARTY_ITEMWISE_RATE_LIFTING "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

Private Sub GenViewForRG1()
Dim QRY As String
Dim ADDIV As String

On Error GoTo VWERR

If txtDVCD <> Empty Then
   ADDIV = " AND DVCD='" & txtDVCD.Tag & "' "
Else
   ADDIV = ""
End If

'1.PACKING WITHOUT GR AND WASTAGE : PPF
QRY = "CREATE VIEW VW_RG1 AS " & _
      "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.VBDT AS DATE,DIVMST.CHAPTERNO,'PPF' AS VTYP,ISNULL(SUM(NTWGT),0) AS QNTY,'' AS RMRK FROM BOXREGISTER " & _
      "INNER JOIN DIVMST ON DIVMST.COMP = BOXREGISTER.COMP AND DIVMST.UNIT = BOXREGISTER.UNIT AND DIVMST.CODE = BOXREGISTER.DVCD  " & _
      "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & txtUNIT.Tag & _
      "' AND BOXREGISTER.RECSTAT<>'D' " & ADDIV & " AND DBCD<>'000004' AND DBCD<>'000006' " & _
      " AND VBDT <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' GROUP BY BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.VBDT,DIVMST.CHAPTERNO "


'2.PACKING ONLY GR : GRP
QRY = QRY & "UNION " & _
      "SELECT GRPACKING.COMP,GRPACKING.UNIT,GRPACKING.DVCD,GRPACKING.VBDT AS DATE,DIVMST.CHAPTERNO,'GRP' AS VTYP,ISNULL(SUM(NETWGT),0) AS QNTY,'' AS RMRK FROM GRPACKING " & _
      "INNER JOIN DIVMST ON DIVMST.COMP = GRPACKING.COMP AND DIVMST.UNIT = GRPACKING.UNIT AND DIVMST.CODE = GRPACKING.DVCD  " & _
      "WHERE GRPACKING.COMP='" & compPth & "' AND GRPACKING.UNIT='" & txtUNIT.Tag & _
      "' AND GRPACKING.RECSTAT<>'D' " & ADDIV & "  " & _
      " AND VBDT <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' GROUP BY GRPACKING.COMP,GRPACKING.UNIT,GRPACKING.DVCD,GRPACKING.VBDT,DIVMST.CHAPTERNO "
      
'3.PACKING ONLY WASTAGE
'DBCD='000006' "
'(LOTNO='' OR LOTNO='WASTE' OR LOTNO IS NULL)
QRY = QRY & "UNION " & _
      "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.VBDT AS DATE,UNTCFG.WCHAP AS CHAPTERNO,'PPF' AS VTYP,ISNULL(SUM(NTWGT),0) AS QNTY,'' AS RMRK FROM BOXREGISTER " & _
      "INNER JOIN UNTCFG ON UNTCFG.COMP = BOXREGISTER.COMP AND UNTCFG.UNIT = BOXREGISTER.UNIT " & _
      "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & txtUNIT.Tag & _
      "' AND BOXREGISTER.RECSTAT<>'D' " & ADDIV & " AND DBCD='000006' " & _
      " AND VBDT <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' GROUP BY BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.VBDT,UNTCFG.WCHAP "
      
'4.NEW GR CLEARANCE
QRY = QRY & "UNION " & _
      "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.VBDT,DIVMST.CHAPTERNO,'GRD' AS VTYP,ISNULL(SUM(NTWGT),0) AS QNTY,'' AS RMRK FROM BOXREGISTER " & _
      "INNER JOIN DIVMST ON DIVMST.COMP = BOXREGISTER.COMP AND DIVMST.UNIT = BOXREGISTER.UNIT AND DIVMST.CODE = BOXREGISTER.DVCD  " & _
      "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & txtUNIT.Tag & "' AND BOXREGISTER.RECSTAT<>'D' " & ADDIV & _
      " AND PCOD='GRPACK' AND VBDT <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' " & _
      "GROUP BY BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,VBDT,DIVMST.CHAPTERNO "
            
'4.DISPATCH WITHOUT WASTAGE,EXPORT AND CAPTIVE
QRY = QRY & "UNION " & _
      "SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,SPTRAN.DATE,DIVMST.CHAPTERNO,'DPF' AS VTYP,ISNULL(SUM(QNTY),0) AS QNTY,'' AS RMRK FROM SPTRAN " & _
      "INNER JOIN DIVMST ON DIVMST.COMP = SPTRAN.COMP AND DIVMST.UNIT = SPTRAN.UNIT AND DIVMST.CODE = SPTRAN.DVCD  " & _
      "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & txtUNIT.Tag & "' AND SPTRAN.RECSTAT<>'D' " & ADDIV & _
      " AND VTYP='DPF' AND DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' AND DBCD NOT IN ('000002','000004','000005') " & _
      " AND SPTRAN.LTNO<>'' AND SPTRAN.LTNO<>'WASTE' AND SPTRAN.LTNO IS NOT NULL " & _
      "GROUP BY SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,DATE,DIVMST.CHAPTERNO "
      
'5.DISPATCH ONLY EXPORT AND CAPTIVE
QRY = QRY & "UNION " & _
      "SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,SPTRAN.DATE,DIVMST.CHAPTERNO,'DDF' AS VTYP,ISNULL(SUM(QNTY),0) AS QNTY,'' AS RMRK FROM SPTRAN " & _
      "INNER JOIN DIVMST ON DIVMST.COMP = SPTRAN.COMP AND DIVMST.UNIT = SPTRAN.UNIT AND DIVMST.CODE = SPTRAN.DVCD  " & _
      "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & txtUNIT.Tag & "' AND SPTRAN.RECSTAT<>'D' " & ADDIV & _
      " AND SPTRAN.VTYP='DPF' AND SPTRAN.DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & _
      "' AND SPTRAN.DBCD IN ('000002','000004') " & _
      "GROUP BY SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,DATE,DIVMST.CHAPTERNO "
      
'6.DISPATCH ONLY WASTAGE
QRY = QRY & "UNION " & _
      "SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,SPTRAN.DATE,UNTCFG.WCHAP AS CHAPTERNO,'DPF' AS VTYP,ISNULL(SUM(QNTY),0) AS QNTY,'' AS RMRK FROM SPTRAN " & _
      "INNER JOIN UNTCFG ON UNTCFG.COMP = SPTRAN.COMP AND UNTCFG.UNIT = SPTRAN.UNIT " & _
      "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & txtUNIT.Tag & "' AND SPTRAN.RECSTAT<>'D' " & ADDIV & _
      " AND SPTRAN.VTYP='DPF' AND SPTRAN.DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & _
      "' AND (SPTRAN.LTNO='' OR SPTRAN.LTNO='WASTE' OR SPTRAN.LTNO IS NULL) " & _
      "GROUP BY SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,DATE,UNTCFG.WCHAP "
      
'7.EXCISABLE SALE
QRY = QRY & "UNION " & _
      "SELECT BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP AS CHAPTERNO,'SAL' AS VTYP,ISNULL(SUM(BILLMAIN.ITOT),0) AS QNTY,'' AS RMRK FROM BILLMAIN " & _
      "INNER JOIN EGPMAN ON EGPMAN.COMP = BILLMAIN.COMP AND EGPMAN.UNIT = BILLMAIN.UNIT AND EGPMAN.VTYP = BILLMAIN.VTYP AND EGPMAN.DBCD = BILLMAIN.DBCD AND EGPMAN.VBNO = BILLMAIN.VBNO " & _
      "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & txtUNIT.Tag & "' AND BILLMAIN.RECSTAT<>'D' " & ADDIV & _
      " AND BILLMAIN.VTYP='SAL' AND BILLMAIN.DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & _
      "' AND BILLMAIN.ITOT <> 0 AND (BILLMAIN.CENVAT+BILLMAIN.EDUCESS+BILLMAIN.H_ED_CESS)<>0 GROUP BY BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP "
'8.NON-EXCISABLE SALE
QRY = QRY & "UNION " & _
      "SELECT BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP AS CHAPTERNO,'SSL' AS VTYP,ISNULL(SUM(BILLMAIN.ITOT),0) AS QNTY,'' AS RMRK FROM BILLMAIN " & _
      "INNER JOIN EGPMAN ON EGPMAN.COMP = BILLMAIN.COMP AND EGPMAN.UNIT = BILLMAIN.UNIT AND EGPMAN.VTYP = BILLMAIN.VTYP AND EGPMAN.DBCD = BILLMAIN.DBCD AND EGPMAN.VBNO = BILLMAIN.VBNO " & _
      "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & txtUNIT.Tag & "' AND BILLMAIN.RECSTAT<>'D' " & ADDIV & _
      " AND BILLMAIN.VTYP='SAL' AND BILLMAIN.DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & _
      "' AND BILLMAIN.ITOT <> 0 AND (BILLMAIN.CENVAT+BILLMAIN.EDUCESS+BILLMAIN.H_ED_CESS)=0 GROUP BY BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP "
'9. FRO CENVAT
QRY = QRY & "UNION " & _
      "SELECT BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP AS CHAPTERNO,'CEN' AS VTYP,ISNULL(SUM(BILLMAIN.CENVAT),0) AS QNTY,'' AS RMRK FROM BILLMAIN " & _
      "INNER JOIN EGPMAN ON EGPMAN.COMP = BILLMAIN.COMP AND EGPMAN.UNIT = BILLMAIN.UNIT AND EGPMAN.VTYP = BILLMAIN.VTYP AND EGPMAN.DBCD = BILLMAIN.DBCD AND EGPMAN.VBNO = BILLMAIN.VBNO " & _
      "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & txtUNIT.Tag & "' AND BILLMAIN.RECSTAT<>'D' " & ADDIV & _
      " AND BILLMAIN.VTYP='SAL' AND BILLMAIN.DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & _
      "' AND BILLMAIN.ITOT <> 0 AND BILLMAIN.CENVAT<>0 GROUP BY BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP "
'10. FOR EDU_CESS
QRY = QRY & "UNION " & _
      "SELECT BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP AS CHAPTERNO,'EDC' AS VTYP,ISNULL(SUM(BILLMAIN.EDUCESS),0) AS QNTY,'' AS RMRK FROM BILLMAIN " & _
      "INNER JOIN EGPMAN ON EGPMAN.COMP = BILLMAIN.COMP AND EGPMAN.UNIT = BILLMAIN.UNIT AND EGPMAN.VTYP = BILLMAIN.VTYP AND EGPMAN.DBCD = BILLMAIN.DBCD AND EGPMAN.VBNO = BILLMAIN.VBNO " & _
      "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & txtUNIT.Tag & "' AND BILLMAIN.RECSTAT<>'D' " & ADDIV & _
      " AND BILLMAIN.VTYP='SAL' AND BILLMAIN.DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & _
      "' AND BILLMAIN.ITOT <> 0 AND BILLMAIN.EDUCESS<>0 GROUP BY BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP "
'11. FOR H_ED_CESS
QRY = QRY & "UNION " & _
      "SELECT BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP AS CHAPTERNO,'HEC' AS VTYP,ISNULL(SUM(BILLMAIN.H_ED_CESS),0) AS QNTY,'' AS RMRK FROM BILLMAIN " & _
      "INNER JOIN EGPMAN ON EGPMAN.COMP = BILLMAIN.COMP AND EGPMAN.UNIT = BILLMAIN.UNIT AND EGPMAN.VTYP = BILLMAIN.VTYP AND EGPMAN.DBCD = BILLMAIN.DBCD AND EGPMAN.VBNO = BILLMAIN.VBNO " & _
      "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & txtUNIT.Tag & "' AND BILLMAIN.RECSTAT<>'D' " & ADDIV & _
      " AND BILLMAIN.VTYP='SAL' AND BILLMAIN.DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & _
      "' AND BILLMAIN.ITOT <> 0 AND BILLMAIN.H_ED_CESS<>0 GROUP BY BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP "
'12 FOR REMARKS : INVOICE NO. FROM FIRST - LAST (011011-061011)
QRY = QRY & "UNION " & _
      "SELECT BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP AS CHAPTERNO,'RMK' AS VTYP,0 AS QNTY," & _
      "LEFT(MIN(BILLMAIN.VBNO),6) + '-' + LEFT(MAX(BILLMAIN.VBNO),6) AS RMRK FROM BILLMAIN " & _
      "INNER JOIN EGPMAN ON EGPMAN.COMP = BILLMAIN.COMP AND EGPMAN.UNIT = BILLMAIN.UNIT AND EGPMAN.VTYP = BILLMAIN.VTYP AND EGPMAN.DBCD = BILLMAIN.DBCD AND EGPMAN.VBNO = BILLMAIN.VBNO " & _
      "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & txtUNIT.Tag & "' AND BILLMAIN.RECSTAT<>'D' " & ADDIV & _
      " AND BILLMAIN.VTYP='SAL' AND BILLMAIN.DATE >='" & Format(dtFrom.Text, "MM/DD/YYYY") & _
      "' AND BILLMAIN.ITOT <> 0 AND (BILLMAIN.CENVAT+BILLMAIN.EDUCESS+BILLMAIN.H_ED_CESS) <> 0 GROUP BY BILLMAIN.COMP,BILLMAIN.UNIT,BILLMAIN.DVCD,BILLMAIN.DATE,EGPMAN.CHAP "
      
CN.Execute "IF ( OBJECT_ID('VW_RG1') IS NOT NULL ) DROP VIEW VW_RG1 "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

Public Sub FillCombo()

Dim SQL As String
Dim rsGeneral As ADODB.Recordset
Set rsGeneral = New Recordset
txtExcChapter.Clear

    SQL = "SELECT DISTINCT COMP,UNIT,CODE,CHAPTERNO FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
          "' AND CODE='" & txtDVCD.Tag & "' "
    
    If rsGeneral.State = 1 Then rsGeneral.Close
    rsGeneral.Open SQL, CN
    
    If rsGeneral.EOF = False Then
       txtExcChapter = Trim(rsGeneral(3))
    End If
    
    Do While rsGeneral.EOF = False
        If rsGeneral(3) & "" <> "" Then txtExcChapter.AddItem Trim(rsGeneral(3))
        rsGeneral.MoveNext
    Loop
    rsGeneral.Close
    
    SQL = "SELECT DISTINCT COMP,UNIT,WCHAP AS CHAPTERNO FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "'"
    If rsGeneral.State = 1 Then rsGeneral.Close
    rsGeneral.Open SQL, CN
    Do While rsGeneral.EOF = False
        If rsGeneral(2) & "" <> "" Then txtExcChapter.AddItem Trim(rsGeneral(2))
        rsGeneral.MoveNext
    Loop
    rsGeneral.Close
    
End Sub


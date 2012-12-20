VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRptSpecification 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Specification Wise Report"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6795
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   19
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
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
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
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
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
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6615
      Begin VB.TextBox MERGE 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   4815
      End
      Begin VB.OptionButton OptLedger 
         Caption         =   "Ledger"
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
         Left            =   5400
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptDetail 
         Caption         =   "Detail"
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
         Left            =   3480
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptStock 
         Caption         =   "Summary"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtIGRP 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtITEM 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "&Merge No.      "
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
         TabIndex        =   24
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "&Item Group  "
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
         Width           =   1695
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
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   6615
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   9
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
         TabIndex        =   8
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
         Image           =   "frmRptSpecification.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   10
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
         Image           =   "frmRptSpecification.frx":0452
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
         TabIndex        =   11
         Top             =   240
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmRptSpecification"
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
Dim SPECI As String
Dim MRGN As String

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

 '   If cboReports.ListIndex = -1 Then
 '       MsgBox "Please Select Report Format ", vbInformation
 '       cboReports.SetFocus
 '       SendKeys "{DOWN}"
 '       Exit Sub
 '   End If
     
     If txtIGRP = Empty Then
        MsgBox "Please Select Item Group", vbOKOnly
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
    
    Call FindSpecification
               
    CRPT.Reset
    crptConnect CRPT
     
            
    ReportName = Empty
    RPTN = Empty
     
    If (SPECI = 0 Or SPECI = 1) And MRGN = "Y" And OptStock.Value = True Then
      ReportName = App.PATH & "\Reports\RPT_STORESTOCKSTATUS_PCS.rpt"
      RPTN = "PCS + MERGE WISE STOCK STATUS SUMMARY"
    ElseIf (SPECI = 0 Or SPECI = 1) And MRGN = "Y" And OptDetail.Value = True Then
      ReportName = App.PATH & "\Reports\RPT_MRGSTORESTOCKDETAIL_PCS.rpt"
      RPTN = "PCS + MERGE WISE STOCK DETAIL"
    ElseIf (SPECI = 0 Or SPECI = 1) And MRGN = "Y" And OptLedger.Value = True Then
      ReportName = App.PATH & "\Reports\MergeWiseItemLedger_PCS.rpt"
      RPTN = "PCS + MERGE WISE  LEDGER"
    ElseIf SPECI = 3 And MRGN = "Y" And OptStock.Value = True Then
      ReportName = App.PATH & "\Reports\RPT_STORESTOCKSTATUS.rpt"
      RPTN = "COPS + MERGE WISE STOCK STATUS SUMMARY"
    ElseIf SPECI = 3 And MRGN = "Y" And OptDetail.Value = True Then
      ReportName = App.PATH & "\Reports\RPT_MRGSTORESTOCKDETAIL.rpt"
      RPTN = "COPS + MERGE WISE STOCK DETAIL"
    ElseIf SPECI = 3 And MRGN = "Y" And OptLedger.Value = True Then
      ReportName = App.PATH & "\Reports\MergeWiseItemLedger.rpt"
      RPTN = "COPS + MERGE WISE  LEDGER"
      
    ' WITHOUT MERGE NO. WISE
    
    ElseIf SPECI = 0 And MRGN <> "Y" And OptStock.Value = True Then
      ReportName = App.PATH & "\Reports\RPT_STORESTOCKSTATUS_PCSWOMRGN.rpt"
      RPTN = "PCS + QUANTITY WISE STOCK SUMMARY"
    ElseIf SPECI = 0 And MRGN <> "Y" And OptDetail.Value = True Then
      ReportName = App.PATH & "\Reports\RPT_STORESTOCKDETAIL_PCSWOMRGN.rpt"
      RPTN = "PCS + QUANTITY WISE STOCK DETAIL"
    ElseIf SPECI = 0 And MRGN <> "Y" And OptLedger.Value = True Then
      ReportName = App.PATH & "\Reports\MergeWiseItemLedger_PCSWOMRGN.rpt"
      RPTN = "PCS + QUANTITY WISE  LEDGER"
    ElseIf SPECI = 3 And MRGN <> "Y" And OptStock.Value = True Then
       ReportName = App.PATH & "\Reports\RPT_STORESTOCKSTATUSWOMRGN.rpt"
      RPTN = "COPS + QUANTITY WISE STOCK SUMMARY"
    ElseIf SPECI = 3 And MRGN <> "Y" And OptDetail.Value = True Then
       ReportName = App.PATH & "\Reports\RPT_MRGSTORESTOCKDETAILWOMRGN.rpt"
       RPTN = "COPS + QUANTITY WISE STOCK DETAIL"
    ElseIf SPECI = 3 And MRGN <> "Y" And OptLedger.Value = True Then
      ReportName = App.PATH & "\Reports\MergeWiseItemLedgerCOPSWOMRGN.rpt"
      RPTN = "COPS + QUANTITY WISE  LEDGER"
    End If
           
    
    
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
        If Me.Tag <> "ISS" Then
          .Formulas(5) = "OPDT=#" & Format(DateAdd("D", -1, dtFrom), "MM/dd/yyyy") & "#"
        End If
        
         RPTN = RPTN + Space(5) + ReportName
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        If cUName = "ADMIN" Then
           CRPT.WindowShowPrintBtn = True
           CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000064", 8, "R") Then
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
     Me.KeyPreview = True
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
       
       .AddItem "Cops + Merge Wise Stock Summary"
       .AddItem "Pcs/Box + Merge Wise Stck Summary"
       .AddItem "Cops + Merge Wise Stock Detail"
       .AddItem "Pcs/Box + Merge Wise Stock Detail"
       .AddItem "Store Item Ledger(Cops + MergeNo.Wise)"
       .AddItem "Store Item Ledger(Pcs/Box + MergeNo. Wise)"
       .AddItem "Cops + Quantity Wise Stock Summary"
       .AddItem "Pcs/Box + Quantity Wise Stock Summary"
       .AddItem "Cops + Quantity Wise Stock Detail"
       .AddItem "Pcs/Box + Quantity Wise Stock Detail"
       .AddItem "Store Item Ledger (Cops + Quantity Wise)"
       .AddItem "Store Item Ledger (Pcs + Quantity Wise)"
       
        cboReports.ListIndex = 0
        Me.Caption = "Store Item Ledger / Summary "
    
 End With
 
 Me.KeyPreview = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
    Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtFrom.Text = Format(FSDT, "dd/MM/yyyy")
    dtTo.Text = Format(FEDT, "dd/MM/yyyy")
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    
    Me.Caption = ""
End Sub

Private Sub MERGE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
NEW_VISIBLE = False
M_DESC = Empty
Key = Empty
If txtITEM <> Empty Then
MERGE.Text = SearchList1(" SELECT DISTINCT MRGN,MRGN FROM MRGMST WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & "' AND ICOD = '" & GetCode("ITMMST", txtITEM, "NAME", "CODE") & "'", 0, MERGE.Text, "SELECT MERGE NO. FROM LIST ")
End If
End If

If KeyCode = vbKeyDelete Then
  MERGE = Empty
End If

End Sub

Private Sub txtIGRP_GotFocus()
 txtIGRP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTIGRP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtIGRP = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtIGRP.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM IGMMST WHERE SPECIFICATION <> 2 ", 0, "", "List Of Item Group")
        txtIGRP.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtIGRP = Empty
        txtIGRP.Tag = Empty
    End If
    
    
End Sub

Private Sub txtIGRP_LostFocus()
   txtIGRP.BackColor = vbWhite
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

Private Sub SetSQL()
 rptsql = Empty
 
 
 rptsql = "{STORETRAN.COMP}='" & compPth & "' AND {STORETRAN.UNIT} = '" & txtUNIT.Tag & "' AND {STORETRAN.DVCD}='000001' "
 rptsql = rptsql & " AND {STORETRAN.RECSTAT}<>'D' "
 
 If Me.Tag = "ISS" Then
   rptsql = rptsql & " AND {STORETRAN.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {STORETRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
 Else
   rptsql = rptsql & " AND {STORETRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")"
 End If
 
 'rptsql = rptsql & "AND {STORETRAN.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & _
 '") AND {STORETRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {STORETRAN.RECSTAT}<>'D' "
 
 If Me.Tag = "ISS" Then
    rptsql = rptsql & "AND {STORETRAN.OPER}='-'"
 End If
 
    If txtIGRP <> Empty And UCase(txtITEM) <> "RCOPS" Then rptsql = rptsql & " AND {IGMMST.NAME}='" & txtIGRP & "'"
    If txtITEM <> Empty Then rptsql = rptsql & " AND {ITMMST.NAME}='" & txtITEM & "'"
    If MERGE <> Empty Then rptsql = rptsql & " AND {STORETRAN.MRGN}='" & Trim(MERGE) & "'"
    
    
End Sub

Private Sub Vw_Merge_Stock()
On Error Resume Next
CN.Execute "DROP VIEW VW_MERGE_STOCK"
CN.Execute "CREATE VIEW VW_MERGE_STOCK AS" & _
" SELECT COMP,UNIT,DVCD,ICOD,ITEM,LTNO,ISNULL(SUM(INWCOPS -OUTCOPS),0) AS COPS,ISNULL(SUM(INWQTY - OUTQTY),0) AS QTY FROM" & _
" (SELECT STORETRAN.COMP,STORETRAN.UNIT,STORETRAN.DVCD,ICOD,ITMMST.NAME AS ITEM,LTNO," & _
" ISNULL(SUM(COPS),0) AS INWCOPS,ISNULL(SUM(QNTY),0) AS INWQTY,0 AS OUTCOPS,0 AS OUTQTY FROM STORETRAN" & _
" INNER JOIN ITMMST ON ITMMST.CODE = STORETRAN.ICOD " & _
" WHERE RECSTAT<>'D' AND OPER='+' AND  LTRIM(RTRIM(LTNO))<>'' AND LTNO IS NOT NULL  AND DATE <= '" & Format(dtTo, "MM/DD/YYYY") & "' " & _
" GROUP BY STORETRAN.COMP,STORETRAN.UNIT,STORETRAN.DVCD,ICOD,ITMMST.NAME,LTNO " & _
" Union" & _
" SELECT STORETRAN.COMP,STORETRAN.UNIT,STORETRAN.DVCD, " & _
" ICOD,ITMMST.NAME AS ITEM,LTNO,0 AS INWCOPS,0 AS INWQTY, " & _
" ISNULL(SUM(COPS),0) AS OUTCOPS,ISNULL(SUM(QNTY),0) AS OUTQTY FROM STORETRAN " & _
" INNER JOIN ITMMST ON ITMMST.CODE = STORETRAN.ICOD " & _
" WHERE RECSTAT<>'D' AND OPER='-'  AND LTRIM(RTRIM(LTNO))<>'' AND DATE <= '" & Format(dtTo, "MM/DD/YYYY") & "'  AND LTNO IS NOT NULL " & _
" GROUP BY STORETRAN.COMP,STORETRAN.UNIT,STORETRAN.DVCD, " & _
" ICOD,ITMMST.NAME,LTNO)A1 " & _
" GROUP BY A1.COMP,A1.UNIT,A1.DVCD,A1.ICOD,A1.ITEM,A1.LTNO " & _
" Having ISNULL(Sum(INWCOPS - OUTCOPS), 0) > 0 And ISNULL(Sum(INWQTY - OUTQTY), 0) > 0"
  
End Sub

Private Sub FindSpecification()
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "SELECT *  FROM IGMMST WHERE NAME = '" & Trim(txtIGRP) & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
SPECI = RS!SPECIFICATION
MRGN = RS!MERGE
End If
End Sub

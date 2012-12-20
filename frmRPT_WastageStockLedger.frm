VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_WastageStockLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wastage Stock Ledger Report"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6840
   Begin VB.CheckBox chkAllowZero 
      Caption         =   "Display Zero Balance records ?"
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
      TabIndex        =   8
      Top             =   3240
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   0
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
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   720
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   6615
      Begin VB.TextBox txtITEM 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox TXTDVCD 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtMACHINE 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label2 
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
         TabIndex        =   22
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "&Division"
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
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "&Machine"
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
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   3480
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
         TabIndex        =   10
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
         Image           =   "frmRPT_WastageStockLedger.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   11
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
         Image           =   "frmRPT_WastageStockLedger.frx":0452
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
         TabIndex        =   12
         Top             =   240
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmRPT_WastageStockLedger"
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
    
    If TXTDVCD = Empty Then
       MsgBox "Please Select Division", vbInformation
       TXTDVCD.SetFocus
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
    
    Call SetViewForWastage
    
    If cboReports.ListIndex = 0 Then
       ReportName = App.PATH & "\Reports\Wastage Stock Ledger Report.rpt"
       RPTN = "ITEM WISE WASTAGE STOCK LEDGER REPORT "
    Else
       ReportName = App.PATH & "\Reports\Division+Item+WastageStock.rpt"
       RPTN = "ITEM WISE WASTAGE STOCK STATUS REPORT "
    End If

    rptsql = Empty
    rptsql = "{VW_WASTELEDGER.COMP}='" & compPth & "' AND {VW_WASTELEDGER.UNIT} = '" & txtUNIT.Tag & "' AND " & _
    "{VW_WASTELEDGER.DVCD}<>'000001' AND {VW_WASTELEDGER.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
    
    If TXTDVCD <> Empty Then rptsql = rptsql & " AND {VW_WASTELEDGER.DVCD}='" & TXTDVCD.Tag & "'"
    If txtMACHINE <> Empty Then rptsql = rptsql & " AND {VW_WASTELEDGER.MCCD}='" & txtMACHINE.Tag & "'"
    If txtITEM <> Empty Then rptsql = rptsql & " AND {VW_WASTELEDGER.ICOD}='" & txtITEM.Tag & "'"
                              
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
        .Formulas(5) = "OPDT=#" & Format(DateAdd("D", -0, dtFrom), "MM/dd/yyyy") & "#"
        .Formulas(6) = "DIVISION='" & TXTDVCD & "'"
         RPTN = RPTN + Space(5) + ReportName
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        If ReadConfigMaster("000048", 8, "R") Then
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
    
     With cboReports
        .AddItem "Item Wise Wastage Stock Ledger"
        .AddItem "Item Wise Wastage Stock Status Report"
        cboReports.ListIndex = 0
     End With
End Sub

Private Sub txtDVCD_GotFocus()
 TXTDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
If txtUNIT = Empty Then txtUNIT.Enabled = True: txtUNIT.SetFocus: Exit Sub

    If KeyCode = vbKeyF2 Or TXTDVCD = Empty Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTDVCD.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & _
        "' AND UNIT='" & txtUNIT.Tag & "' AND RECSTAT='A'  AND CODE<>'000001'", 0, "", "List Of Division")
        
        TXTDVCD.Tag = Key
        
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    
        TXTDVCD = Empty
        TXTDVCD.Tag = Empty
        
    End If
    
End Sub

Private Sub txtDVCD_LostFocus()
   TXTDVCD.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
  txtITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtITEM = SearchList1("Select TOP 20 Code,Name From FINITMMST WHERE COMP='" & compPth & _
                "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & TXTDVCD.Tag & "'", 0, Empty, "Select Item")
        txtITEM.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtITEM = Empty
        txtITEM.Tag = Empty
    End If
End Sub

Private Sub txtItem_LostFocus()
 txtITEM.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
  txtMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        If TXTDVCD = Empty Then TXTDVCD.Enabled = True: TXTDVCD.SetFocus: Exit Sub
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtMACHINE = SearchList1("Select TOP 20 Code,Name From MACMST WHERE COMP='" & compPth & _
        "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & TXTDVCD.Tag & "'", 0, Empty, "Select Machine")
        txtMACHINE.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtMACHINE = Empty
        txtMACHINE.Tag = Empty
    End If
End Sub

Private Sub txtMACHINE_LostFocus()
 txtMACHINE.BackColor = vbWhite
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

Private Sub SetViewForWastage()
Dim QRY As String
On Error GoTo VWERR
  
    QRY = "CREATE VIEW VW_WASTEITM AS " & _
          "SELECT COMP,UNIT,DVCD,ICOD,ISNULL(SUM(INWGT - OUTWGT),0) AS INWGT FROM " & _
          "(SELECT COMP,UNIT,DVCD,ICOD,ISNULL(SUM(GRSWGT),0) AS INWGT,0 AS OUTWGT FROM BOXREGISTER " & _
          "WHERE DBCD='000006' AND RECSTAT<>'D' AND GRSWGT<>0 " & _
          " AND VBDT<='" & Format(dtTo, "MM/dd/yyyy") & "' " & _
          "GROUP BY COMP,UNIT,DVCD,ICOD " & _
          "Union " & _
          "SELECT COMP,UNIT,DVCD,ICOD,0 AS INWGT,ISNULL(SUM(QNTY),0) AS OUTWGT FROM SPTRAN " & _
          "WHERE RECSTAT<>'D' AND VTYP='DPF' AND DBCD='000005' " & _
          " AND DATE<='" & Format(dtTo, "MM/dd/yyyy") & "' " & _
          "GROUP BY COMP,UNIT,DVCD,ICOD)T2 " & _
          "GROUP BY T2.COMP,T2.UNIT,T2.DVCD,T2.ICOD "
          If chkAllowZero.Value = 0 Then
             QRY = QRY + " Having IsNull(Sum(INWGT - OUTWGT), 0) > 0"
          End If
          
    CN.Execute "IF ( OBJECT_ID('VW_WASTEITM') IS NOT NULL ) DROP VIEW VW_WASTEITM"
    CN.Execute QRY
   
   QRY = "CREATE VIEW VW_WASTELEDGER AS " & _
         "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,'PPF' AS VTYP,BOXREGISTER.ICOD,BOXREGISTER.VBDT AS DATE,BOXREGISTER.VBNO,BOXREGISTER.GRSWGT AS QNTY FROM BOXREGISTER " & _
         "INNER JOIN VW_WASTEITM ON VW_WASTEITM.COMP=BOXREGISTER.COMP AND VW_WASTEITM.UNIT=BOXREGISTER.UNIT " & _
         "AND VW_WASTEITM.DVCD=BOXREGISTER.DVCD AND VW_WASTEITM.ICOD=BOXREGISTER.ICOD " & _
         "WHERE DBCD='000006' AND RECSTAT<>'D' AND GRSWGT<>0 " & _
         "UNION " & _
         "SELECT SPTRAN.COMP,SPTRAN.UNIT,SPTRAN.DVCD,'DPF' AS VTYP,SPTRAN.ICOD,SPTRAN.DATE,SPTRAN.VBNO,SPTRAN.QNTY FROM SPTRAN " & _
         "INNER JOIN VW_WASTEITM ON VW_WASTEITM.COMP=SPTRAN.COMP AND VW_WASTEITM.UNIT=SPTRAN.UNIT " & _
         "AND VW_WASTEITM.DVCD=SPTRAN.DVCD AND VW_WASTEITM.ICOD=SPTRAN.ICOD " & _
         "WHERE SPTRAN.RECSTAT<>'D' AND SPTRAN.VTYP='DPF' AND (SPTRAN.LTNO='WASTE' OR SPTRAN.LTNO='') "
         
CN.Execute "IF ( OBJECT_ID('VW_WASTELEDGER') IS NOT NULL ) DROP VIEW VW_WASTELEDGER "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

Private Sub SetViewForMachineWastage()
Dim QRY As String
On Error GoTo VWERR
            
 
       
   
    QRY = "CREATE VIEW VW_WASTELEDGER AS " & _
          "SELECT COMP,UNIT,DVCD,'PPF' AS VTYP,MCCD,ICOD,VBDT AS DATE,VBNO,GRSWGT AS QNTY FROM BOXREGISTER " & _
          "WHERE DBCD='000006' AND RECSTAT<>'D' AND GRSWGT<>0 " & _
          "UNION " & _
          "SELECT COMP,UNIT,DVCD,'DPF' AS VTYP,PCOD AS MCCD,ICOD,DATE,VBNO,QNTY FROM SPTRAN " & _
          "WHERE RECSTAT<>'D' AND VTYP='DPF' AND DBCD='000005' "
 
         
         
CN.Execute "IF ( OBJECT_ID('VW_WASTELEDGER') IS NOT NULL ) DROP VIEW VW_WASTELEDGER "
CN.Execute QRY
Exit Sub
VWERR:
MsgBox ERR.Description
End Sub



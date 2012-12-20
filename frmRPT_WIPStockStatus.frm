VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_WIPStockStatus 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WIP Stock Status"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   6615
      Begin VB.TextBox txtIGRP 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtITEM 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label5 
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
         TabIndex        =   13
         Top             =   240
         Width           =   1695
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
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   4680
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
         Image           =   "frmRPT_WIPStockStatus.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   20
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
         Image           =   "frmRPT_WIPStockStatus.frx":0452
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
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   6615
      Begin VB.TextBox txtMACHINE 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox TXTDVCD 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   4815
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
         TabIndex        =   11
         Top             =   600
         Width           =   1215
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
Attribute VB_Name = "frmRPT_WIPStockStatus"
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
       Call SetViewForWastage
       ReportName = App.PATH & "\Reports\Division Wise WIP Stock Status Report.rpt"
       RPTN = "DIVISION WISE WIP STOCK STATUS REPORT "
       If txtMACHINE <> Empty Then RPTN = RPTN & " ( " & txtMACHINE & ")"
    ElseIf cboReports.ListIndex = 1 Then
       Call SetViewForMachineWastage
       ReportName = App.PATH & "\Reports\Division+Machine Wise WIP Stock Status Report.rpt"
       RPTN = "DIVISION + MACHINE WISE WIP STOCK STATUS REPORT"
    ElseIf cboReports.ListIndex = 2 Then
       ReportName = App.PATH & "\Reports\Division+Machine WIP Ledger.rpt"
       RPTN = "DIVISION + MACHINE WISE WIP LEDGER REPORT"
    ElseIf cboReports.ListIndex = 3 Then
       Call SetViewForWastage
       ReportName = App.PATH & "\Reports\Division+Item Wise WIP Stock Status Report.rpt"
       RPTN = "DIVISION + ITEM WISE WIP STOCK STATUS REPORT"
    ElseIf cboReports.ListIndex = 4 Then
       Call SetViewForMachineWastage
       ReportName = App.PATH & "\Reports\Division+Machine+Item Wise WIP Stock Status Report.rpt"
       RPTN = "DIVISION + MACHINE + ITEM WISE WIP STOCK STATUS REPORT"
    End If

    rptsql = Empty
    rptsql = "{STORETRAN.COMP}='" & compPth & "' AND {STORETRAN.UNIT} = '" & txtUNIT.Tag & "' AND " & _
    "{STORETRAN.DVCD}<>'000001' AND {STORETRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") " & _
    "AND {STORETRAN.RECSTAT}<>'D'  "
    
    If TXTDVCD <> Empty Then rptsql = rptsql & " AND {STORETRAN.DVCD}='" & TXTDVCD.Tag & "'"
    If txtMACHINE <> Empty Then rptsql = rptsql & " AND {STORETRAN.PCOD}='" & txtMACHINE.Tag & "'"
    
    If txtIGRP <> Empty And UCase(txtITEM) <> "RCOPS" Then
       RPTN = RPTN + "(" & txtIGRP & ")"
       rptsql = rptsql & " AND {STORETRAN.IGCD} IN [" & txtIGRP.Tag & "]"
    End If
        
    If txtITEM <> Empty Then rptsql = rptsql & " AND {STORETRAN.ITEM}='" & txtITEM & "'"

                              
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
        .AddItem "Division Wise WIP Stock"
        .AddItem "Division + M/c Wise WIP Stock"
        .AddItem "Division + M/c Wise WIP Stock Ledger"
        .AddItem "Division + Item Wise WIP Stock"
        .AddItem "Division + M/c + Item Wise WIP Stock"
        cboReports.ListIndex = 0
     End With
End Sub

Private Sub txtDVCD_GotFocus()
 TXTDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
If txtUNIT = Empty Then txtUNIT.Enabled = True: txtUNIT.SetFocus: Exit Sub

    If KeyCode = vbKeyF2 Then
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

Private Sub txtIGRP_GotFocus()
 txtIGRP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTIGRP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        MSelCata = "G"
        LOAD frm_MulSelectCode
        frm_MulSelectCode.Show 1
        
        If Len(MSelName) > 0 Then
            txtIGRP.Text = Trim(MSelName)
            txtIGRP.Tag = Trim(MSelCode)
        End If
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
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
   
   QRY = "CREATE VIEW VW_WASTAGE AS " & _
         "SELECT COMP,UNIT,DVCD,ISNULL(SUM(GRSWGT),0) AS WASTAGE_QNTY FROM BOXREGISTER " & _
         "WHERE DBCD='000006' AND RECSTAT<>'D' " & _
         " AND VBDT >='" & Format(dtFrom.Text, "MM/DD/YYYY") & _
         "' AND VBDT <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' GROUP BY COMP,UNIT,DVCD " & _
         "UNION " & _
         "SELECT COMP,UNIT,DVCD,ISNULL(SUM(QNTY),0) AS WASTAGE_QNTY FROM PKGMAN " & _
         "WHERE DBCD='000006' AND RECSTAT<>'D' " & _
         " AND DATE >='" & Format(dtFrom.Text, "MM/DD/YYYY") & _
         "' AND DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' GROUP BY COMP,UNIT,DVCD "
         
CN.Execute "IF ( OBJECT_ID('VW_WASTAGE') IS NOT NULL ) DROP VIEW VW_WASTAGE "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

Private Sub SetViewForMachineWastage()
Dim QRY As String
On Error GoTo VWERR
   
   QRY = "CREATE VIEW VW_WASTAGE AS " & _
         "SELECT COMP,UNIT,DVCD,MCCD,ISNULL(SUM(GRSWGT),0) AS WASTAGE_QNTY FROM BOXREGISTER " & _
         "WHERE DBCD='000006' AND RECSTAT<>'D' " & _
         " AND VBDT >='" & Format(dtFrom.Text, "MM/DD/YYYY") & _
         "' AND VBDT <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' GROUP BY COMP,UNIT,DVCD,MCCD " & _
         "UNION " & _
         "SELECT COMP,UNIT,DVCD,MCCD,ISNULL(SUM(QNTY),0) AS WASTAGE_QNTY FROM PKGMAN " & _
         "WHERE DBCD='000006' AND RECSTAT<>'D' " & _
         " AND DATE >='" & Format(dtFrom.Text, "MM/DD/YYYY") & _
         "' AND DATE <='" & Format(dtTo.Text, "MM/DD/YYYY") & "' GROUP BY COMP,UNIT,DVCD,MCCD "
         
CN.Execute "IF ( OBJECT_ID('VW_WASTAGE') IS NOT NULL ) DROP VIEW VW_WASTAGE "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

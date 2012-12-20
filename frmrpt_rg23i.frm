VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrpt_rg23i 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RG23-I"
   ClientHeight    =   3525
   ClientLeft      =   5580
   ClientTop       =   4260
   ClientWidth     =   4620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4620
   Begin VB.Frame Frame8 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   4860
      Visible         =   0   'False
      Width           =   4560
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   3285
      End
      Begin VB.Label Label14 
         Caption         =   "Division :"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   2745
      Width           =   4560
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Text            =   "100"
         Top             =   270
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3600
         TabIndex        =   15
         Top             =   210
         Width           =   915
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&View"
         Height          =   405
         Left            =   2520
         TabIndex        =   14
         Top             =   210
         Width           =   915
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2400
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label13 
         Caption         =   "Report Zoom %"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   0
      TabIndex        =   6
      Top             =   1695
      Width           =   4560
      Begin VB.ComboBox txtitem 
         Height          =   315
         ItemData        =   "frmrpt_rg23i.frx":0000
         Left            =   1200
         List            =   "frmrpt_rg23i.frx":0002
         TabIndex        =   10
         Top             =   600
         Width           =   3330
      End
      Begin VB.TextBox txtParty 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4185
         Width           =   3210
      End
      Begin VB.ComboBox cboFormat 
         Height          =   315
         ItemData        =   "frmrpt_rg23i.frx":0004
         Left            =   1200
         List            =   "frmrpt_rg23i.frx":0006
         TabIndex        =   8
         Top             =   240
         Width           =   3330
      End
      Begin VB.TextBox txtGroup 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3345
         Width           =   3210
      End
      Begin VB.TextBox txtCategory 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3765
         Width           =   3210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chapter No."
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Party Name :"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   4245
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Format :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   277
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Group :"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   3405
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Category :"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   3825
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   0
      TabIndex        =   16
      Top             =   930
      Width           =   4560
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   2955
         TabIndex        =   17
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   18284545
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1050
         TabIndex        =   18
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   18284545
         CurrentDate     =   38429
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2595
         TabIndex        =   20
         Top             =   255
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "From :"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4560
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   255
         Width           =   3285
      End
      Begin VB.Label Label8 
         Caption         =   "Unit Name :"
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmrpt_rg23i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboFormat_LostFocus()
 cboFormat.BackColor = vbWhite
End Sub

Private Sub cmdpreview_Click()
    CRPT.Reset
    crptConnect CRPT
    
    Call CollectStockData
    
    If txtUNIT = Empty Then
        MsgBox "Please Select Unit !!", vbInformation, "Unit Is Key Field Missing"
        txtUNIT.SetFocus
    End If

    rptsql = Empty
    ReportName = Empty
    
    Select Case cboFormat.ListIndex
        Case 0
            ReportName = App.PATH & "\Reports\RPT_RG23A-I.RPT"
            RPTN = "RG23-I Register"
        Case 1
            ReportName = App.PATH & "\Reports\RPT_STORE_ITM_PERIOD_INOUT.rpt"
            RPTN = "STOCK INWARD / OUTWARD ONLY REPORT"
        Case Else
            MsgBox "Please Select Valid Report Format !!", vbInformation, "Wrong Report Format"
            Exit Sub
    End Select
    
    If ReportName = Empty Then
        MsgBox "Report Is Not Configured !!", vbInformation, "Under Development"
        Exit Sub
    End If
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    If cboFormat.ListIndex = 0 Then
        rptsql = "{STORETRAN.COMP}='" & compPth & "' AND {STORETRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {STORETRAN.UNIT}='" & txtUNIT.Tag & "' AND {STORETRAN.RECSTAT}='A' AND ({STORETRAN.OPER}='+' OR {STORETRAN.OPER}='-') AND {STORETRAN.DVCD}='000001'  "
    Else
        rptsql = "{STORETRAN.COMP}='" & compPth & "' AND {STORETRAN.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {STORETRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {STORETRAN.UNIT}='" & txtUNIT.Tag & "' AND {STORETRAN.RECSTAT}='A' AND {STORETRAN.DVCD}='000001' "
    End If
    
    If Not txtitem = Empty Then rptsql = rptsql & " AND {IGMMST.CHAP}='" & txtitem.Text & "'"
    If Not txtCategory = Empty Then rptsql = rptsql & " AND {IGMMST.IHCD} IN [" & txtCategory.Tag & "]"
    If Not txtGroup = Empty Then rptsql = rptsql & " AND {IGMMST.CODE} IN [" & txtGroup.Tag & "]"
    If Not txtParty = Empty Then rptsql = rptsql & " AND {STORETRAN.PCOD} IN [" & txtParty.Tag & "]"
    
    PERIOD = dtFrom & " To " & dtTo
    
    CRPT.ReportFileName = ReportName
    
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
        .Formulas(4) = "UNIT='" & txtUNIT & "'"
        '.Formulas(5) = "DIVISION='" & txtDVCD & "'"
        .Formulas(5) = "OPDT=#" & Format(DateAdd("D", -1, dtFrom), "MM/dd/yyyy") & "#"
        
        RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000080", 8, "R") Then
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
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
End Sub

Private Sub Form_Activate()
    Call ColorComponent(Me)
    If Not txtUNIT = Empty Then Exit Sub
    Call txtUnit_KeyDown(vbKeyF2, 0)
    If txtUNIT = Empty Then cmdPreview.Enabled = False
    If txtitem.ListCount > 0 Then txtitem.ListIndex = 0
    If cboFormat.ListCount > 0 Then cboFormat.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtDVCD = Empty And ActiveControl.NAME = "txtDVCD" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCategory_GotFocus()
 txtCategory.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtCategory_LostFocus()
 txtCategory.BackColor = vbWhite
End Sub

Private Sub txtGroup_GotFocus()
 txtGroup.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtGroup_LostFocus()
 txtGroup.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
 txtitem.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub



Private Sub cboFormat_Click()
    SendKeys "{END}+{HOME}"
End Sub

Private Sub cboFormat_GotFocus()
 cboFormat.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "%{DOWN}"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    Dim PRM As String
    
    Call CenterChild(frm_Main, Me)
    
    dtFrom = FSDT
    dtTo = GetMaxDate
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    
    With cboFormat
        .AddItem "Chapter Wise Stock Register"
        .ListIndex = 0
    End With
    If RS.State = 1 Then RS.Close
    RS.Open "select distinct chap from igmmst", CN, adOpenDynamic, adLockOptimistic
    txtitem.Clear
    Do While Not RS.EOF
     If Trim(RS!CHAP & "") = "" Then
      Else
       txtitem.AddItem RS!CHAP & ""
     End If
     RS.MoveNext
    Loop
    
    If txtitem.ListCount > 0 Then txtitem.ListIndex = 0
    If cboFormat.ListCount > 0 Then cboFormat.ListIndex = 0
End Sub

Private Sub txtCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
'        txtCategory = Empty
'        txtCategory.Tag = Empty
'        M_DESC = Empty
'        NEW_VISIBLE = False
'        txtCategory = SearchItemList("Select TOP 20 Code,Name From SCAT_MST", 0, Empty, "Select Denier !!")
'        txtCategory.Tag = Key
        MSelCata = "C"
        LOAD frm_MulSelectCode
        frm_MulSelectCode.Show 1
        
        If Len(MSelName) > 0 Then
            txtCategory.Text = Trim(MSelName)
            txtCategory.Tag = Trim(MSelCode)
        End If
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtCategory = Empty
        txtCategory.Tag = Empty
    End If
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' and Unit='" & txtUNIT.Tag & "'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
End Sub

Private Sub txtGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
'        txtGroup = Empty
'        txtGroup.Tag = Empty
'        M_DESC = Empty
'        NEW_VISIBLE = False
'        txtGroup = SearchItemList("Select TOP 20 Code,Name From IGMMST", 0, Empty, "Select Denier !!")
'        txtGroup.Tag = Key
        MSelCata = "G"
        LOAD frm_MulSelectCode
        frm_MulSelectCode.Show 1
        
        If Len(MSelName) > 0 Then
            txtGroup.Text = Trim(MSelName)
            txtGroup.Tag = Trim(MSelCode)
        End If
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtGroup = Empty
        txtGroup.Tag = Empty
    End If
End Sub

Private Sub txtItem_LostFocus()
 txtitem.BackColor = vbWhite
End Sub

Private Sub txtParty_GotFocus()
 txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
'        NEW_VISIBLE = False
'        M_DESC = Empty
'        Key = Empty
'        txtParty.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM ACCMST", 0)
'        txtParty.Tag = Key
        MSelCata = "P"
        LOAD frm_MulSelectCode
        frm_MulSelectCode.Show 1
        
        If Len(MSelName) > 0 Then
            txtParty.Text = Trim(MSelName)
            txtParty.Tag = Trim(MSelCode)
        End If
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtParty.Tag = Empty
        txtParty = Empty
    End If
End Sub

Private Sub txtParty_LostFocus()
 txtParty.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
 txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
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


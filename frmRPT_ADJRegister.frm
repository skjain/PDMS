VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_ADJRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjustment Register"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6810
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   19
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
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
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
         Style           =   2  'Dropdown List
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
         TabIndex        =   18
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
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   6615
      Begin VB.TextBox txtItem 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
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
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Item"
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
         Top             =   960
         Width           =   1095
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   6615
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   8
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
         TabIndex        =   9
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
         Image           =   "frmRPT_ADJRegister.frx":0000
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
         Image           =   "frmRPT_ADJRegister.frx":0452
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
Attribute VB_Name = "frmRPT_ADJRegister"
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
       ReportName = App.PATH & "\Reports\Date wise Adjustment Register.rpt"
       RPTN = "DATE WISE ADJUSTMENT REGISTER"
       If txtMACHINE <> Empty Then RPTN = RPTN & " ( " & txtMACHINE & ")"
    ElseIf cboReports.ListIndex = 1 Then
       ReportName = App.PATH & "\Reports\Division Wise Adjustment Register.rpt"
       RPTN = "DIVISION WISE ADJUSTMENT REGISTER"
    ElseIf cboReports.ListIndex = 2 Then
       ReportName = App.PATH & "\Reports\Division+Machine wise Adjustment Register.rpt"
       RPTN = "DIVISION + MACHINE WISE ADJUSTMENT REGISTER"
    End If

    rptsql = Empty
    rptsql = "{STORETRAN.COMP}='" & compPth & "' AND {STORETRAN.UNIT} = '" & txtUNIT.Tag & "' AND " & _
    "{STORETRAN.DVCD}<>'000001' AND {STORETRAN.VTYP}='WIP' AND {STORETRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") " & _
    "AND {STORETRAN.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {STORETRAN.RECSTAT}<>'D' "
    
    If TXTDVCD <> Empty Then rptsql = rptsql & " AND {STORETRAN.DVCD}='" & TXTDVCD.Tag & "'"
    If txtMACHINE <> Empty Then rptsql = rptsql & " AND {STORETRAN.PCOD}='" & txtMACHINE.Tag & "'"
    If txtItem <> Empty Then rptsql = rptsql & " AND {STORETRAN.ICOD}='" & txtItem.Tag & "'"
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
 With cboReports
   .AddItem "Date Wise Adjustment Register"
   .AddItem "Division Wise Adjustment Register"
   .AddItem "Division + M/c Wise Adjustment Register"
   cboReports.ListIndex = 0
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
End Sub

Private Sub txtDVCD_GotFocus()
 TXTDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
'If txtUNIT = Empty Then txtUNIT.Enabled = True: txtUNIT.SetFocus: Exit Sub

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

Private Sub txtItem_GotFocus()
    txtItem.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtItem = SearchList1("Select  TOP 20 CODE,NAME From ITMMST Where COMP='" & compPth & "'", 0, Empty, "Select Item from List")
        txtItem.Tag = Key
        
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    
        txtItem = Empty
        txtItem.Tag = Empty
        
    End If
End Sub

Private Sub txtItem_LostFocus()
    txtItem.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
  txtMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        'If TXTDVCD = Empty Then TXTDVCD.Enabled = True: TXTDVCD.SetFocus: Exit Sub
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

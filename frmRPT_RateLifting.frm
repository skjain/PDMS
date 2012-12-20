VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_RateLifting 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Average Rate Realisation Report"
   ClientHeight    =   4845
   ClientLeft      =   2940
   ClientTop       =   870
   ClientWidth     =   6660
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   120
      TabIndex        =   25
      Top             =   1920
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   3960
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
         TabIndex        =   17
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3720
         TabIndex        =   18
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
         Image           =   "frmRPT_RateLifting.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   19
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
         Image           =   "frmRPT_RateLifting.frx":0452
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
         TabIndex        =   16
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   6495
      Begin VB.TextBox txtParty 
         Enabled         =   0   'False
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
         TabIndex        =   13
         Top             =   600
         Width           =   4845
      End
      Begin VB.TextBox txtItem 
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
         TabIndex        =   15
         Top             =   960
         Width           =   4845
      End
      Begin VB.TextBox TXTDCOD 
         Enabled         =   0   'False
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
         TabIndex        =   11
         Top             =   240
         Width           =   4845
      End
      Begin VB.Label LblParty 
         BackStyle       =   0  'Transparent
         Caption         =   "Consigner "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label LBLConsignee 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame11 
      Height          =   585
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   6495
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1560
         TabIndex        =   5
         Top             =   180
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
         TabIndex        =   7
         Top             =   180
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
         TabIndex        =   6
         Top             =   210
         Width           =   735
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
         TabIndex        =   4
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame Frame10 
      Height          =   585
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   6495
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
         ForeColor       =   &H8000000C&
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
      TabIndex        =   20
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
Attribute VB_Name = "frmRPT_RateLifting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RPTN As String
Dim PERIOD As String

Private Sub cboFormats_Click()

Select Case cboFormats.ListIndex
Case 0
     TXTDCOD.Enabled = True
     txtParty.Enabled = False
     LBLConsignee.Enabled = True
     LblParty.Enabled = False
Case 1
     txtParty.Enabled = True
     TXTDCOD.Enabled = False
     LBLConsignee.Enabled = False
     LblParty.Enabled = True
End Select

End Sub

Private Sub dtFrom_Validate(Cancel As Boolean)
    If Not IsDate(dtFrom) And dtFrom <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtFrom.SetFocus
    End If
End Sub

Private Sub dtTo_Validate(Cancel As Boolean)
    If Not IsDate(dtTo) And dtTo <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtTo.SetFocus
    End If
End Sub

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
    End If
        
    If cboFormats.ListIndex = -1 Then cboFormats.ListIndex = 0
     
    CRPT.Reset
    crptConnect CRPT
    
    rptsql = Empty
    ReportName = Empty
    
    Select Case cboFormats.ListIndex
       Case 0
           ReportName = App.PATH & "\Reports\RPT_Consigneewise_Avg_RateRealisation.RPT"
           RPTN = "Consignee wise Average Rate Realisation Report"
           Call GenViewForConsignee
       Case 1
           ReportName = App.PATH & "\Reports\RPT_Consignerwise_Avg_RateRealisation.RPT"
           RPTN = "Consigner wise Average Rate Realisation Report"
           Call GenViewForConsigner
    End Select
        
    RPTN = RPTN
        
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
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        .Formulas(3) = "DIVISION='" & txtDVCD & "'"
        .Formulas(4) = "PERIOD='" & PERIOD & "'"
        .Formulas(5) = "REPORTHEAD='" & RPTN & "'"
        
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
         .AddItem "Consignee+Item wise Average Rate Realisation Report"
         .AddItem "Consigner+Item wise Average Rate Realisation Report"
    End With
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
  If txtDVCD = Empty Then txtDVCD.SetFocus: Exit Sub

  If KeyCode = vbKeyF2 Then
     NEW_VISIBLE = False
     CANCEL_VISIBLE = True
     M_DESC = Empty
     txtITEM = SearchList1("Select  TOP 20 Code,Name From FINITMMST WHERE COMP='" & compPth & _
                              "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Dealer From List")
     txtITEM.Tag = Key
  ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtITEM = Empty
     txtITEM.Tag = Empty
  End If
  
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

Private Sub TXTDCOD_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTDCOD = Empty
  ElseIf KeyCode = vbKeyF2 Then
     M_DESC = Empty:   NEW_VISIBLE = False
     TXTDCOD = SearchList1("Select DISTINCT CODE,NAME From PADDMST", 0, Empty, "Select Consignee Name ")
     TXTDCOD.Tag = Key
  End If
  Me.KeyPreview = True
End Sub

Private Sub txtParty_LostFocus()
 txtParty.BackColor = vbWhite
End Sub

Private Sub txtParty_GotFocus()
 txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True:  M_DESC = Empty
        txtParty = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select Consigner From List")
        txtParty.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtParty = Empty
        txtParty.Tag = Empty
    End If
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

Private Sub TXTDCOD_GotFocus()
 TXTDCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDCOD_LostFocus()
 TXTDCOD.BackColor = vbWhite
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

Private Sub txtItem_LostFocus()
 txtITEM.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
 txtITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
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

If txtPCOD <> Empty Then QRY = QRY & " AND SPTRAN.PCOD='" & txtParty.Tag & "' "
If txtITEM <> Empty Then QRY = QRY & " AND SPTRAN.ICOD='" & txtITEM.Tag & "' "

QRY = QRY & " GROUP BY SPTRAN.COMP,SPTRAN.UNIT,ACCMST.NAME,FINITMMST.NAME"
       
CN.Execute "IF ( OBJECT_ID('VW_PARTY_ITEMWISE_RATE_LIFTING') IS NOT NULL ) DROP VIEW VW_PARTY_ITEMWISE_RATE_LIFTING "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

Private Sub GenViewForConsignee()
Dim QRY As String
On Error GoTo VWERR

QRY = "CREATE VIEW VW_CONSIGNEE_ITEMWISE_RATE_LIFTING AS " & _
   "SELECT SPTRAN.COMP,SPTRAN.UNIT,PADDMST.NAME AS PARTY,FINITMMST.NAME AS DENIER, " & _
   "ISNULL(SUM(SPTRAN.QNTY * SPTRAN.RATE),0) / ISNULL(SUM(SPTRAN.QNTY),1) AS RATE  FROM SPTRAN " & _
   "INNER JOIN PADDMST ON PADDMST.CODE = SPTRAN.DCOD " & _
   "INNER JOIN FINITMMST ON FINITMMST.COMP = SPTRAN.COMP AND FINITMMST.UNIT = SPTRAN.UNIT " & _
   "AND FINITMMST.DVCD = SPTRAN.DVCD AND FINITMMST.CODE = SPTRAN.ICOD " & _
   "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & txtUNIT.Tag & _
   "' AND SPTRAN.VTYP='DPF' AND SPTRAN.DVCD='" & txtDVCD.Tag & _
   "' AND SPTRAN.DATE >='" & Format(dtFrom.Text, "MM/DD/YYYY") & _
   "' AND SPTRAN.DATE<='" & Format(dtTo.Text, "MM/DD/YYYY") & "' AND SPTRAN.RECSTAT<>'D' "

If TXTDCOD <> Empty Then QRY = QRY & " AND SPTRAN.DCOD='" & TXTDCOD.Tag & "' "
If txtITEM <> Empty Then QRY = QRY & " AND SPTRAN.ICOD='" & txtITEM.Tag & "' "

QRY = QRY & " GROUP BY SPTRAN.COMP,SPTRAN.UNIT,PADDMST.NAME,FINITMMST.NAME"
       
CN.Execute "IF ( OBJECT_ID('VW_CONSIGNEE_ITEMWISE_RATE_LIFTING') IS NOT NULL ) DROP VIEW VW_CONSIGNEE_ITEMWISE_RATE_LIFTING "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub



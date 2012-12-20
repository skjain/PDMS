VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_LOTMASTER 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOT  MASTER"
   ClientHeight    =   1785
   ClientLeft      =   1965
   ClientTop       =   2145
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000B&
      Height          =   1080
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5730
      Begin VB.TextBox txtItem 
         Height          =   315
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   4380
      End
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   165
         Width           =   4395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "&Item Name :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Division :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   645
      End
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   0
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin WelchButton.lvButtons_H cmdpreview 
      Height          =   405
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
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
      Image           =   "frmRpt_LOTMASTER.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   4440
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
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
      Image           =   "frmRpt_LOTMASTER.frx":0452
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmRPT_LOTMASTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPrintReport
    
    Dim I As Long
    
    CRPT.Reset
    crptConnect CRPT
    
    If Me.Tag = "LOT" Then
        ReportName = App.PATH & "\Reports\RPT_LOT_MASTER.rpt"
        rptsql = " {TXULOT.COMP}='" & compPth & "' AND {TXULOT.UNIT}='" & UNCD & "' "
    Else
        ReportName = App.PATH & "\Reports\RPT_FIN_ITM_MST.rpt"
        rptsql = " {FINITMMST.COMP}='" & compPth & "' AND {FINITMMST.UNIT}='" & UNCD & "' "
    End If
    
    If Not txtDVCD.Text = Empty Then rptsql = rptsql & " AND {DIVMST.CODE}='" & Trim(txtDVCD.Tag) & "'"
    If Not txtITEM = Empty Then rptsql = rptsql & " AND {FINITMMST.CODE}='" & txtITEM.Tag & "'"
        
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.DiscardSavedData = True
    CRPT.Destination = crptToWindow
    CRPT.WindowState = crptMaximized
    CRPT.WindowShowProgressCtls = True
    CRPT.WindowShowPrintBtn = True
    CRPT.WindowShowPrintSetupBtn = True
    CRPT.WindowShowRefreshBtn = True
    CRPT.WindowShowSearchBtn = True
    
    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula rptsql
    
    CRPT.WindowShowExportBtn = True
    CRPT.WindowTitle = " Master Report" & Space(5) & "Report : " & ReportName
    CRPT.PageLast
    CRPT.PageFirst
    CRPT.ACTION = 1

    Exit Sub
    
errPrintReport:
    
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description & vbCrLf & " Error In Report " & ReportName
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
On Error GoTo errLoad
    Call CenterChild(frm_Main, Me)
    Me.KeyPreview = True
    Me.Tag = RPTPARA
    
    If Me.Tag = "LOT" Then
    'ok
    Else
    Me.Caption = "FINSH ITEM MASTER"
    End If
    
    Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub txtDVCD_GotFocus()
txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Public Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & UNCD & "' AND RECSTAT='A'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
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
        If txtDVCD.Text = Empty Then
        txtITEM = SearchList1("Select TOP 20 Code,Name From FINITMMST", 0, Empty, "Select Item")
        Else
        txtITEM = SearchList1("Select TOP 20 Code,Name From FINITMMST WHERE DVCD='" & Trim(txtDVCD.Tag) & "'", 0, Empty, "Select Item")
        End If
        txtITEM.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtITEM = Empty
    End If
End Sub

Private Sub txtItem_LostFocus()
txtITEM.BackColor = vbWhite
End Sub

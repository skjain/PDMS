VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRPT_PartyWiseReturnableReg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PartyWise Returnable Metallic Cops/Pallet Register"
   ClientHeight    =   3360
   ClientLeft      =   2895
   ClientTop       =   2790
   ClientWidth     =   5895
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame9 
      Height          =   1005
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   5700
      Begin VB.TextBox txtAGENT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   4365
      End
      Begin VB.TextBox TXTPARTY 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   4365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent  :"
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
         TabIndex        =   8
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Party :"
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
         TabIndex        =   6
         Top             =   300
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   5700
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Text            =   "100"
         Top             =   255
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4440
         TabIndex        =   13
         Top             =   255
         Width           =   1140
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Width           =   1140
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2385
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Zoom %"
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
         TabIndex        =   10
         Top             =   315
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   105
      TabIndex        =   17
      Top             =   45
      Width           =   5700
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50397185
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50397185
         CurrentDate     =   38429
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   308
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   308
         Width           =   1005
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   120
      TabIndex        =   16
      Top             =   750
      Width           =   5700
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   4365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   308
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmRPT_PartyWiseReturnableReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As ADODB.Recordset
Dim rptsql As String

Private Sub Form_KeyDown(KeyCode As Integer, SHIFT As Integer)
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = vbKeyReturn Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    Call CenterChild(frm_Main, Me)
    dtFrom = GetMinDate
    dtTo = Date
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()

If txtUNIT.Text = Empty Then
   MsgBox "Please Select Unit !!", vbInformation, "Key Field Unit Is Missing"
   txtUNIT.SetFocus
   Exit Sub
End If

    CRPT.Reset
    crptConnect CRPT

ReportName = Empty
rptsql = Empty

PERIOD = dtFrom & " To " & dtTo

   ReportName = App.PATH & "\Reports\PartywiseReturnableCopsRegister.rpt"
   RPTN = "PARTY WISE RETURNABLE REGISTER  "

If Dir(ReportName, vbNormal) = Empty Then
   ReportErrorMessage 1001
   Exit Sub
End If
    
rptsql = "{PKGSTK.COMP}='" & compPth & "' AND {PKGSTK.UNIT} IN [" & txtUNIT.Tag & _
            "] AND {PKGSTK.DATE}>=DATE(" & Year(dtFrom) & _
            "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {PKGSTK.DATE}<=DATE(" & _
            Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")"
    
    
    If TXTPARTY <> Empty Then rptsql = rptsql & " AND {PKGSTK.PCOD} = '" & TXTPARTY.Tag & "'"
    If txtAGENT <> Empty Then rptsql = rptsql & " AND {PKGSTK.BRCD} = '" & txtAGENT.Tag & "'"
        
    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
        .Formulas(4) = "UNIT='" & txtUNIT.Text & "'"
        
        RPTN = RPTN + Space(5) + ReportName
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    
    Exit Sub
    
errPreview:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub TXTPARTY_KeyDown(KeyCode As Integer, SHIFT As Integer)
   If KeyCode = vbKeyF2 Then
     TXTPARTY.Text = SearchList1("SELECT CODE,NAME FROM ACCMST ", 0, TXTPARTY.Text, "Select Party")
     TXTPARTY.Tag = Key
   ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTPARTY.Text = Empty
     TXTPARTY.Tag = Empty
   End If
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, SHIFT As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        LOAD frm_askunit
        If frm_askunit.LSTUNIT.ListCount > 0 Then
            frm_askunit.Show 1
        End If
        txtUNIT = sel_untnam
        txtUNIT.Tag = sel_untcod
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If
End Sub

Private Sub TXTAGENT_KeyDown(KeyCode As Integer, SHIFT As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtAGENT = SearchList1("Select TOP 20 Code,Name From REFMST WHERE CATA='B'", 0, Empty, "Select AGENT")
        txtAGENT.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtAGENT = Empty
        txtAGENT.Tag = Empty
    End If
End Sub

Private Sub TXTPARTY_GotFocus():  TXTPARTY.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTPARTY_LostFocus():  TXTPARTY.BackColor = vbWhite: End Sub

Private Sub TXTAGENT_GotFocus(): txtAGENT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTAGENT_LostFocus(): txtAGENT.BackColor = vbWhite: End Sub

Private Sub TXTZOOM_GotFocus():: txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTZOOM_LostFocus(): txtZoom.BackColor = vbWhite: End Sub

Private Sub txtUNIT_GotFocus(): txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtUNIT_LostFocus():  txtUNIT.BackColor = vbWhite: End Sub

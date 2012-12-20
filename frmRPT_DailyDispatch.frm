VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_DailyDispatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consignee/Consigner Wise Lifting Report"
   ClientHeight    =   5130
   ClientLeft      =   2940
   ClientTop       =   870
   ClientWidth     =   6675
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "As On :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   2160
      Begin MSComCtl2.DTPicker dtAson 
         Height          =   330
         Left            =   600
         TabIndex        =   13
         Top             =   300
         Width           =   1410
         _ExtentX        =   2487
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
         Format          =   53542913
         CurrentDate     =   39343
      End
      Begin VB.Label Label2 
         Caption         =   "Da&te :"
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
         Top             =   300
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   120
      TabIndex        =   22
      Top             =   1320
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   4320
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
         TabIndex        =   15
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3720
         TabIndex        =   16
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
         Image           =   "frmRPT_DailyDispatch.frx":0000
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
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Image           =   "frmRPT_DailyDispatch.frx":0452
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
         TabIndex        =   14
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   120
      TabIndex        =   20
      Top             =   1920
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
         TabIndex        =   9
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
         TabIndex        =   11
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   10
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
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame10 
      Height          =   585
      Left            =   120
      TabIndex        =   19
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
      TabIndex        =   18
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
Attribute VB_Name = "frmRPT_DailyDispatch"
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
    dtAson = Now
    Call SetReportFormat
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub cmdPreview_Click()
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
    
    rptsql = "{SPTRAN.COMP}='" & compPth & "' AND {SPTRAN.UNIT}='" & txtUNIT.Tag & _
    "' AND {SPTRAN.DVCD}='" & txtDVCD.Tag & "' AND {SPTRAN.VTYP}='DPF' And {SPTRAN.DATE}>=DATE(" & Year(dtAson) & _
    "," & Month(dtAson) & ",1) AND {SPTRAN.DATE}<=DATE(" & Year(dtAson) & "," & Month(dtAson) & _
    "," & Day(dtAson) & ") And {SPTRAN.RECSTAT}<>'D' "
    
    If TXTDCOD <> Empty Then rptsql = rptsql & "AND {SPTRAN.DCOD}='" & TXTDCOD.Tag & "' "
    If txtParty <> Empty Then rptsql = rptsql & "AND {SPTRAN.PCOD}='" & txtParty.Tag & "' "
    If txtItem <> Empty Then rptsql = rptsql & "AND {SPTRAN.ICOD}='" & txtItem.Tag & "' "
        
    Select Case cboFormats.ListIndex
       Case 0
           ReportName = App.PATH & "\Reports\ConsigneeWiseDailyLifting.rpt"
           RPTN = "Consignee wise Daily Lifting Report"
       Case 1
           ReportName = App.PATH & "\Reports\ConsignerWiseDailyLifting.rpt"
           RPTN = "Consigner wise Daily Lifting Report"
    End Select
        
    RPTN = RPTN
        
    If ReportName = Empty Then
        ReportErrorMessage 0
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
    
    CRPT.ReplaceSelectionFormula rptsql
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
          
    If Len(Month(dtAson)) = 1 Then
       PERIOD = "01" & "/0" & Month(dtAson) & "/" & Year(dtAson) & " To " & Format(dtAson, "dd/mm/yyyy")
    Else
       PERIOD = "01" & "/" & Month(dtAson) & "/" & Year(dtAson) & " To " & Format(dtAson, "dd/mm/yyyy")
    End If
    
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
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub SetReportFormat()
    With cboFormats
         .Clear
         .AddItem "Consigneewise Daily Lifting"
         .AddItem "Consignerwise Daily Lifting"
    End With
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
  If txtDVCD = Empty Then txtDVCD.SetFocus: Exit Sub

  If KeyCode = vbKeyF2 Then
     NEW_VISIBLE = False
     CANCEL_VISIBLE = True
     M_DESC = Empty
     txtItem = SearchList1("Select  TOP 20 Code,Name From FINITMMST WHERE COMP='" & compPth & _
                              "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Dealer From List")
     txtItem.Tag = Key
  ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtItem = Empty
     txtItem.Tag = Empty
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

Private Sub TXTParty_GotFocus()
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

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub TXTITEM_LostFocus()
 txtItem.BackColor = vbWhite
End Sub

Private Sub TXTITEM_GotFocus()
 txtItem.BackColor = RGB(BRED, BGREEN, BBLUE)
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


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FRMRPT_PCKSMRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing Summary"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6435
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   6255
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
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4500
      End
      Begin VB.TextBox TXTDVCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   4500
      End
      Begin VB.TextBox TXTPACKINGTYPE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   4500
      End
      Begin VB.Label lblUnit 
         Caption         =   "&Unit Name "
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
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "&Packing Type"
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
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1440
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1035
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   6285
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
         TabIndex        =   22
         Text            =   "100"
         Top             =   300
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2640
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3240
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Preview"
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
         Image           =   "FRMRPT_PCKSMRY.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4560
         TabIndex        =   24
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
         Image           =   "FRMRPT_PCKSMRY.frx":0D8A
         cBack           =   -2147483633
      End
      Begin VB.Label Label13 
         Caption         =   "Report &Zoom %"
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
         TabIndex        =   21
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   6285
      Begin VB.TextBox TXTSUBGRD 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox TXTMACHINE 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   4560
      End
      Begin VB.TextBox txtDENIER 
         Height          =   315
         Left            =   1545
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   4560
      End
      Begin VB.TextBox txtLTNO 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtGRADE 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label LBLSUBGRD 
         Caption         =   "&Sub Grade"
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
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "&Machine Name"
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
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label11 
         Caption         =   "Finish &Item "
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
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label15 
         Caption         =   "L&ot No."
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
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label8 
         Caption         =   "&Grade"
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
         TabIndex        =   19
         Top             =   1320
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6255
      Begin VB.ComboBox cmbFormat 
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
         TabIndex        =   8
         Top             =   240
         Width           =   4545
      End
      Begin MSMask.MaskEdBox DTFROM 
         Height          =   330
         Left            =   1560
         TabIndex        =   10
         Top             =   720
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
      Begin MSMask.MaskEdBox DTTO 
         Height          =   330
         Left            =   4785
         TabIndex        =   12
         Top             =   720
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
      Begin VB.Label Label6 
         Caption         =   "F&rom Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "T&o Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   11
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label7 
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
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FRMRPT_PCKSMRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_DBCD As String
Dim PERIOD As String
Dim M_MCNO As String
Dim PACK As String

Private Sub cmbFormat_Click()
If cmbFormat.ListIndex = 6 Then
   txtDENIER = Empty
   TXTMACHINE = Empty
   txtLTNo = Empty
   txtGrade = Empty
   TXTSUBGRD = Empty
   
   txtDENIER.Enabled = False
   TXTMACHINE.Enabled = False
   txtLTNo.Enabled = False
   txtGrade.Enabled = False
   TXTSUBGRD.Enabled = False
Else
   txtDENIER.Enabled = True
   TXTMACHINE.Enabled = True
   txtLTNo.Enabled = True
   txtGrade.Enabled = True
   TXTSUBGRD.Enabled = True
End If
End Sub

Private Sub CMBFORMAT_GotFocus()
 cmbFormat.BackColor = RGB(BRED, BGREEN, BBLUE)
  With cmbFormat
            .Clear
            cmbFormat.AddItem "Machine+Item+Lot+Grade Wise Packing Summary"
            cmbFormat.AddItem "Machine+Item+Grade Wise Packing Summary"
            cmbFormat.AddItem "Item+Month Wise Packing Summary"
            cmbFormat.AddItem "Item+Grade Wise Packing Summary"
            cmbFormat.AddItem "Machine+Item+Grade Wise Packing Summary {%}"
            cmbFormat.AddItem "Date and M/c Wise Packing Summary"
            
            If SetIsShadeReq(txtDVCD.Tag) = "Y" Then
               .AddItem "Machine+Denier+Lot+Shade+Grade Wise Packing Summary"
            End If
            
            cmbFormat.AddItem "Grade+SubGrade Wise Packing Summary {%}"
            
            .ListIndex = 0
    End With
End Sub

Private Sub cmbFormat_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then dtFrom.SetFocus
End Sub

Private Sub CMBFORMAT_LostFocus()
  cmbFormat.BackColor = vbWhite
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
Dim FLAG As Boolean
Dim PKG As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim M_DVCD As String

Dim PKGTYP As String
PKGTYP = PackingType(txtDVCD.Tag)

Call SetViewForPackingSummary

    CRPT.Reset
    crptConnect CRPT
    
    If Not RIGHTDATA Then Exit Sub
    
    M_MCNO = Empty
    rptsql = Empty
  
  Call SetFormula(PKGTYP)
  
      If cmbFormat.ListIndex = 0 Then
            If PKGTYP = "C" Then
               ReportName = App.PATH & "\Reports\Machine+Item+Lot+Grade WisePkgBoxSummary.rpt"
            Else
               ReportName = App.PATH & "\Reports\Machine+Item+Lot+Grade WisePkgSummary.rpt"
            End If
        RPTN = "MACHINE + ITEM + LOT + GRADE WISE PACKING SUMMARY"
     ElseIf cmbFormat.ListIndex = 1 Then
           If PKGTYP = "C" Then
              ReportName = App.PATH & "\Reports\Machine+Item+Grade+WisePkgBoxSummary.rpt"
           Else
              ReportName = App.PATH & "\Reports\Machine+Item+Grade+WisePkgSummary.rpt"
           End If
        RPTN = "MACHINE + ITEM + GRADE WISE PACKING SUMMARY"
     ElseIf cmbFormat.ListIndex = 2 Then
           If PKGTYP = "C" Then
              ReportName = App.PATH & "\Reports\Item+Month+WisePkgBoxSummary.rpt"
           Else
              ReportName = App.PATH & "\Reports\Item+Month+WisePkgSummary.rpt"
           End If
        RPTN = "ITEM + MONTH WISE PACKING SUMMARY"
     ElseIf cmbFormat.ListIndex = 3 Then
           If PKGTYP = "C" Then
              ReportName = App.PATH & "\Reports\Item+Grade+WisePkgBoxSummary.rpt"
           Else
              ReportName = App.PATH & "\Reports\Item+Grade+WisePkgSummary.rpt"
           End If
           RPTN = "ITEM + GRADE WISE PACKING SUMMARY"
     ElseIf cmbFormat.ListIndex = 4 Then
            If PKGTYP = "C" Then
               ReportName = App.PATH & "\Reports\Machine+Item+Grade_SubGrade WisePkgBoxSummary.rpt"
            Else
               MsgBox "Lumpsump Production Summary Report Under Development"
               Exit Sub
            End If
            RPTN = "Machine+Item+Grade+SubGrade Wise Production Summary"
     ElseIf cmbFormat.ListIndex = 5 Then
            If PKGTYP = "C" Then
               ReportName = App.PATH & "\Reports\Datewise+MachinewisePkgBoxSummary.RPT"
            Else
               MsgBox "Lumpsump Production Summary Report Under Development"
               Exit Sub
            End If
            RPTN = "Machine+Item+Grade+SubGrade Wise Production Summary"
     ElseIf cmbFormat.ListIndex = 6 Then
            If PKGTYP = "C" Then
               ReportName = App.PATH & "\Reports\Grade_SubGrade WisePkgBoxSummary.rpt"
            Else
               MsgBox "Lumpsump Production Summary Report Under Development"
               Exit Sub
            End If
            RPTN = "Grade+SubGrade Wise Production Summary"
     End If
          
    If Trim(cmbFormat.Text) = "Machine+Denier+Lot+Shade+Grade Wise Packing Summary" Then
       ReportName = App.PATH & "\Reports\Machine+Item+Lot+Shade+Grade WisePkgBoxSummary.rpt"
       RPTN = "Machine+Denier+Lot+Shade+Grade Wise Packing Summary Report"
    End If
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    

    CRPT.ReportFileName = ReportName
    PERIOD = dtFrom & " To " & dtTo
    PACK = TXTPACKINGTYPE.Text
    CRPT.ReplaceSelectionFormula rptsql
    
        With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        .Formulas(3) = "DIVISION='" & txtDVCD & "'"
        .Formulas(4) = "PERIOD='" & PERIOD & "'"
        .Formulas(5) = "RPTN='" & RPTN & "'"
        RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        If ReadConfigMaster("000061", 8, "R") Then
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
 MsgBox ERR.Description
 End Sub

Private Sub dtFrom_GotFocus()
  dtFrom.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub dtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtTo.SetFocus
End Sub

Private Sub dtFrom_LostFocus()
  dtFrom.BackColor = vbWhite
End Sub

Private Sub dtTo_GotFocus()
  dtTo.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub dtTo_LostFocus()
  dtTo.BackColor = vbWhite
End Sub

Private Sub dtFrom_Validate(Cancel As Boolean)
    If Not IsDate(dtFrom) And dtFrom <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtFrom.SetFocus
    End If
End Sub

Private Sub Form_Activate()
     Call ColorComponent(Me)
    If Not txtUNIT = Empty Then Exit Sub
    
    txtUnit_KeyDown vbKeyReturn, 0
    
    If txtUNIT = Empty Then
        cmdpreview.Enabled = False
    End If
    
    If txtUNIT = Empty Then Unload Me: Exit Sub
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtDVCD" And txtDVCD = Empty Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  Me.Left = Me.Left - 900
  dtFrom.Text = Format(FSDT, "dd/MM/yyyy")
  dtTo.Text = Format(FEDT, "dd/MM/yyyy")
  
  txtUNIT = UntNm
  txtUNIT.Tag = UNCD
  
  Call FillReportFormat
End Sub

Private Sub txtDenier_GotFocus()
   txtDENIER.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDenier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        txtDENIER = Empty
        txtDENIER.Tag = Empty
        M_DESC = Empty
        NEW_VISIBLE = False
        txtDENIER = SearchList1("Select TOP 20 Code,Name  From FINITMMST WHERE COMP='" & compPth & _
        "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Denier !!")
        txtDENIER.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDENIER = Empty
        txtDENIER.Tag = Empty
    End If
    If KeyCode = vbKeyReturn Then TXTMACHINE.SetFocus
End Sub

Private Sub txtDenier_LostFocus()
 txtDENIER.BackColor = vbWhite
End Sub

Private Sub txtDVCD_Change()

TXTMACHINE.Text = Empty
'TXTPKGSTATION.Text = Empty
txtDENIER.Text = Empty
txtLTNo.Text = Empty
txtGrade.Text = Empty
'txtPackType.Text = Empty
'txtLOC.Text = Empty
End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "'  AND RECSTAT='A'  AND CODE<>'000001'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    End If
    If KeyCode = vbKeyReturn Then TXTPACKINGTYPE.SetFocus
End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtDVCD_Validate(Cancel As Boolean)
If SetIsShadeReq(txtDVCD.Tag) = "Y" Then
     LBLSUBGRD.Caption = "Shade"
  End If
End Sub

Private Sub txtgrade_GotFocus()
txtGrade.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtgrade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   txtGrade = SearchList1("SELECT DISTINCT CODE,GRAD FROM GRDMST", 0, txtGrade, "SELECT GRADE FROM LIST")
   txtGrade.Tag = Key
ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
   txtGrade = Empty
   txtGrade.Tag = Empty
End If
 If KeyCode = vbKeyReturn Then txtZoom.SetFocus
End Sub

Private Sub txtgrade_LostFocus()
 txtGrade.BackColor = vbWhite
End Sub

Private Sub txtltno_GotFocus()
   txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   txtLTNo = SearchList1("SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
   "' AND DVCD = '" & txtDVCD.Tag & "'", 0, txtLTNo, "SELECT LOTNO FROM LIST")
   txtLTNo.Tag = Key
ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
   txtLTNo = Empty
   txtLTNo.Tag = Empty
End If
  If KeyCode = vbKeyReturn Then txtGrade.SetFocus
End Sub

Private Sub txtltno_LostFocus()
  txtLTNo.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
 TXTMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   TXTMACHINE = SearchList1("SELECT CODE,NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
   "' AND DVCD = '" & txtDVCD.Tag & "'", 0, TXTMACHINE, "SELECT MACHINE FROM LIST")
   TXTMACHINE.Tag = Key
ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
   TXTMACHINE = Empty
   TXTMACHINE.Tag = Empty
End If
If KeyCode = vbKeyReturn Then txtLTNo.SetFocus
End Sub

Private Sub txtMACHINE_LostFocus()
TXTMACHINE.BackColor = vbWhite
End Sub

Private Sub TXTPACKINGTYPE_GotFocus()
  TXTPACKINGTYPE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPACKINGTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTPACKINGTYPE = Empty) Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   TXTPACKINGTYPE = SearchList1("SELECT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & _
   "' AND UNIT='" & txtUNIT.Tag & "' AND VTYP='PPF' AND FYCD='" & FYCD & "'", 0, TXTPACKINGTYPE, "SELECT PACKING TYPE FROM LIST")
   TXTPACKINGTYPE.Tag = Key
End If
If KeyCode = vbKeyReturn Then cmbFormat.SetFocus
End Sub

Private Sub TXTPACKINGTYPE_LostFocus()
TXTPACKINGTYPE.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_GotFocus()
  TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL As String
Me.KeyPreview = False

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
     TXTSUBGRD.Tag = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = 13 And TXTSUBGRD = Empty) Then
   Key = Empty: M_DESC = Empty:  NEW_VISIBLE = False
      
   SQL = "SELECT DISTINCT SUBGRD,NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
         "' AND DVCD='" & txtDVCD.Tag & "' "
         
   If SetIsShadeReq(txtDVCD.Tag) = "N" Then
      If txtGrade = Empty Then Exit Sub
      SQL = SQL & " AND GRAD='" & GetCode("GRDMST", txtGrade, "GRAD", "CODE") & "'"
   End If
         
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
   TXTSUBGRD.Tag = Key
End If

Me.KeyPreview = True
End Sub

Private Sub TXTSUBGRD_LostFocus()
TXTSUBGRD.BackColor = vbWhite
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
If KeyCode = vbKeyReturn Then txtDVCD.SetFocus
End Sub

Private Sub txtUNIT_LostFocus()
   txtUNIT.BackColor = vbWhite
End Sub

Private Sub TXTZOOM_GotFocus()
   txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtZoom_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then cmdpreview.SetFocus
End Sub

Private Sub TXTZOOM_LostFocus()
   txtZoom.BackColor = vbWhite
End Sub

Private Function RIGHTDATA() As Boolean
 RIGHTDATA = True
 
 If txtUNIT = Empty Then
    RIGHTDATA = False
    MsgBox "Please Select Unit !!", vbInformation, "Unit Is Key Field Missing"
    txtUNIT.SetFocus
    Exit Function
 End If
    
 If txtDVCD = Empty Then
    RIGHTDATA = False
    MsgBox "Please Select Division !!", vbInformation, "Division Missing !!"
    txtDVCD.SetFocus
    txtDVCD_KeyDown vbKeyReturn, 0
    Exit Function
 End If
 
 If TXTPACKINGTYPE = Empty Then
    RIGHTDATA = False
    MsgBox "Please Select Packing Type !!", vbInformation, "Packing Type is Missing !!"
    TXTPACKINGTYPE.SetFocus
    TXTPACKINGTYPE_KeyDown vbKeyReturn, 0
    Exit Function
 End If
     
 If cmbFormat.ListIndex = -1 Then
    RIGHTDATA = False
    MsgBox "Please Select Report Format !!", vbInformation, "Report Format Missing !!"
    cmbFormat.SetFocus
    Exit Function
 End If
    
    
 If Not IsDate(dtFrom) Then
    RIGHTDATA = False
    MsgBox "Please enter valid Starting Date !!", vbInformation, "Date Error"
    dtFrom.SetFocus
    Exit Function
 End If
    
 If Not IsDate(dtTo) Then
    RIGHTDATA = False
    MsgBox "Please enter valid Ending Date !!", vbInformation, "Date Error"
    dtTo.SetFocus
    Exit Function
 End If
End Function

Private Sub FillReportFormat()
  cmbFormat.Clear
  cmbFormat.AddItem "Machine+Item+Lot+Grade Wise Packing Summary"
  cmbFormat.AddItem "Machine+Item+Grade Wise Packing Summary"
  cmbFormat.AddItem "Item+Month Wise Packing Summary"
  cmbFormat.AddItem "Item+Grade Wise Packing Summary"
  cmbFormat.AddItem "Machine+Item+Grade Wise Packing Summary {%}"
  cmbFormat.AddItem "Date and M/c Wise Packing Summary"
If cmbFormat.ListCount > 0 Then cmbFormat.ListIndex = 0
End Sub



Private Sub SetFormula(PKGTYP As String)
If PKGTYP = "C" Then
    rptsql = "{BOXREGISTER.COMP}='" & compPth & "' AND {BOXREGISTER.UNIT}='" & txtUNIT.Tag & _
    "' AND {BOXREGISTER.DVCD}='" & txtDVCD.Tag & "'"
    
    
    If cmbFormat.ListIndex <> 6 Then
       rptsql = rptsql & " AND {BOXREGISTER.RECSTAT}<>'D' AND  {BOXREGISTER.VBDT} >= DATE(" & Year(dtFrom) & _
                "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {BOXREGISTER.VBDT} <= DATE(" & Year(dtTo) & _
                "," & Month(dtTo) & "," & Day(dtTo) & ") "
    End If
    
    If Not TXTPACKINGTYPE = Empty Then rptsql = rptsql & " AND {BOXREGISTER.DBCD}='" & TXTPACKINGTYPE.Tag & "'"
    If Not txtDENIER = Empty Then rptsql = rptsql & " AND {BOXREGISTER.ICOD}='" & txtDENIER.Tag & "'"
    If Not txtLTNo = Empty Then rptsql = rptsql & " AND {BOXREGISTER.LOTNO}= '" & txtLTNo & "'"
    If Not txtGrade = Empty Then rptsql = rptsql & " AND {BOXREGISTER.GRAD}= " & txtGrade.Tag & ""
    If Not TXTMACHINE = Empty Then rptsql = rptsql & " AND {BOXREGISTER.MCCD} = '" & TXTMACHINE.Tag & "'"
    
    If Not TXTSUBGRD = Empty Then
       Dim SQL As String
       SQL = "SELECT SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
             "' AND DVCD='" & txtDVCD.Tag & "' AND NAME = '" & TXTSUBGRD & "' "
             
       If Not txtGrade = Empty Then
          SQL = SQL & "AND GRAD='" & GetCode("GRDMST", txtGrade, "GRAD", "CODE") & "' "
       End If
       
       If RS.State = 1 Then RS.Close
       RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
       If Not RS.EOF Then
          rptsql = rptsql & " AND {BOXREGISTER.SUBGRD}= '" & RS!SUBGRD & "' "
       End If
    End If
    
Else
    rptsql = "{PKGMAN.COMP}='" & compPth & "' AND  {PKGMAN.DATE} >= DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {PKGMAN.DATE} <= DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND ({PKGMAN.VTYP}='PPF' OR {PKGMAN.VTYP}='OPN' OR {PKGMAN.VTYP}='WST' OR {PKGMAN.VTYP}='DPF') AND {PKGMAN.UNIT}='" & txtUNIT.Tag & "' AND {PKGMAN.RECSTAT}<>'D' AND {PKGMAN.DVCD}='" & txtDVCD.Tag & "'"
    If Not TXTPACKINGTYPE = Empty Then rptsql = rptsql & " AND {PKGMAN.DBCD}='" & TXTPACKINGTYPE.Tag & "'"
    If Not txtDENIER = Empty Then rptsql = rptsql & " AND {PKGMAN.FINITMCOD}='" & txtDENIER.Tag & "'"
    If Not txtLTNo = Empty Then rptsql = rptsql & " AND {PKGMAN.LOTNO}= '" & txtLTNo & "'"
    If Not txtGrade = Empty Then rptsql = rptsql & " AND {PKGMAN.GRAD}= " & txtGrade.Tag & ""
    If Not TXTMACHINE = Empty Then rptsql = rptsql & " AND {PKGMAN.MCCD} = '" & TXTMACHINE.Tag & "'"
End If
End Sub

Private Sub SetViewForPackingSummary()
Dim QRY As String
On Error GoTo VWERR

QRY = "CREATE VIEW VW_MC_ITM_NTWGTSMRY AS " & _
     "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.DBCD,BOXREGISTER.MCCD," & _
     "BOXREGISTER.ICOD,ISNULL(SUM(BOXREGISTER.NTWGT),0) AS NTWGT FROM BOXREGISTER " & _
     "WHERE BOXREGISTER.RECSTAT<>'D' AND BOXREGISTER.VBDT >='" & Format(dtFrom, "MM/DD/YYYY") & _
     "' AND BOXREGISTER.VBDT <='" & Format(dtTo, "MM/DD/YYYY") & _
     "' GROUP BY BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.DBCD,BOXREGISTER.MCCD,BOXREGISTER.ICOD "

CN.Execute "IF ( OBJECT_ID('VW_MC_ITM_NTWGTSMRY') IS NOT NULL ) DROP VIEW VW_MC_ITM_NTWGTSMRY "
CN.Execute QRY


QRY = "CREATE VIEW VW_MHEAD AS " & _
     "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.DBCD," & _
     "ISNULL(SUM(BOXREGISTER.NTWGT),0) AS NTWGT FROM BOXREGISTER " & _
     "WHERE BOXREGISTER.RECSTAT<>'D' AND BOXREGISTER.VBDT >='" & Format(dtFrom, "MM/DD/YYYY") & _
     "' AND BOXREGISTER.VBDT <='" & Format(dtTo, "MM/DD/YYYY") & _
     "' GROUP BY BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.DBCD " & _
     " HAVING ISNULL(SUM(BOXREGISTER.NTWGT),0) > 0"

CN.Execute "IF ( OBJECT_ID('VW_MHEAD') IS NOT NULL ) DROP VIEW VW_MHEAD "
CN.Execute QRY

QRY = "CREATE VIEW VW_GRD_SUBGRD_NTWGTSMRY AS " & _
     "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.DBCD,BOXREGISTER.GRAD,BOXREGISTER.SUBGRD," & _
     "ISNULL(SUM(BOXREGISTER.NTWGT),0) AS NTWGT,VW_MHEAD.NTWGT AS TTLWGT,(ISNULL(SUM(BOXREGISTER.NTWGT),0) * 100)/VW_MHEAD.NTWGT AS PERC FROM BOXREGISTER " & _
     "INNER JOIN VW_MHEAD ON VW_MHEAD.COMP=BOXREGISTER.COMP AND VW_MHEAD.UNIT=BOXREGISTER.UNIT AND VW_MHEAD.DVCD=BOXREGISTER.DVCD AND VW_MHEAD.DBCD=BOXREGISTER.DBCD  " & _
     "WHERE BOXREGISTER.RECSTAT<>'D' AND BOXREGISTER.VBDT >='" & Format(dtFrom, "MM/DD/YYYY") & _
     "' AND BOXREGISTER.VBDT <='" & Format(dtTo, "MM/DD/YYYY") & _
     "' GROUP BY BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,BOXREGISTER.DBCD,BOXREGISTER.GRAD,BOXREGISTER.SUBGRD,VW_MHEAD.NTWGT "

CN.Execute "IF ( OBJECT_ID('VW_GRD_SUBGRD_NTWGTSMRY') IS NOT NULL ) DROP VIEW VW_GRD_SUBGRD_NTWGTSMRY "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

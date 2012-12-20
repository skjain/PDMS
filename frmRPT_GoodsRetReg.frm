VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_GoodsRetReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finish Goods Return Register"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6480
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   6300
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   4905
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
         TabIndex        =   24
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame framDIVISION 
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   6330
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
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   195
         Width           =   4860
      End
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   4860
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
         TabIndex        =   22
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "&Division "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   660
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   6330
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Top             =   195
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
         Left            =   4320
         TabIndex        =   8
         Top             =   195
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
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   19
         Top             =   195
         Width           =   885
      End
      Begin VB.Label Label5 
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
         TabIndex        =   18
         Top             =   195
         Width           =   885
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
      Height          =   1425
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   6330
      Begin VB.TextBox txtDENIER 
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4860
      End
      Begin VB.TextBox txtLTNo 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   4860
      End
      Begin VB.TextBox txtGrade 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   4860
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label9 
         Caption         =   "&Lot No"
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
         TabIndex        =   15
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label10 
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
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   6330
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         Image           =   "frmRPT_GoodsRetReg.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4560
         TabIndex        =   11
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
         Image           =   "frmRPT_GoodsRetReg.frx":0452
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
         TabIndex        =   12
         Top             =   300
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmRPT_GoodsRetReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_DBCD As String
Dim PERIOD As String
Dim M_MCNO As String
Dim PACK As String
   
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
Dim FLAG As Boolean
Dim PKG As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim M_DVCD As String
    CRPT.Reset
    crptConnect CRPT
    
    If Not RIGHTDATA Then Exit Sub
    
    M_MCNO = Empty
    rptsql = Empty
    
       rptsql = "{GRPACKING.COMP}='" & compPth & "' "
       rptsql = rptsql & " AND {GRPACKING.VBDT} >= DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {GRPACKING.VBDT} <= DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
       rptsql = rptsql & " AND {GRPACKING.UNIT}='" & txtUNIT.Tag & "' AND {GRPACKING.RECSTAT}<>'D' AND {GRPACKING.DVCD}='" & txtDVCD.Tag & "'"
        
'FILTERATION CRITERIA

    If Not txtDENIER = Empty Then rptsql = rptsql & " AND {GRPACKING.ICOD}='" & txtDENIER.Tag & "'"
    If Not txtLTNo = Empty Then rptsql = rptsql & " AND {GRPACKING.LOTNO}= '" & txtLTNo & "'"
    If Not txtGrade = Empty Then rptsql = rptsql & " AND {GRPACKING.GRAD}= " & GetCode("GRDMST", txtGrade, "GRAD", "CODE") & ""
     
    If cmbFormat.ListIndex = 0 Then
        ReportName = App.PATH & "\Reports\DateWiseGRRegister.RPT"
        RPTN = "DATEWISE FINISH GOODS RETURN REGISTER"
    ElseIf cmbFormat.ListIndex = 1 Then
        ReportName = App.PATH & "\Reports\PartyWiseGRRegister.RPT"
        RPTN = "PARTYWISE FINISH GOODS RETURN REGISTER"
    ElseIf cmbFormat.ListIndex = 2 Then
        ReportName = App.PATH & "\Reports\PendingDateWiseGRRegister.RPT"
        RPTN = "DATEWISE PENDING GR FOR PRODUCTION"
    ElseIf cmbFormat.ListIndex = 3 Then
        ReportName = App.PATH & "\Reports\PendingPartyWiseGRRegister.RPT"
        RPTN = "PARTYWISE PENDING GR FOR PRODUCTION"
    ElseIf cmbFormat.ListIndex = 4 Then
        ReportName = App.PATH & "\Reports\PendingItemWiseGRRegister.RPT"
        RPTN = "ITEMWISE PENDING GR FOR PRODUCTION"
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
        .Formulas(3) = "COMPANY='" & compNm & "'"
        .Formulas(4) = "DIVISION='" & txtDVCD & "'"
        .Formulas(5) = "PERIOD='" & PERIOD & "'"
        .Formulas(6) = "RPTN='" & RPTN & "'"
        
        '.Formulas(5) = "SHWDELRECS=" & chkShwDelRec.Value
        '.Formulas(6) = "PAGEBRK=" & chkPGBRK.Value
         RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        
        If ReadConfigMaster("000060", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        
        .WindowShowPrintSetupBtn = True
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
    
End Sub

Private Sub txtDenier_LostFocus()
 txtDENIER.BackColor = vbWhite
End Sub

Private Sub txtDVCD_Change()

txtDENIER.Text = Empty
txtLTNo.Text = Empty
txtGrade.Text = Empty

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
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "'  AND CODE<>'000001' AND RECSTAT='A'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    End If
End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
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
End Sub

Private Sub txtltno_LostFocus()
 txtLTNo.BackColor = vbWhite
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
  cmbFormat.AddItem "Datewise Finish Goods Return Register"
  cmbFormat.AddItem "Partywise Finish Goods Return Register"
  cmbFormat.AddItem "Datewise Pending GR for Production"
  cmbFormat.AddItem "Partywise Pending GR for Production"
  cmbFormat.AddItem "Itemwise Pending GR for Production"
If cmbFormat.ListCount > 0 Then cmbFormat.ListIndex = 0
End Sub


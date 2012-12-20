VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_Form402 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form 402"
   ClientHeight    =   6225
   ClientLeft      =   2370
   ClientTop       =   1335
   ClientWidth     =   6570
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   6375
      Begin VB.ComboBox cmbformat 
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
         ItemData        =   "frmRPT_Form402.frx":0000
         Left            =   1320
         List            =   "frmRPT_Form402.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
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
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Text            =   "100"
         Top             =   390
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2640
         Top             =   1395
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3480
         TabIndex        =   21
         Top             =   120
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
         Image           =   "frmRPT_Form402.frx":002A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4800
         TabIndex        =   22
         Top             =   120
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
         Image           =   "frmRPT_Form402.frx":047C
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Caption         =   "Format"
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
         Left            =   1320
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Zoom %"
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
         TabIndex        =   20
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3000
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   6375
      Begin MSComctlLib.ListView lstBills 
         Height          =   2610
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4604
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inv. No"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2143
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   3597
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Tot. Carton"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Bill Amount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "VATCST"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2310
      Left            =   120
      TabIndex        =   17
      Top             =   75
      Width           =   6375
      Begin VB.ComboBox cmbSaleType 
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
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Width           =   5025
      End
      Begin VB.TextBox txtParty 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   56950785
         CurrentDate     =   39343
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   3360
         TabIndex        =   11
         Top             =   1800
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   56950785
         CurrentDate     =   39343
      End
      Begin WelchButton.lvButtons_H cmdSEARCH 
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Search"
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
         Image           =   "frmRPT_Form402.frx":0A16
         cBack           =   -2147483633
      End
      Begin VB.Label Label5 
         Caption         =   "Party Name"
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
         TabIndex        =   6
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Date From :"
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
         Left            =   135
         TabIndex        =   8
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "To :"
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
         Left            =   2880
         TabIndex        =   10
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label Label14 
         Caption         =   "Sale Type :"
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
         Left            =   135
         TabIndex        =   4
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Division :"
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
         Left            =   135
         TabIndex        =   2
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmRPT_Form402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_DBCD As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview

    Dim M_VBNO As String
    Dim ctr As Integer
    Dim VBNO As String
    Dim BNET As Double
    Dim TQTY As Double
    Dim TPCS As Integer
    Dim VatCst As Double
    Dim LAST_VBNO As String
    
    CRPT.Reset
    crptConnect CRPT
    
    If lstBills.ListItems.COUNT < 1 Then
        MsgBox "No Item Found In List !!", vbInformation
        Exit Sub
    End If
    
    For ctr = 1 To lstBills.ListItems.COUNT
        If lstBills.ListItems(ctr).Checked = True Then
            If M_VBNO <> Empty Then M_VBNO = M_VBNO & ","
            M_VBNO = M_VBNO & "'" & lstBills.ListItems(ctr).Text & "'"
            VBNO = VBNO & Left(lstBills.ListItems(ctr).Text, 6) & ", "
            LAST_VBNO = "'" & lstBills.ListItems(ctr).Text & "'"
            BNET = BNET + Val(lstBills.ListItems.Item(ctr).SubItems(5))
            TQTY = TQTY + Val(lstBills.ListItems.Item(ctr).SubItems(3))
            TPCS = TPCS + Val(lstBills.ListItems.Item(ctr).SubItems(4))
            VatCst = VatCst + Val(lstBills.ListItems.Item(ctr).SubItems(6))
        End If
    Next
    
    M_DBCD = GetDBCDPDMS("CODE", "NAME", cmbSaleType, txtUNIT.Tag, "SAL")
    
    VBNO = Left(VBNO, Len(VBNO) - 2)
    
    If M_VBNO = Empty Then
        MsgBox "No Item Selected !!", vbInformation, "Select Invoice From List !!"
        lstBills.SetFocus
        Exit Sub
    End If
    
    If cmbFormat.ListIndex = 0 Then
        ReportName = App.PATH & "\Reports\RPT_FORM402.rpt"
    Else
        ReportName = App.PATH & "\Reports\RPT_FORM402WIN.rpt"
    End If
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
    
    rptsql = "{BILLMAIN.COMP}='" & compPth & "' AND {BILLMAIN.VTYP}='SAL' And {BILLMAIN.RECSTAT}<>'D' AND {BILLMAIN.UNIT}='" & txtUNIT.Tag & _
    "' AND {BILLMAIN.DVCD}='" & txtDVCD.Tag & "' AND {BILLMAIN.DBCD}='" & M_DBCD & _
    "' AND {BILLMAIN.VBNO} IN [" & LAST_VBNO & "]"
    
    CRPT.ReplaceSelectionFormula rptsql
    RPTN = "FORM 402 - "
    
    With CRPT
        .Formulas(1) = "BILLNO='" & VBNO & "'"
        .Formulas(2) = "BILLAMOUNT=" & BNET
        .Formulas(3) = "TOTALPCS=" & TPCS
        .Formulas(4) = "TOTALQTY=" & TQTY
        .Formulas(5) = "VATCST=" & VatCst
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
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdSearch_Click()
Dim Item As ListItem
    
    If txtUNIT = Empty Then
        txtUNIT.SetFocus
        Exit Sub
    End If
    
    If txtDVCD = Empty Then
        txtDVCD.SetFocus
        Exit Sub
    End If
    
    If cmbSaleType = Empty Then
        cmbSaleType.SetFocus
        Exit Sub
    End If
    
    M_DBCD = GetDBCDPDMS("CODE", "NAME", cmbSaleType, txtUNIT.Tag, "SAL")
        
    SQL = "SELECT * FROM BILLMAIN INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD WHERE VTYP='SAL' AND RECSTAT<>'D' " & _
          "AND DATE>='" & Format(dtFrom, "MM/dd/yyyy") & "' AND DATE<='" & Format(dtTo, "MM/dd/yyyy") & _
        "' AND COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "' And DVCD='" & txtDVCD.Tag & _
        "' AND DBCD='" & M_DBCD & "' "
        
    If txtParty <> Empty Then
       SQL = SQL & " AND BILLMAIN.PCOD ='" & txtParty.Tag & "' "
    End If
    
    Set rsTemp = New Recordset
    
    rsTemp.Open SQL, CN
    lstBills.ListItems.Clear
    Do While Not rsTemp.EOF
    
        Set Item = lstBills.ListItems.ADD
        Item.Text = Trim(rsTemp!VBNO)
        Item.SubItems(1) = rsTemp!Date
        Item.SubItems(2) = rsTemp!NAME
        Item.SubItems(3) = Format(rsTemp!TQTY, "#.000")
        Item.SubItems(4) = rsTemp!TPCS
        Item.SubItems(5) = Format(rsTemp!BNET, "#.00")
        Item.SubItems(6) = Format(rsTemp!VAT + rsTemp!CST, "#.00")
        
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    
    If lstBills.ListItems.COUNT > 0 Then lstBills.SetFocus: cmdpreview.Default = True Else cmdpreview.Default = False
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  cmbSaleType.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtDVCD = Empty And ActiveControl.NAME = "txtDVCD" And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtDBCD" And txtDBCD = Empty And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtParty" And txtParty = Empty And KeyCode = 13 Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
 Call ColorComponent(Me)
 Call CenterChild(frm_Main, Me)
    
    dtFrom = GetMinDate
    dtTo = GetMaxDate
    Me.KeyPreview = True
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    
    Call SetSaleType
End Sub

Private Sub lstBills_GotFocus()
lstBills.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstBills_LostFocus()
lstBills.BackColor = vbWhite
End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtParty_GotFocus()
 txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtParty = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtParty = SearchList1("SELECT TOP 20 Code,NAME From ACCMST", 0, Empty, "Select Party")
        txtParty.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtParty = Empty
        txtParty.Tag = Empty
    End If
End Sub

Private Sub txtParty_LostFocus()
 txtParty.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(txtUNIT) = Empty Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("SELECT TOP 20 Code,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("SELECT TOP 20 Code,NAME From DIVMST Where COMP='" & compPth & "' and UNIT='" & txtUNIT.Tag & "'  AND CODE<>'000001' AND RECSTAT='A'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
End Sub

Private Sub txtUNIT_LostFocus()
txtUNIT.BackColor = vbWhite
End Sub

Private Sub SetSaleType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbSaleType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop

If cmbSaleType.ListCount > 1 Then cmbSaleType.ListIndex = 0

End Sub




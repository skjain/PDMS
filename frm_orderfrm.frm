VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frm_orderfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Form"
   ClientHeight    =   6795
   ClientLeft      =   3780
   ClientTop       =   2535
   ClientWidth     =   8340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8340
   Begin VB.Frame Frame4 
      Height          =   4680
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   8145
      Begin MSComctlLib.ListView lstInvoice 
         Height          =   4335
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   7646
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order. No"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   4588
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agent  Name"
            Object.Width           =   4588
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   8130
      Begin VB.TextBox txtPCOD 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   6615
      End
      Begin VB.TextBox txtSM 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1320
         TabIndex        =   5
         Top             =   960
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
         Left            =   3600
         TabIndex        =   7
         Top             =   960
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
      Begin WelchButton.lvButtons_H cmdSEARCH 
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   960
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
         Image           =   "frm_orderfrm.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
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
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Salesman Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "&From Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "&To Date : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   6
         Top             =   960
         Width           =   765
      End
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   120
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin WelchButton.lvButtons_H cmdprv 
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   6240
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
      Image           =   "frm_orderfrm.frx":039A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdclose 
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   6240
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
      Image           =   "frm_orderfrm.frx":07EC
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frm_orderfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub cmdSearch_Click()
 Call GenInvList
End Sub

Private Sub dtFrom_Validate(Cancel As Boolean)
    If Not IsDate(dtFrom) And dtFrom <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtFrom.SetFocus
    End If
End Sub

Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cmdSearch.SetFocus
End If
End Sub

Private Sub dtTo_Validate(Cancel As Boolean)
    If Not IsDate(dtTo) And dtTo <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtTo.SetFocus
    End If
End Sub

Public Sub lstInvoice_GotFocus()
    lstInvoice.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstInvoice_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  If Item.Checked = True Then
    SendKeys "{DOWN}"
  End If
End Sub

Private Sub lstInvoice_LostFocus()
lstInvoice.BackColor = vbWhite
End Sub

Private Sub CMDPRV_Click()
  Dim M_ORDN As String
  Dim I As Long
  Dim LED_BALN As Double
  Dim ACC_COD As String
  Dim ORD_DAT As Date
  LED_BALN = 0
  If RS.State = 1 Then RS.Close
  
  Dim NWDVCD As String
  Dim NWDVNM As String
  Dim ORDN As String
  
  For I = 1 To lstInvoice.ListItems.COUNT
   If lstInvoice.ListItems(I).Checked = True Then
      If M_ORDN <> Empty Then M_ORDN = M_ORDN & ","
      M_ORDN = M_ORDN & "'" & Trim(lstInvoice.ListItems(I)) & "'"
   End If
  Next
        
  If M_ORDN = Empty Then
     MsgBox "No Item Selected !!", vbInformation, "No Information Found !!"
     If lstInvoice.ListItems.COUNT < 1 Then Call lstInvoice_GotFocus
     Exit Sub
  End If
  
  CRPT.Reset
  crptConnect CRPT
  ReportName = App.PATH & "\Reports\RPT_SALEORDER_APR.RPT"
  If Dir(ReportName, vbNormal) = Empty Then
     ReportErrorMessage 1001
     Exit Sub
  End If
  
  CRPT.DiscardSavedData = True
  CRPT.WindowShowRefreshBtn = True
  CRPT.WindowShowPrintBtn = True
  CRPT.ReportFileName = ReportName
  CRPT.ReplaceSelectionFormula "{ORDMAN.COMP}='" & compPth & "' AND TRIM({ORDMAN.ORDN}) IN [" & M_ORDN & "] "
  CRPT.WindowState = crptMaximized
  CRPT.WindowShowPrintBtn = True
  CRPT.WindowShowPrintSetupBtn = True
  CRPT.WindowShowSearchBtn = True
  CRPT.WindowShowExportBtn = True
  CRPT.WindowShowRefreshBtn = False
  CRPT.WindowTitle = "Order Report " & Space(5) & "Report : " & ReportName
  CRPT.ACTION = 1
End Sub

Private Sub Form_Activate()
    If lstInvoice.ListItems.COUNT >= 1 Then Call lstInvoice.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  dtFrom = Format(Now, "DD/MM/YYYY")
  dtTo = Format(Now, "DD/MM/YYYY")
End Sub

Public Sub GenInvList()
Dim SQL As String
Dim Item As ListItem
Dim ORDDBCD As String

If Not IsDate(dtFrom) Then
     dtFrom.SetFocus
     Exit Sub
ElseIf Not IsDate(dtTo) Then
     dtTo.SetFocus
     Exit Sub
End If

If RS.State = 1 Then RS.Close
RS.Open "select code from salmanmst where name='" & txtSM.Text & "'"
If Not RS.EOF Then
  ORDDBCD = RS!CODE
Else
  Exit Sub
End If

SQL = Empty
SQL = "SELECT ORDMAN.ORDN AS ORDNO,ORDMAN.ORDT,ACCMST.NAME AS PARTY,REFMST.NAME AS AGENT FROM ORDMAN "
SQL = SQL & "INNER JOIN ACCMST ON ORDMAN.PCOD=ACCMST.CODE "
SQL = SQL & "INNER JOIN REFMST ON ORDMAN.BRCD=REFMST.CODE "
SQL = SQL & "WHERE ORDMAN.ORDT>='" & Format(dtFrom, "MM/dd/yyyy") & "' and ORDMAN.ORDT<= '" & Format(dtTo, "MM/dd/yyyy") & "' AND DBCD='" & ORDDBCD & "' "
    
    If txtPCOD <> Empty Then
       SQL = SQL & " AND ORDMAN.PCOD = '" & txtPCOD.Tag & "' "
    End If
    
    SQL = SQL & " GROUP BY ORDMAN.ORDN,ORDMAN.ORDT,ACCMST.NAME,REFMST.NAME "
    LSQL = SQL

    Set rsTemp = New Recordset
    rsTemp.Open SQL, CN
    
    lstInvoice.ListItems.Clear
    
    Do While Not rsTemp.EOF
        Set Item = lstInvoice.ListItems.ADD
        Item.Text = rsTemp!ORDNO & ""
        Item.SubItems(1) = rsTemp!ORDT
        Item.SubItems(2) = rsTemp!PARTY
        Item.SubItems(3) = rsTemp!AGENT
        rsTemp.MoveNext
    Loop
    rsTemp.Close
End Sub

Private Sub txtPCOD_GotFocus()
 txtPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtPCOD = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select Party From List")
        txtPCOD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtPCOD = Empty
        txtPCOD.Tag = Empty
    End If
End Sub

Private Sub txtPCOD_LostFocus()
 txtPCOD.BackColor = vbWhite
End Sub


Private Sub txtSM_GotFocus()
txtSM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtSM_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtSM = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtSM = SearchList1("Select  TOP 20 Code,Name From SALMANMST", 0, Empty, "Select Sales Man From List")
        txtSM.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtSM = Empty
        txtSM.Tag = Empty
    End If

End Sub

Private Sub txtSM_LostFocus()
  txtSM.BackColor = vbWhite
End Sub

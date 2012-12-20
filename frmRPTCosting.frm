VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPTCosting 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costing Report"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6615
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   6450
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1320
         TabIndex        =   5
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
         TabIndex        =   7
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
         TabIndex        =   4
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
         TabIndex        =   6
         Top             =   195
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      Height          =   795
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   6405
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         Image           =   "frmRPTCosting.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4560
         TabIndex        =   12
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
         Image           =   "frmRPTCosting.frx":0452
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
         TabIndex        =   9
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.Frame framDIVISION 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6450
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   195
         Width           =   4860
      End
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   3
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
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
         Left            =   240
         TabIndex        =   2
         Top             =   660
         Width           =   885
      End
   End
   Begin MSComctlLib.ListView lstCostHead 
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   7011
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Costing Head"
         Object.Width           =   10584
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmRPTCosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ROWNO As Long
Dim SWITCH As Boolean

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
   
  Dim CHCOD As String, SQL As String
   
  If txtUNIT = Empty Then
     MsgBox "Please Select Unit", vbInformation
     txtUNIT.SetFocus
     Exit Sub
  End If
  
  If txtDVCD = Empty Then
     MsgBox "Please Select Division", vbInformation
     txtDVCD.SetFocus
     Exit Sub
  End If
    
  If Not IsDate(dtFrom) Then
     MsgBox "Pleas Select Correct Starting Date", vbInformation
     dtFrom.SetFocus
     Exit Sub
  End If
    
  If Not IsDate(dtTo) Then
     MsgBox "Pleas Select Correct Ending Date", vbInformation
     dtTo.SetFocus
     Exit Sub
  End If
  
  Dim I As Long
   For I = 1 To lstCostHead.ListItems.COUNT
      If lstCostHead.ListItems(I).Checked = True Then
         If CHCOD <> Empty Then CHCOD = CHCOD & ","
         CHCOD = CHCOD & "'" & Trim(lstCostHead.ListItems(I).SubItems(1)) & "'"
      End If
   Next

  Dim TMPRS As ADODB.Recordset
  Set TMPRS = New ADODB.Recordset
  
  Dim INRS As ADODB.Recordset
  Set INRS = New ADODB.Recordset
  
  Dim PRD_QNTY As Double, CONQNTY As Double, ISS_QNTY As Double, ISS_VAL As Double, AVG_RATE As Double
  Dim RAW_VALUE As Double, CHVAL As Double
    
  'PRODUCTION QNTY
  If TMPRS.State = 1 Then TMPRS.Close
  TMPRS.Open "SELECT ISNULL(SUM(NTWGT),0) AS FINQTY FROM BOXREGISTER WHERE COMP='" & compPth & _
             "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & _
             "' AND VBDT >='" & Format(dtFrom, "MM/DD/YYYY") & _
             "' AND VBDT <='" & Format(dtTo, "MM/DD/YYYY") & _
             "' AND RECSTAT<>'D' ", CN, adOpenDynamic, adLockOptimistic
  If Not TMPRS.EOF Then
     PRD_QNTY = Round(Val(TMPRS!FINQTY & ""), 3)
  End If
  
  SQL = "CREATE VIEW VW_COSTING AS " & _
         "SELECT 1 AS SRNO,'TOTAL PRODUCTION QUANTITY' AS NAME," & PRD_QNTY & " AS VALUE FROM STORETRAN "
  
  'CONSUMPTION QNTY
  If TMPRS.State = 1 Then TMPRS.Close
  TMPRS.Open "SELECT ISNULL(SUM(QNTY),0) AS CONQTY FROM STORETRAN WHERE COMP='" & compPth & _
             "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & _
             "' AND DATE >='" & Format(dtFrom, "MM/DD/YYYY") & _
             "' AND DATE <='" & Format(dtTo, "MM/DD/YYYY") & _
             "' AND RECSTAT<>'D' AND VTYP='PPF' AND OPER='-' ", CN, adOpenDynamic, adLockOptimistic
  If Not TMPRS.EOF Then
     CONQNTY = Round(Val(TMPRS!CONQTY & ""), 3)
  End If
  
  SQL = SQL & "UNION " & _
         "SELECT 2 AS SRNO,'RAW CONSUMPTION QUANTITY' AS NAME," & CONQNTY & " AS VALUE FROM STORETRAN "
    
  'ISSUED QNTY FOR AVERAGE RATE
  If TMPRS.State = 1 Then TMPRS.Close
  TMPRS.Open "SELECT ISNULL(SUM(QNTY),0) AS ISSQTY FROM STORETRAN WHERE COMP='" & compPth & _
             "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & _
             "' AND DATE >='" & Format(dtFrom, "MM/DD/YYYY") & _
             "' AND DATE <='" & Format(dtTo, "MM/DD/YYYY") & _
             "' AND RECSTAT<>'D' AND OPER='+' AND VTYP='ISS' AND " & _
             "ICOD IN (SELECT DISTINCT TXULOT.RICD FROM BOXREGISTER " & _
             " INNER JOIN TXULOT ON TXULOT.COMP=BOXREGISTER.COMP AND TXULOT.UNIT=BOXREGISTER.UNIT " & _
             " AND TXULOT.DVCD=BOXREGISTER.DVCD AND TXULOT.LTNO=BOXREGISTER.LOTNO " & _
             " WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & txtUNIT.Tag & "' AND BOXREGISTER.DVCD='" & txtDVCD.Tag & _
             "' AND VBDT >='" & Format(dtFrom, "MM/DD/YYYY") & _
             "' AND VBDT <='" & Format(dtTo, "MM/DD/YYYY") & _
             "' AND BOXREGISTER.RECSTAT<>'D')", CN, adOpenDynamic, adLockOptimistic
             
  If Not TMPRS.EOF Then
     ISS_QNTY = Val(TMPRS!ISSQTY & "")
  End If
  
  'ISSUED VALUE FOR AVERAGE RATE
  If TMPRS.State = 1 Then TMPRS.Close
  TMPRS.Open "SELECT ISNULL(SUM(QNTY * RATE),0) AS ISSVALUE FROM STORETRAN WHERE COMP='" & compPth & _
             "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & _
             "' AND DATE >='" & Format(dtFrom, "MM/DD/YYYY") & _
             "' AND DATE <='" & Format(dtTo, "MM/DD/YYYY") & _
             "' AND RECSTAT<>'D' AND OPER='+' AND VTYP='ISS' AND " & _
             "ICOD IN (SELECT DISTINCT TXULOT.RICD FROM BOXREGISTER " & _
             " INNER JOIN TXULOT ON TXULOT.COMP=BOXREGISTER.COMP AND TXULOT.UNIT=BOXREGISTER.UNIT " & _
             " AND TXULOT.DVCD=BOXREGISTER.DVCD AND TXULOT.LTNO=BOXREGISTER.LOTNO " & _
             " WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & txtUNIT.Tag & "' AND BOXREGISTER.DVCD='" & txtDVCD.Tag & _
             "' AND VBDT >='" & Format(dtFrom, "MM/DD/YYYY") & _
             "' AND VBDT <='" & Format(dtTo, "MM/DD/YYYY") & _
             "' AND BOXREGISTER.RECSTAT<>'D')", CN, adOpenDynamic, adLockOptimistic
             
  If Not TMPRS.EOF Then
     ISS_VAL = Val(TMPRS!ISSVALUE & "")
  End If
  
  'AVG. CONSUMPTION RATE
  If ISS_VAL > 0 Or ISS_QNTY > 0 Then
        AVG_RATE = Round(ISS_VAL / ISS_QNTY, 2)
  Else
        AVG_RATE = 0
  End If
  
   SQL = SQL & "UNION " & _
         "SELECT 3 AS SRNO,'AVERAGE RATE OF RAW MATERIAL' AS NAME," & AVG_RATE & " AS VALUE FROM STORETRAN "
   
  'RAW VALUE
   RAW_VALUE = Round(AVG_RATE * CONQNTY, 2)
   SQL = SQL & "UNION " & _
         "SELECT 4 AS SRNO,'TOTAL RAW MATERIAL VALUE' AS NAME," & RAW_VALUE & " AS VALUE FROM STORETRAN "

   
  'COST HEAD WISE COSTING :
   SQL = SQL & "UNION " & _
         "SELECT 5 AS SRNO,REFMST.NAME AS NAME,ISNULL(SUM(AMNT),0) AS VALUE FROM STORETRAN " & _
         "INNER JOIN REFMST ON REFMST.CODE = STORETRAN.CSHD AND REFMST.CATA='N' " & _
         "WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & _
         "' AND DATE >='" & Format(dtFrom, "MM/DD/YYYY") & "' AND DATE <='" & Format(dtTo, "MM/DD/YYYY") & _
         "' AND RECSTAT<>'D' AND OPER='+' AND VTYP='ISS' AND CSHD IN (" & CHCOD & " ) GROUP BY REFMST.NAME "
    
   CN.Execute "IF ( OBJECT_ID('VW_COSTING') IS NOT NULL ) DROP VIEW VW_COSTING "
   CN.Execute SQL
      
   CRPT.Reset
   crptConnect CRPT
   ReportName = Empty
   ReportName = App.PATH & "\Reports\RPTCosting.rpt"
   Debug.Print ReportName
    
   If Dir(ReportName, vbNormal) = Empty Then
      ReportErrorMessage 1001
      Exit Sub
   End If
    
   CRPT.ReportFileName = ReportName
   PERIOD = dtFrom & " To " & dtTo
     
    With CRPT
       .Formulas(1) = "COMPANY='" & compNm & "'"
       .Formulas(2) = "UNIT='" & txtUNIT & "'"
       .Formulas(3) = "DIVISION='" & txtDVCD & "'"
       .Formulas(4) = "PERIOD='" & PERIOD & "'"
       .Formulas(5) = "TOT_FIN_QTY=" & PRD_QNTY & ""
       
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
        If txtUNIT.Enabled Then txtUNIT.SetFocus
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
End Sub

Private Sub dtFrom_LostFocus()
  dtFrom.BackColor = vbWhite
End Sub

Private Sub dtTo_GotFocus()
  dtTo.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub dtTo_LostFocus()
  dtTo.BackColor = vbWhite
End Sub

Private Sub Form_Load()
 Call ColorComponent(Me)
 Call CenterChild(frm_Main, Me)
    
 dtFrom.Text = Format(FSDT, "dd/MM/yyyy")
 dtTo.Text = Format(FEDT, "dd/MM/yyyy")
 txtUNIT = UntNm
 txtUNIT.Tag = UNCD
 Call SetCostHead
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lstCostHead_GotFocus()
  lstCostHead.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstCostHead_LostFocus()
  lstCostHead.BackColor = vbWhite
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


Private Sub SetCostHead()

lstCostHead.ListItems.Clear

Dim SQL As String, ctr As Long
Dim Item As ListItem

SQL = "Select * From REFMST WHERE CATA='N' ORDER BY NAME"

If RS.State = 1 Then RS.Close
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Do While Not RS.EOF
   Set Item = lstCostHead.ListItems.ADD
   Item.Text = Trim(RS!NAME & "")
   Item.SubItems(1) = Trim(RS!CODE & "")
   RS.MoveNext
Loop
RS.Close

End Sub

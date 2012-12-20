VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RPT_EXCISEREG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excise Register"
   ClientHeight    =   2850
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5670
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox Regtype 
         Height          =   315
         ItemData        =   "RPT_EXCISEREG.frx":0000
         Left            =   2040
         List            =   "RPT_EXCISEREG.frx":0010
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   3975
      End
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3420
         Visible         =   0   'False
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   3255
         TabIndex        =   8
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   18284545
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1245
         TabIndex        =   6
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   18284545
         CurrentDate     =   38429
      End
      Begin VB.Label Label3 
         Caption         =   "Type of Register"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label Label8 
         Caption         =   "Unit Name :"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   870
      End
      Begin VB.Label Label14 
         Caption         =   "Division :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   3465
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   825
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2775
         TabIndex        =   7
         Top             =   1065
         Width           =   330
      End
   End
   Begin VB.Frame Frame4 
      Height          =   810
      Left            =   120
      TabIndex        =   11
      Top             =   1950
      Width           =   5430
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1335
         TabIndex        =   14
         Text            =   "100"
         Top             =   1905
         Width           =   480
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   4200
         TabIndex        =   13
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&View"
         Height          =   405
         Left            =   3000
         TabIndex        =   12
         Top             =   225
         Width           =   1080
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2280
         Top             =   1155
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label12 
         Caption         =   "Report &Zoom %"
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   1905
         Width           =   1140
      End
   End
End
Attribute VB_Name = "RPT_EXCISEREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
    If txtUNIT.Tag = "" Then Exit Sub
    Call create_table
    Call gen_RGREG
    CRPT.Reset
    
    crptConnect CRPT
    
    ReportName = App.PATH & "\Reports\RPT_MODVATELEDGER.RPT"
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    
    SQL = Empty

    SQL = "{PLATEMP.COMP}='" & compPth & "' AND {PLATEMP.CENVAT}<>0 AND {PLATEMP.TDAT}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ")  AND {PLATEMP.TDAT}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {PLATEMP.UNIT} IN [" & sel_untcod & "] AND {PLATEMP.EXTRA1}='" & ComputerName & "'"
    
    If txtDVCD <> Empty Then SQL = SQL & " AND {PLATEMP.DVCD}='" & txtDVCD.Tag & "'"
    

    CRPT.ReportFileName = ReportName
    
    CRPT.ReplaceSelectionFormula SQL
    RPTN = Regtype.Text
    
    
    
    
    PERIOD = dtFrom & " To " & dtTo
    
    If txtDVCD = Empty Then
        M_DVCD = "N/A"
    Else
        M_DVCD = txtDVCD
    End If
    Dim opncenvat As Double
    Dim OPNEDUCESS As Double
    Dim opnhredcess As Double
    Dim OPNAED As Double
    opncenvat = 0
    OPNEDUCESS = 0
    opnhredcess = 0
    OPNAED = 0
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT ISNULL(SUM(ADUTY),0) AS ADUTY,ISNULL(SUM(CENVAT),0) AS CENVAT, ISNULL(SUM(EDUCESS),0) AS EDUCESS,ISNULL(SUM(HEDCESS),0) AS HEDCESS,ISNULL(SUM(ADUTY),0) AS ADUTY FROM PLATEMP WHERE COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ") AND TDAT<'" & Format(dtFrom, "MM/DD/YYYY") & "' AND VTYP<>'EXD'"
    If Not RS.EOF Then
      'Chage as per required by sumeet
      opncenvat = RS!CENVAT + RS!ADUTY
      OPNEDUCESS = RS!EDUCESS
      opnhredcess = RS!HEDCESS
      OPNAED = RS!ADUTY
    End If
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT ISNULL(SUM(ADUTY),0) AS ADUTY,ISNULL(SUM(CENVAT),0) AS CENVAT, ISNULL(SUM(EDUCESS),0) AS EDUCESS,ISNULL(SUM(HEDCESS),0) AS HEDCESS,ISNULL(SUM(ADUTY),0) AS ADUTY FROM PLATEMP WHERE COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ") AND TDAT<'" & Format(dtFrom, "MM/DD/YYYY") & "' AND VTYP='EXD'"
    If Not RS.EOF Then
      'Chage as per required by sumeet
      opncenvat = opncenvat - RS!CENVAT - RS!ADUTY
      OPNEDUCESS = OPNEDUCESS - RS!EDUCESS
      opnhredcess = opnhredcess - RS!HEDCESS
      OPNAED = OPNAED - RS!ADUTY
    End If
    
    
    With CRPT

        .Formulas(1) = "rptn='" & RPTN & "'"
        .Formulas(2) = "PERIOD='" & PERIOD & "'"
        .Formulas(3) = "OPNCENVAT=" & opncenvat & ""
        .Formulas(4) = "OPNEDUCESS=" & OPNEDUCESS & ""
        .Formulas(5) = "OPNHREDCESS=" & opnhredcess & ""
        .Formulas(6) = "OPNAED=" & OPNAED & ""
        .ReportTitle = RPTN
        RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000078", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = 13 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
        
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtFrom = GetMinDate
    dtTo = GetMaxDate
    Regtype.ListIndex = 0
End Sub

Private Sub Regtype_GotFocus()
 Regtype.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub Regtype_LostFocus()
 Regtype.BackColor = vbWhite
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        If txtUNIT = Empty Then Exit Sub
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("SELECT TOP 20 Code,NAME From DIVMST Where COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ")", 0, Empty, "Select UNIT")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
    End If

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

Private Sub gen_RGREG()
  On Error Resume Next
  CN.Execute "IF NOT EXISTS(SELECT * FROM SYSOBJECTS WHERE NAME='PLA_TEMP') Select * into PLATEMP from PLATRN where 1=2"
  CN.Execute "ALTER TABLE PLATEMP ADD ICOD CHAR(10) NULL"
  CN.Execute "ALTER TABLE PLATEMP ADD CHAP CHAR(15) NULL"
  CN.Execute "ALTER TABLE PLATEMP ADD VTYP CHAR(3) NULL"
  CN.Execute "ALTER TABLE PLATEMP ADD VBNO CHAR(20) NULL"
  CN.Execute "ALTER TABLE PLATEMP ADD PCOD CHAR(6) NULL"
  CN.Execute "ALTER TABLE PLATEMP ADD ASSV NUMERIC(18,3) NOT NULL DEFAULT (0)"
  CN.Execute "ALTER TABLE PLATEMP ADD QNTY NUMERIC(18,3) NOT NULL DEFAULT (0)"
  CN.Execute "ALTER TABLE PLATEMP ADD CESS NUMERIC(18,3) NOT NULL DEFAULT (0)"
  CN.Execute "ALTER TABLE PLATEMP ALTER COLUMN VBNO CHAR(20) NULL"
  CN.Execute "ALTER TABLE PLATEMP ADD VBDT [datetime] NULL"
  CN.Execute "ALTER TABLE PLATMEP DROP COLUMN ROWGUID"
  
  CN.Execute "DELETE FROM PLATEMP"
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  Dim CNTR As Double
  CNTR = 1
  On Error GoTo LAST
  
  Set MSTDAT = New ADODB.Recordset
  Dim PLAREG As New ADODB.Recordset
  Set PLAREG = New ADODB.Recordset
  Select Case Regtype.Text
   Case "RG23-A-II Register"
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM ER1DATA WHERE RCOD='RG23-A' AND VTYP='OPN' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ")", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = "RG23-A"
     RS!TDAT = PLAREG!TDAT
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!opncenvat
     RS!CESS = PLAREG!OPNCESS
     RS!ADUTY = PLAREG!OPNADUTY
     RS!NCCD = PLAREG!OPNNCCD
     RS!EDUCESS = PLAREG!OPNEDUCESS
     RS!HEDCESS = PLAREG!opnhredcess
     RS!VTYP = "OPN"
     RS!ICOD = "YYYYYYYYYY"
     RS!CHAP = ""
     RS!ASSV = 0
     RS!EXTRA1 = ComputerName
     RS!PCOD = "OPNBAL"
     
     
     RS.Update
     PLAREG.MoveNext
     CNTR = CNTR + 1
    Loop
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM EGPMAN WHERE TTYP='RG23-A' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ") AND ((CENVAT+EDUCESS+H_ED_CESS+A_DUTY)<>0 OR (SRVTAX)<>0) AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     Dim M_CHAP As String
     M_CHAP = ""
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM ITMMST WHERE CODE='" & PLAREG!ICOD & "'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       M_CHAP = MSTDAT!igcd
       If MSTDAT.State = 1 Then MSTDAT.Close
       MSTDAT.Open "SELECT * FROM IGMMST WHERE CODE='" & M_CHAP & "'", CN, adOpenDynamic, adLockOptimistic
       If Not MSTDAT.EOF Then
         M_CHAP = MSTDAT!CHAP & ""
       End If
     End If
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = PLAREG!TTYP
     RS!TDAT = PLAREG!Date
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!CENVAT
     RS!CESS = PLAREG!CESS
     RS!ADUTY = PLAREG!A_DUTY
     RS!NCCD = PLAREG!NCCD
     RS!EDUCESS = PLAREG!EDUCESS
     RS!HEDCESS = PLAREG!H_ED_CESS
     RS!VTYP = PLAREG!VTYP
     RS!ICOD = PLAREG!ICOD
     RS!CHAP = M_CHAP
     RS!ASSV = PLAREG!ITOT
     RS!PCOD = PLAREG!DRAC
     RS!EXTRA1 = ComputerName
     RS!EXTRA3 = PLAREG!SRNO & ""
     If PLAREG!DRAC = "ER1DAT" And PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = "ER1 For the Month"
     ElseIf PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = PLAREG!EXTRA5
     End If
     
     
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT VBNO,PODT FROM PURMAN WHERE COMP='" & PLAREG!COMP & "' AND VTYP='" & PLAREG!VTYP & "' AND SRNO='" & PLAREG!SRNO & "'", CN, adOpenDynamic, adLockOptimistic
     If MSTDAT.EOF = False Then
       RS!VBNO = MSTDAT!VBNO
       If IsNull(MSTDAT!PODT) = True Then
        RS!VBDT = PLAREG!Date
       Else
        RS!VBDT = MSTDAT!PODT
       End If
      Else
       RS!VBNO = Trim(PLAREG!chln)
       RS!VBDT = PLAREG!Date
     End If
     RS!VBDT = PLAREG!CHDT
     RS!QNTY = PLAREG!QNTY
     RS.Update
     CNTR = CNTR + 1
     PLAREG.MoveNext
    Loop
   Case "RG23-C-II Register"
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM ER1DATA WHERE RCOD='RG23-C' AND VTYP='OPN' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ")", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = "RG23-C"
     RS!TDAT = PLAREG!TDAT
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!opncenvat
     RS!CESS = PLAREG!OPNCESS
     RS!ADUTY = PLAREG!OPNADUTY
     RS!NCCD = PLAREG!OPNNCCD
     RS!EDUCESS = PLAREG!OPNEDUCESS
     RS!HEDCESS = PLAREG!opnhredcess
     RS!VTYP = "OPN"
     RS!ICOD = "YYYYYYYYYY"
     RS!CHAP = ""
     RS!ASSV = 0
     RS!PCOD = "OPNBAL"
     RS!EXTRA1 = ComputerName
     RS.Update
     PLAREG.MoveNext
     CNTR = CNTR + 1
    Loop
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM EGPMAN WHERE TTYP='RG23-C' AND EXTRA3='True' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ")  AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     
     M_CHAP = ""
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM ITMMST WHERE CODE='" & PLAREG!ICOD & "'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       M_CHAP = MSTDAT!igcd
       If MSTDAT.State = 1 Then MSTDAT.Close
       MSTDAT.Open "SELECT * FROM IGMMST WHERE CODE='" & M_CHAP & "'", CN, adOpenDynamic, adLockOptimistic
       If Not MSTDAT.EOF Then
         M_CHAP = MSTDAT!CHAP & ""
       End If
     End If
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = PLAREG!TTYP
     RS!TDAT = PLAREG!Date
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!CENVAT
     RS!CESS = PLAREG!CESS
     RS!ADUTY = PLAREG!A_DUTY
     RS!NCCD = PLAREG!NCCD
     RS!EDUCESS = PLAREG!EDUCESS
     RS!HEDCESS = PLAREG!H_ED_CESS
     RS!VTYP = PLAREG!VTYP
     RS!ICOD = PLAREG!ICOD
     RS!CHAP = M_CHAP
     RS!ASSV = PLAREG!ITOT
     RS!EXTRA1 = ComputerName
     RS!PCOD = PLAREG!DRAC
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT VBNO,PODT FROM PURMAN WHERE COMP='" & PLAREG!COMP & "' AND VTYP='" & PLAREG!VTYP & "' AND SRNO='" & PLAREG!SRNO & "'", CN, adOpenDynamic, adLockOptimistic
     If MSTDAT.EOF = False Then
       RS!VBNO = MSTDAT!VBNO
       If IsNull(MSTDAT!PODT) = True Then
        RS!VBDT = PLAREG!Date
       Else
        RS!VBDT = MSTDAT!PODT
       End If
      Else
       RS!VBNO = Trim(PLAREG!chln)
       RS!VBDT = PLAREG!Date
     End If
     
     RS!QNTY = PLAREG!QNTY
     If PLAREG!DRAC = "ER1DAT" And PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = "ER1 For the Month"
     ElseIf PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = "PLAREG!EXTRA5"
     End If
     RS!EXTRA3 = PLAREG!SRNO & ""
     RS!VBDT = PLAREG!CHDT
     RS.Update
     CNTR = CNTR + 1
     PLAREG.MoveNext
    Loop
   Case "Service Tax"
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM ER1DATA WHERE RCOD='SRVTAX' AND VTYP='OPN' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ")", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = "SRVTAX"
     RS!TDAT = PLAREG!TDAT
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!opncenvat
     RS!CESS = PLAREG!OPNCESS
     RS!ADUTY = PLAREG!OPNADUTY
     RS!NCCD = PLAREG!OPNNCCD
     RS!EDUCESS = PLAREG!OPNEDUCESS
     RS!HEDCESS = PLAREG!opnhredcess
     RS!VTYP = "OPN"
     RS!ICOD = "YYYYYYYYYY"
     RS!CHAP = ""
     RS!ASSV = 0
     RS!PCOD = "OPNBAL"
     RS!EXTRA1 = ComputerName
     RS.Update
     PLAREG.MoveNext
     CNTR = CNTR + 1
    Loop
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM EGPMAN WHERE (TTYP='SRVTAX' OR TTYP='SERVICE TAX') AND EXTRA3='True' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ") and recstat='A'", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     
     M_CHAP = ""
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM ITMMST WHERE CODE='" & PLAREG!ICOD & "'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       M_CHAP = MSTDAT!igcd
       If MSTDAT.State = 1 Then MSTDAT.Close
       MSTDAT.Open "SELECT * FROM IGMMST WHERE CODE='" & M_CHAP & "'", CN, adOpenDynamic, adLockOptimistic
       If Not MSTDAT.EOF Then
         M_CHAP = MSTDAT!CHAP & ""
       End If
     End If
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = "SRVTAX"
     RS!TDAT = PLAREG!Date
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!CENVAT
     RS!CESS = PLAREG!CESS
     RS!ADUTY = PLAREG!A_DUTY
     RS!NCCD = PLAREG!NCCD
     RS!EDUCESS = PLAREG!EDUCESS
     RS!HEDCESS = PLAREG!H_ED_CESS
     RS!VTYP = PLAREG!VTYP
     RS!ICOD = PLAREG!ICOD
     RS!CHAP = M_CHAP
     RS!ASSV = PLAREG!ITOT
     RS!EXTRA1 = ComputerName
     RS!PCOD = PLAREG!DRAC
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT VBNO,PODT FROM PURMAN WHERE COMP='" & PLAREG!COMP & "' AND VTYP='" & PLAREG!VTYP & "' AND SRNO='" & PLAREG!SRNO & "'", CN, adOpenDynamic, adLockOptimistic
     If MSTDAT.EOF = False Then
       RS!VBNO = MSTDAT!VBNO
       If IsNull(MSTDAT!PODT) = True Then
        RS!VBDT = PLAREG!Date
       Else
        RS!VBDT = MSTDAT!PODT
       End If
      Else
       RS!VBNO = PLAREG!VBNO
       RS!VBDT = PLAREG!Date
     End If
     RS!QNTY = PLAREG!QNTY
     If PLAREG!DRAC = "ER1DAT" And PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = "ER1 For the Month"
     ElseIf PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = "PLAREG!EXTRA5"
     End If
     RS!EXTRA3 = PLAREG!SRNO & ""
     RS.Update
     CNTR = CNTR + 1
     PLAREG.MoveNext
    Loop
   Case "PLA Register"
    '---------------------------------
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM ER1DATA WHERE RCOD='PLAREG' AND VTYP='OPN' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ")", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = "PLAREG"
     RS!TDAT = PLAREG!TDAT
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!opncenvat
     RS!CESS = PLAREG!OPNCESS
     RS!ADUTY = PLAREG!OPNADUTY
     RS!NCCD = PLAREG!OPNNCCD
     RS!EDUCESS = PLAREG!OPNEDUCESS
     RS!HEDCESS = PLAREG!opnhredcess
     RS!VTYP = "OPN"
     RS!ICOD = "YYYYYYYYYY"
     RS!CHAP = ""
     RS!ASSV = 0
     RS!PCOD = "OPNBAL"
     RS!EXTRA1 = ComputerName
     
     RS.Update
     PLAREG.MoveNext
     CNTR = CNTR + 1
    Loop
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM EGPMAN WHERE TTYP='PLAREG' AND COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ")  AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     
     M_CHAP = ""
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM ITMMST WHERE CODE='" & PLAREG!ICOD & "'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       M_CHAP = MSTDAT!igcd
       If MSTDAT.State = 1 Then MSTDAT.Close
       MSTDAT.Open "SELECT * FROM IGMMST WHERE CODE='" & M_CHAP & "'", CN, adOpenDynamic, adLockOptimistic
       If Not MSTDAT.EOF Then
         M_CHAP = MSTDAT!CHAP & ""
       End If
     End If
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = PLAREG!TTYP
     RS!TDAT = PLAREG!Date
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!CENVAT
     RS!CESS = PLAREG!CESS
     RS!ADUTY = PLAREG!A_DUTY
     RS!NCCD = PLAREG!NCCD
     RS!EDUCESS = PLAREG!EDUCESS
     RS!HEDCESS = PLAREG!H_ED_CESS
     RS!VTYP = PLAREG!VTYP
     RS!ICOD = PLAREG!ICOD
     RS!CHAP = M_CHAP
     RS!ASSV = PLAREG!ITOT
     RS!PCOD = PLAREG!DRAC
     RS!EXTRA1 = ComputerName
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT VBNO,PODT FROM PURMAN WHERE COMP='" & PLAREG!COMP & "' AND VTYP='" & PLAREG!VTYP & "' AND SRNO='" & PLAREG!SRNO & "'", CN, adOpenDynamic, adLockOptimistic
     If MSTDAT.EOF = False Then
       RS!VBNO = MSTDAT!VBNO
       RS!VBDT = MSTDAT!PODT
      Else
       RS!VBNO = PLAREG!VBNO
       RS!VBDT = PLAREG!Date
     End If
     RS!QNTY = PLAREG!QNTY
     If PLAREG!DRAC = "ER1DAT" And PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = "ER1 For the Month"
     ElseIf PLAREG!VTYP = "EXD" Then
       RS!EXTRA2 = ""
     End If
     RS!EXTRA3 = PLAREG!SRNO & ""
     RS.Update
     CNTR = CNTR + 1
     PLAREG.MoveNext
    Loop
    'Platrn For TR6 Challan
    If PLAREG.State = 1 Then PLAREG.Close
    PLAREG.Open "SELECT * FROM PLATRN WHERE COMP='" & compPth & "' AND UNIT IN (" & txtUNIT.Tag & ") AND RCOD='PLAREG'", CN, adOpenDynamic, adLockOptimistic
    Do While Not PLAREG.EOF
     
     M_CHAP = ""
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM PLATEMP WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     RS.AddNew
     RS!COMP = PLAREG!COMP
     RS!unit = PLAREG!unit
     RS!RCOD = "PLAREG"
     RS!TDAT = PLAREG!TDAT
     RS!TR6N = CNTR
     RS!CENVAT = PLAREG!CENVAT
     RS!CESS = PLAREG!CESS
     RS!ADUTY = PLAREG!ADUTY
     RS!NCCD = PLAREG!NCCD
     RS!EDUCESS = PLAREG!EDUCESS
     RS!HEDCESS = PLAREG!HEDCESS
     RS!VTYP = "TR6"
     RS!ICOD = ""
     RS!CHAP = M_CHAP
     RS!ASSV = 0
     RS!PCOD = "T.R.6"
     RS!EXTRA1 = ComputerName
     RS!EXTRA2 = "TR6 No. " + PLAREG!TR6N
     RS.Update
     CNTR = CNTR + 1
     PLAREG.MoveNext
    Loop
    '---------------------------------
  End Select
  Exit Sub
LAST:
  MsgBox ERR.Description
  Resume
End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub


Private Sub create_table()
 On Error Resume Next
 Dim crtsql As String
 crtsql = "CREATE TABLE [dbo].[ER1DATA] ([COMP] [char] (4) NOT NULL , " & _
            "[UNIT] [char] (6) NOT NULL ,  " & _
            "[VTYP] [char] (3) NOT NULL , " & _
            "[RCOD] [char] (6) NOT NULL , " & _
            "[TDAT] [datetime] NOT NULL , " & _
            "[ITOT] [decimal](18, 2) NOT NULL , " & _
            "[OPNCENVAT] [decimal](18, 2) NOT NULL , " & _
            "[OPNNCCD] [decimal](18, 2) NOT NULL , " & _
            "[OPNADUTY] [decimal](18, 2) NOT NULL , " & _
            "[OPNEDUCESS] [decimal](18, 2) NOT NULL , " & _
            "[OPNHREDCESS] [decimal](18, 2) NOT NULL , " & _
            "[OPNSTAX] [decimal](18, 2) NOT NULL , " & _
            "[OPNSTAX_EDCESS] [decimal](18, 2) NOT NULL , " & _
            "[OPNSTAX_HEDCESS] [decimal](18, 2) NOT NULL , " & _
            "[CENCENVAT] [decimal](18, 2) NOT NULL , " & _
            "[CENNCCD] [decimal](18, 2) NOT NULL , " & _
            "[CENADUTY] [decimal](18, 2) NOT NULL , " & _
            "[CENEDUCESS] [decimal](18, 2) NOT NULL , " & _
            "[CENHREDCESS] [decimal](18, 2) NOT NULL , " & _
            "[CENSTAX] [decimal](18, 2) NOT NULL , " & _
            "[CENSTAX_EDCESS] [decimal](18, 2) NOT NULL , " & _
            "[CENSTAX_HEDCESS] [decimal](18, 2) NOT NULL , " & _
            " [DEBCENVAT] [decimal](18, 2) NOT NULL , " & _
            "[DEBNCCD] [decimal](18, 2) NOT NULL , "
            
    crtsql = crtsql & "[DEBADUTY] [decimal](18, 2) NOT NULL , " & _
                   "[DEBEDUCESS] [decimal](18, 2) NOT NULL , " & _
                   "[DEBHREDCESS] [decimal](18, 2) NOT NULL , " & _
                   "[DEBSTAX] [decimal](18, 2) NOT NULL , " & _
                   "[DEBSTAX_EDCESS] [decimal](18, 2) NOT NULL , " & _
                   "[DEBSTAX_HEDCESS] [decimal](18, 2) NOT NULL , " & _
                   "[EXTRA1] [varchar] (50) NULL , " & _
                   "[EXTRA2] [varchar] (50) NULL , " & _
                   "[EXTRA3] [varchar] (50) NULL , " & _
                   "[EXTRA4] [varchar] (50) NULL , " & _
                   "[EXTRA5] [varchar] (50) NULL " & _
                   ") ON [PRIMARY] "

CN.Execute crtsql
crtsql = "ALTER TABLE [dbo].[ER1DATA] WITH NOCHECK ADD " & _
       "CONSTRAINT [PK_ER1DATA] PRIMARY KEY  CLUSTERED " & _
       "([COMP],[UNIT],[VTYP],[RCOD],[TDAT])  ON [PRIMARY] "

CN.Execute crtsql
crtsql = "ALTER TABLE [dbo].[ER1DATA] ADD " & _
       "CONSTRAINT [DF_ER1DATA_ITOT] DEFAULT (0) FOR [ITOT], " & _
       "CONSTRAINT [DF_ER1DATA_CENVAT] DEFAULT (0) FOR [OPNCENVAT], " & _
       "CONSTRAINT [DF_ER1DATA_NCCD] DEFAULT (0) FOR [OPNNCCD], " & _
       "CONSTRAINT [DF_ER1DATA_ADUTY] DEFAULT (0) FOR [OPNADUTY], " & _
       "CONSTRAINT [DF_ER1DATA_EDUCESS] DEFAULT (0) FOR [OPNEDUCESS], " & _
       "CONSTRAINT [DF_ER1DATA_HREDCESS] DEFAULT (0) FOR [OPNHREDCESS], " & _
       "CONSTRAINT [DF_ER1DATA_STAX] DEFAULT (0) FOR [OPNSTAX], " & _
       "CONSTRAINT [DF_ER1DATA_STAX_EDCESS] DEFAULT (0) FOR [OPNSTAX_EDCESS], " & _
       "CONSTRAINT [DF_ER1DATA_STAX_HEDCESS] DEFAULT (0) FOR [OPNSTAX_HEDCESS], " & _
       "CONSTRAINT [DF_ER1DATA_OPNCENVAT1] DEFAULT (0) FOR [CENCENVAT], " & _
       "CONSTRAINT [DF_ER1DATA_OPNNNCD1] DEFAULT (0) FOR [CENNCCD], " & _
       "CONSTRAINT [DF_ER1DATA_OPNADUTY1] DEFAULT (0) FOR [CENADUTY], " & _
       "CONSTRAINT [DF_ER1DATA_OPNEDUCESS1] DEFAULT (0) FOR [CENEDUCESS], " & _
       "CONSTRAINT [DF_ER1DATA_OPNHREDCESS1] DEFAULT (0) FOR [CENHREDCESS], " & _
       "CONSTRAINT [DF_ER1DATA_OPNSTAX1] DEFAULT (0) FOR [CENSTAX], " & _
       "CONSTRAINT [DF_ER1DATA_OPNSTAX_EDCESS1] DEFAULT (0) FOR [CENSTAX_EDCESS], " & _
       "CONSTRAINT [DF_ER1DATA_OPNSTAX_HEDCESS1] DEFAULT (0) FOR [CENSTAX_HEDCESS], " & _
       "CONSTRAINT [DF_ER1DATA_OPNCENVAT1_1] DEFAULT (0) FOR [DEBCENVAT], " & _
       "CONSTRAINT [DF_ER1DATA_OPNNNCD1_1] DEFAULT (0) FOR [DEBNCCD], " & _
       "CONSTRAINT [DF_ER1DATA_OPNADUTY1_1] DEFAULT (0) FOR [DEBADUTY], " & _
       "CONSTRAINT [DF_ER1DATA_OPNEDUCESS1_1] DEFAULT (0) FOR [DEBEDUCESS], " & _
       "CONSTRAINT [DF_ER1DATA_OPNHREDCESS1_1] DEFAULT (0) FOR [DEBHREDCESS], " & _
       "CONSTRAINT [DF_ER1DATA_OPNSTAX1_1] DEFAULT (0) FOR [DEBSTAX], "
crtsq = crtsq & "CONSTRAINT [DF_ER1DATA_OPNSTAX_EDCESS1_1] DEFAULT (0) FOR [DEBSTAX_EDCESS], " & _
               "CONSTRAINT [DF_ER1DATA_OPNSTAX_HEDCESS1_1] DEFAULT (0) FOR [DEBSTAX_HEDCESS] "

CN.Execute crtsql

End Sub

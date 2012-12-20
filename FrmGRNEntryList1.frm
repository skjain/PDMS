VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmGRNEntryList1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRN List"
   ClientHeight    =   6600
   ClientLeft      =   2070
   ClientTop       =   2250
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9120
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   8955
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
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
         Left            =   6600
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1320
         TabIndex        =   9
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
         Format          =   24313857
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4200
         TabIndex        =   10
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
         Format          =   24313857
         CurrentDate     =   38429
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date: "
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
         Left            =   255
         TabIndex        =   3
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date: "
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
         Left            =   3315
         TabIndex        =   4
         Top             =   285
         Width           =   825
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   8895
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   7
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         TabIndex        =   8
         Top             =   195
         Width           =   1035
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Bindings        =   "FrmGRNEntryList1.frx":0000
      Height          =   4815
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8493
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      ForeColorFixed  =   128
      ForeColorSel    =   -2147483633
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSAdodcLib.Adodc ADOHELP 
      Height          =   330
      Left            =   120
      Top             =   6720
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label LBLDAYBOK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "DAYBOOK: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   5040
   End
   Begin VB.Label LBLDIVNAM 
      BackColor       =   &H00C0E0FF&
      Caption         =   "DIVISION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "FrmGRNEntryList1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  FrmGRNEntry.TXTVBNO = Empty
  Unload Me
End Sub

Private Sub CMDOK_Click()
  Dim SEL_DBCD As String, SEL_VBNO As String
  
  If FLEX.Rows <= 1 Then Exit Sub
  If FLEX.TextMatrix(1, 1) = Empty Then Exit Sub
  
  SEL_VBNO = FLEX.TextMatrix(FLEX.ROW, 6)
  SEL_DBCD = FrmGRNEntry.M_DBCD_DIRIVR
  
  If Trim(SEL_VBNO) = Empty Then
     lstBill.SetFocus
     Exit Sub
  End If
  Dim EDTDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  Dim SQL As String
  Dim SQL1 As String
  With FrmGRNEntry
  SQL = "SELECT * FROM " & .SUMMARYTABLE & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND VTYP='IVR' AND DBCD='" & SEL_DBCD & "' AND VBNO='" & SEL_VBNO & "' AND RECSTAT<>'D'"
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
     Exit Sub
  End If
   
   'IMPORT CASE
     .txtCURNCY = EDTDAT!Currency
     .txtEXRate = EDTDAT!EXRATE
     .txtcha = EDTDAT!CHAVALUE
     .txtfrt = EDTDAT!FRTVALUE
     .txtdty = EDTDAT!DTYVALUE
    '-------------------------------
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ACCMST WHERE CODE='" & EDTDAT!DRAC & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTDBAC = Trim(MSTDAT!NAME & "")
     Else
      .TXTDBAC.Text = Empty
    End If
       
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM TRANSPORTMST WHERE CODE='" & EDTDAT!TRCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTTRNM = Trim(MSTDAT!NAME & "")
     Else
      .TXTTRNM = Empty
    End If
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM TAXMST WHERE CODE='" & EDTDAT!TXCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTTAXNAM = Trim(MSTDAT!NAME & "")
     Else
      .TXTTAXNAM = Empty
    End If
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM LOCMST WHERE CODE='" & EDTDAT!GDNCOD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTGDN = Trim(MSTDAT!NAME & "")
     Else
      .TXTGDN = Empty
    End If
    
    
    .TXTRTORTAX = EDTDAT!RORT & ""
    .cmbSelection = Trim(EDTDAT!TTYP & "")
    
    .TXTSCHLN = Trim(EDTDAT!chln & "")
    
    If IsNull(EDTDAT!CHDT) Then
     Else
      .TXTSCHLNDT = EDTDAT!CHDT & ""
    End If
                
    .TXTVBNO = Trim(EDTDAT!VBNO & "")
    .TXTVBDT = EDTDAT!Date
    
    .TXTLRNO = Trim(EDTDAT!LRNO & "")
    
    If EDTDAT!ACEFFECT = "Y" Then
      .chkAcEffect.Value = 1
    Else
      .chkAcEffect.Value = 0
    End If
    
    If Not IsNull(EDTDAT!LRDT) Then
      .TXTLRDT = EDTDAT!LRDT
    End If
    .TXTVHCL = Trim(EDTDAT!VHCL & "")
    .TXTRMRK = Trim(EDTDAT!BRMK & "")
    .TXTTPCS = EDTDAT!TPCS
    
    .TXTSBILLNO = Trim(EDTDAT!CVBN & "")
    
    If IsNull(EDTDAT!GATD) Then
     Else
      .TXTSBILLDATE = EDTDAT!GATD & ""
    End If
    
    .TXTTQTY = Format(EDTDAT!TQTY, "######.000")
    .TXTITOT = Format(EDTDAT!ITOT, "##########.00")
    .TXTBNET = Format(EDTDAT!BNET, "##########.00")
    Dim i As Double
    Dim J As Double
    i = 0
    For i = 0 To .flexBTRM.Rows - 1
      J = 0
      For J = 0 To EDTDAT.Fields.COUNT - 1
        If Trim(EDTDAT.Fields(J).NAME) = Trim(.flexBTRM.TextMatrix(i, 0)) Then
            .flexBTRM.TextMatrix(i, 2) = Format(EDTDAT.Fields(J).Value, "#########.00")
        End If
        If Trim(EDTDAT.Fields(J).NAME) = "PER" & Trim(.flexBTRM.TextMatrix(i, 0)) Then
           .flexBTRM.TextMatrix(i, 1) = Format(EDTDAT.Fields(J).Value, "######.00")
        End If
        If Trim(EDTDAT.Fields(J).NAME) = Trim(.flexBTRM.TextMatrix(i, 0)) & "_INCOST" Then
           .flexBTRM.TextMatrix(i, 3) = Trim(EDTDAT.Fields(J).Value)
        End If
      Next
    Next
    
    'FIFO-----------------------------
      If FIFOREQ = "Y" Then
         If RS.State = 1 Then RS.Close
         RS.Open "SELECT * FROM GRNTRAN WHERE COMP ='" & compPth & "' AND UNIT ='" & UNCD & _
         "' AND DBCD='" & EDTDAT!dbcd & "" & "' AND VBNO = '" & EDTDAT!VBNO & "" & "' AND GRN_QNTY <> BAL_QNTY ", CN, adOpenDynamic, adLockReadOnly
         If RS.EOF = False Then .M_ISSUE = "Y"
      End If
    '---------------------------------
        
    .FLEX.Rows = 2
    SQL = "SELECT " & .TABLENAME & ".*,ITMMST.NAME FROM " & .TABLENAME & " INNER JOIN ITMMST ON " & .TABLENAME & ".ICOD=ITMMST.CODE WHERE " & .TABLENAME & ".COMP='" & compPth & _
    "' AND " & .TABLENAME & ".UNIT='" & UNCD & "' AND " & .TABLENAME & ".VTYP='IVR' AND " & .TABLENAME & ".DBCD='" & SEL_DBCD & _
    "' AND " & .TABLENAME & ".VBNO='" & SEL_VBNO & "' AND RECSTAT<>'D' ORDER BY " & .TABLENAME & ".SRCH"
           
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL, CN, adOpenForwardOnly, adLockOptimistic
    i = 1
    Do While Not EDTDAT.EOF
     .FLEX.TextMatrix(i, 0) = EDTDAT!SRCH
     .FLEX.TextMatrix(i, 1) = EDTDAT!chln & ""
     If Not IsNull(EDTDAT!CHDT) Then
       .FLEX.TextMatrix(i, 2) = EDTDAT!CHDT
      Else
       .FLEX.TextMatrix(i, 2) = ""
     End If
     
     .FLEX.TextMatrix(i, 3) = EDTDAT!NAME & ""
     .FLEX.TextMatrix(i, 4) = EDTDAT!ltno & ""
     .FLEX.TextMatrix(i, 5) = EDTDAT!grad & ""
     .FLEX.TextMatrix(i, 6) = EDTDAT!COPS
     .FLEX.TextMatrix(i, 7) = EDTDAT!PCES
     
     .FLEX.TextMatrix(i, 8) = Format(EDTDAT!QNTY, "########.000")
     .FLEX.TextMatrix(i, 9) = Format(EDTDAT!GWGT, "######.0000")
     .FLEX.TextMatrix(i, 10) = Val(EDTDAT!QNTY) * Val(EDTDAT!GWGT) 'AMOUNT
     
     If UCase(Trim(M_Currency)) <> UCase(Trim(.txtCURNCY)) Then
        If Val(.txtEXRate) > 0 Then
           .FLEX.TextMatrix(i, 9) = Format(EDTDAT!GWGT / Val(.txtEXRate), "######.00")
           .FLEX.TextMatrix(i, 10) = Val(EDTDAT!QNTY) * Val(EDTDAT!GWGT) 'AMOUNT
        End If
     End If
       
     .FLEX.TextMatrix(i, 11) = EDTDAT!ICOD
     .FLEX.TextMatrix(i, 12) = EDTDAT!RTYP & ""
     .FLEX.TextMatrix(i, 13) = EDTDAT!RSRN & ""
     .FLEX.TextMatrix(i, 16) = EDTDAT!TWST & ""
     .FLEX.TextMatrix(i, 17) = EDTDAT!RSRC & ""
     
     If IsNull(EDTDAT!RTYP) Or Trim(EDTDAT!RTYP) = "" Then
       .ALLOWEDITDEL = True
      Else
       .ALLOWEDITDEL = False
     End If
        
     EDTDAT.MoveNext
     i = i + 1
     If Not EDTDAT.EOF Then
       If i > 1 Then
           .FLEX.Rows = .FLEX.Rows + 1
       End If
     End If
    Loop
    
    EDTDAT.MovePrevious
    
    Dim RETRS As ADODB.Recordset
    Set RETRS = New ADODB.Recordset
    
    Dim NOPLY As Long: NOPLY = 0
    
    If RETRS.State = 1 Then RETRS.Close
    SQL = "SELECT * FROM PKGSTK WHERE COMP='" & compPth & "' AND VTYP='IVR' AND DBCD='000001' AND UNIT='" & UNCD & _
          "' AND CHLN= '" & Trim(EDTDAT!VBNO & "") & "' "
    RETRS.Open SQL, CN, adOpenForwardOnly, adLockOptimistic
    If Not RETRS.EOF Then
       .chkReturnable.Value = 1
       .TxtPallet = Val(RETRS!TOPPLY)
       .txtCops = Val(RETRS!QNTY)
              
       i = 0
       For i = 1 To .FLEXPLY.Cols - 1
         J = 0
            For J = 0 To RETRS.Fields.COUNT - 1
               If Trim(RETRS.Fields(J).NAME) = Trim(.FLEXPLY.TextMatrix(0, i)) Then
                  .FLEXPLY.TextMatrix(1, i) = Val(RETRS.Fields(J).Value)
                  NOPLY = NOPLY + Val(RETRS.Fields(J).Value)
               End If
            Next
       Next
       .chkReturnable.Visible = False
       .txtPly = NOPLY
    Else
       .chkReturnable.Value = 0
       .chkReturnable.Visible = True
    End If
               
    .TXTVBNO = SEL_VBNO
    .TXTDBAC.Enabled = True
  End With
  Unload Me
End Sub

Private Sub FLEX_DblClick()
Call CMDOK_Click
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
  LBLDIVNAM.Caption = DIVNAM
  LBLDAYBOK.Caption = FrmGRNEntry.Caption
  txtFrDate = GetMinDate
  txtToDate = GetMaxDate
  Me.KeyPreview = True
  cmdOk.Enabled = False
  cmdCancel.Enabled = True
End Sub

Private Sub cmdGo_Click()
 'lstBill.ListItems.Clear
 Call SetInitial
  
 With FrmGRNEntry
  SQL = "SELECT DISTINCT CONVERT(VarChar(10), " & .SUMMARYTABLE & ".DATE, 103) As DATE," & .SUMMARYTABLE & ".VBNO AS GRN,ACCMST.NAME AS PARTY," & _
  " " & .SUMMARYTABLE & ".TQTY AS TOTAL_QTY," & .SUMMARYTABLE & ".BNET AS NET_AMOUNT," & _
  " " & .SUMMARYTABLE & ".VTYP AS VTYP," & .SUMMARYTABLE & ".VBNO AS VBNO FROM " & .SUMMARYTABLE & " " & _
  " INNER JOIN ACCMST ON " & .SUMMARYTABLE & ".PCOD=ACCMST.CODE " & _
  " INNER JOIN " & .TABLENAME & " ON " & .SUMMARYTABLE & ".COMP=" & .TABLENAME & ".COMP " & _
  " AND " & .SUMMARYTABLE & ".UNIT=" & .TABLENAME & ".UNIT AND " & .SUMMARYTABLE & ".VTYP=" & .TABLENAME & ".VTYP AND " & .SUMMARYTABLE & ".DBCD=" & .TABLENAME & ".DBCD " & _
  " AND " & .SUMMARYTABLE & ".VBNO=" & .TABLENAME & ".VBNO " & _
  " WHERE " & .SUMMARYTABLE & ".COMP='" & compPth & "' AND " & .SUMMARYTABLE & ".UNIT='" & UNCD & _
  "' AND " & .SUMMARYTABLE & ".VTYP='IVR' " & _
  " AND " & .SUMMARYTABLE & ".BSTS='P' AND " & .SUMMARYTABLE & ".DVCD='" & DIVCOD & "' AND " & .SUMMARYTABLE & ".DBCD='" & FrmGRNEntry.M_DBCD_DIRIVR & _
  "' AND " & .SUMMARYTABLE & ".DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
  "' AND " & .SUMMARYTABLE & ".DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
  "' AND " & .SUMMARYTABLE & ".RECSTAT<>'D' and " & .SUMMARYTABLE & " .extra1 IS NULL  AND " & .SUMMARYTABLE & ".UNIT='" & UNCD & "' "
  '"' ORDER BY " & .SUMMARYTABLE & ".DATE," & .SUMMARYTABLE & ".VBNO"
 End With
  
 ADOHELP.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
 ADOHELP.CommandType = adCmdText
 ADOHELP.RecordSource = SQL
 ADOHELP.Refresh
   
 Call SETFLEX
  
 cmdOk.Enabled = True
 cmdOk.Default = True
  
  If FrmGRNEntryList1.Visible = True Then
     If FLEX.Rows > 1 Then
        If FLEX.TextMatrix(1, 0) <> Empty Then
           FLEX.ROW = 1
           FLEX.COL = 0
           FLEX.SetFocus
        End If
     End If
  End If
  
End Sub

Private Sub lstBill_GotFocus()
lstBill.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstBill_LostFocus()
lstBill.BackColor = vbWhite
End Sub

Private Sub txtFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub SetInitial()
  FLEX.FixedCols = 0
  If FLEX.Rows > 1 Then FLEX.FixedRows = 1
  FLEX.Cols = 2
  FLEX.Rows = 2
End Sub

Private Sub SETFLEX()
    With FLEX
        .Cols = 7
        If .Rows > 1 Then .FixedRows = 1
                
        .TextMatrix(0, 0) = "Date"
        .TextMatrix(0, 1) = "GRN"
        .TextMatrix(0, 2) = "Party"
        .TextMatrix(0, 3) = "Total Qty"
        .TextMatrix(0, 4) = "Net Amt."
        .TextMatrix(0, 5) = "Vtyp"
        .TextMatrix(0, 6) = "VBNO"
               
        .ColWidth(0) = 1200
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        
          
    End With
End Sub

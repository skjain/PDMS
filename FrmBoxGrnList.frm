VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form FrmBoxGrnList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box wise GRN Help List"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framCmd 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   9375
      Begin WelchButton.lvButtons_H cmdOk 
         Height          =   375
         Left            =   6600
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&OK"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   7920
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   9435
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52428801
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52428801
         CurrentDate     =   38429
      End
      Begin WelchButton.lvButtons_H cmdGo 
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Search"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3315
         TabIndex        =   5
         Top             =   285
         Width           =   930
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   285
         Width           =   1185
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Bindings        =   "FrmBoxGrnList.frx":0000
      Height          =   4815
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
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
         Name            =   "Times New Roman Greek"
         Size            =   11.25
         Charset         =   161
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
      Top             =   6600
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
   Begin VB.Label LBLDIVNAM 
      BackColor       =   &H00C0E0FF&
      Caption         =   "DIVISION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4560
   End
   Begin VB.Label LBLDAYBOK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "DAYBOOK: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   0
      Width           =   5160
   End
End
Attribute VB_Name = "FrmBoxGrnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  FRMBOXGRN.M_SRNO = Empty
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim SEL_SRNO As String
  SEL_SRNO = FLEX.TextMatrix(FLEX.ROW, 6)
  If Trim(SEL_SRNO) = Empty Then
    lstBill.SetFocus
     Exit Sub
  End If
  Dim EDTDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  Dim SQL As String
  Dim SQL1 As String
  With FRMBOXGRN
  SQL = "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='IVR' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D'"
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
'     lstBill.SetFocus
     Exit Sub
  End If
  
     'IMPORT CASE
     '.txtCURNCY = EDTDAT!Currency
     ''.txtEXRate = EDTDAT!EXRATE
     '.txtcha = EDTDAT!CHAVALUE
     '.txtfrt = EDTDAT!FRTVALUE
     '.txtdty = EDTDAT!DTYVALUE
    '-------------------------------
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ACCMST WHERE CODE='" & EDTDAT!DRAC & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTDBAC = MSTDAT!NAME & ""
     Else
      .TXTDBAC.Text = Empty
    End If
       
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM TRANSPORTMST WHERE CODE='" & EDTDAT!TRCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTTRNM = MSTDAT!NAME & ""
     Else
      .TXTTRNM = Empty
    End If


   If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM LOCMST  WHERE CODE='" & EDTDAT!GDNCOD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTGDN = MSTDAT!NAME & ""
     Else
      .TXTGDN = Empty
    End If

    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM TAXMST WHERE CODE='" & EDTDAT!TXCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTTAXNAM = MSTDAT!NAME & ""
     Else
      .TXTTAXNAM = Empty
    End If
    
    .TXTRTORTAX = EDTDAT!RORT & ""
    .cmbSelection = EDTDAT!TTYP & ""
    
    .TXTSCHLN = Trim(EDTDAT!CHLN & "")
    
    If IsNull(EDTDAT!CHDT) Then
     Else
      .TXTSCHLNDT = EDTDAT!CHDT & ""
    End If
                
    .TXTVBNO = EDTDAT!VBNO & ""
    .TXTVBDT = EDTDAT!Date
    
    .TXTLRNO = EDTDAT!LRNO & ""
    If Not IsNull(EDTDAT!LRDT) Then
      .TXTLRDT = EDTDAT!LRDT
    End If
    .TXTVHCL = EDTDAT!VHCL & ""
    .TXTRMRK = EDTDAT!BRMK & ""
    .TXTTPCS = EDTDAT!TPCS
    
    .TXTSBILLNO = EDTDAT!CVBN & ""
    
    If IsNull(EDTDAT!GATD) Then
     Else
      .TXTSBILLDATE = EDTDAT!GATD & ""
    End If
    
    .TXTTQTY = Format(EDTDAT!TQTY, "######.000")
    .TXTITOT = Format(EDTDAT!ITOT, "##########.00")
    .TXTBNET = Format(EDTDAT!BNET, "##########.00")
    Dim I As Double
    Dim J As Double
    I = 0
    For I = 0 To .flexBTRM.Rows - 1
      J = 0
      For J = 0 To EDTDAT.Fields.COUNT - 1
        If Trim(EDTDAT.Fields(J).NAME) = Trim(.flexBTRM.TextMatrix(I, 0)) Then
            .flexBTRM.TextMatrix(I, 2) = Format(EDTDAT.Fields(J).Value, "#########.00")
        End If
        If Trim(EDTDAT.Fields(J).NAME) = "PER" & Trim(.flexBTRM.TextMatrix(I, 0)) Then
           .flexBTRM.TextMatrix(I, 1) = Format(EDTDAT.Fields(J).Value, "######.00")
        End If
        
        If Trim(EDTDAT.Fields(J).NAME) = Trim(.flexBTRM.TextMatrix(I, 0)) & "_INCOST" Then
           .flexBTRM.TextMatrix(I, 3) = Trim(EDTDAT.Fields(J).Value)
        End If
        
      Next
    Next
   
    'FIFO-----------------------------
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM GRNTRAN WHERE COMP ='" & compPth & "' AND UNIT ='" & UNCD & _
     "' AND DBCD='" & EDTDAT!dbcd & "" & "' AND VBNO = '" & EDTDAT!VBNO & "" & "' AND GRN_QNTY <> BAL_QNTY ", CN, adOpenDynamic, adLockReadOnly
     If RS.EOF = False Then .M_ISSUE = "Y"
    '---------------------------------
    
    
    If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM TRDBOXREGISTER WHERE COMP ='" & compPth & "' AND UNIT ='" & UNCD & _
     "' AND DBCD='" & EDTDAT!dbcd & "" & "' AND GRNNO = '" & EDTDAT!VBNO & "" & "' AND RVBNO IS NOT NULL", CN, adOpenDynamic, adLockReadOnly
     If RS.EOF = False Then .M_ISSUE = "Y"
        
     .FLEX.Rows = 2
     SQL = "SELECT STORETRAN.*,ITMMST.NAME FROM STORETRAN INNER JOIN ITMMST ON STORETRAN.ICOD=ITMMST.CODE WHERE STORETRAN.COMP='" & compPth & "' AND STORETRAN.VTYP='IVR' AND STORETRAN.SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D' ORDER BY STORETRAN.SRCH"
             
     If EDTDAT.State = 1 Then EDTDAT.Close
        EDTDAT.Open SQL, CN, adOpenForwardOnly, adLockOptimistic
     I = 1
     Do While Not EDTDAT.EOF
        .FLEX.TextMatrix(I, 0) = EDTDAT!SRCH
        .FLEX.TextMatrix(I, 1) = EDTDAT!CHLN & ""
     If Not IsNull(EDTDAT!CHDT) Then
        .FLEX.TextMatrix(I, 2) = EDTDAT!CHDT
      Else
        .FLEX.TextMatrix(I, 2) = ""
     End If
     
     .FLEX.TextMatrix(I, 3) = EDTDAT!NAME & ""
     .FLEX.TextMatrix(I, 4) = Trim(EDTDAT!ltno & "")
     .FLEX.TextMatrix(I, 5) = EDTDAT!grad & ""
     .FLEX.TextMatrix(I, 6) = EDTDAT!COPS
     .FLEX.TextMatrix(I, 7) = EDTDAT!PCES
     .FLEX.TextMatrix(I, 8) = Format(EDTDAT!QNTY, "########.000")
     
    .FLEX.TextMatrix(I, 9) = Format(EDTDAT!GWGT, "######.0000")
     .FLEX.TextMatrix(I, 10) = Val(EDTDAT!QNTY) * Val(EDTDAT!GWGT) 'AMOUNT
     
    ' If UCase(Trim(M_Currency)) <> UCase(Trim(.txtCURNCY)) Then
    '    If Val(.txtEXRate) > 0 Then
    '       .Flex.TextMatrix(I, 9) = Format(EDTDAT!GWGT / Val(.txtEXRate), "######.00")
    '       .Flex.TextMatrix(I, 10) = Val(EDTDAT!QNTY) * Val(EDTDAT!GWGT) 'AMOUNT
    '    End If
    ' End If
     .FLEX.TextMatrix(I, 11) = EDTDAT!ICOD
     .FLEX.TextMatrix(I, 12) = EDTDAT!RTYP & ""
     .FLEX.TextMatrix(I, 13) = EDTDAT!RSRN & ""
     .FLEX.TextMatrix(I, 16) = EDTDAT!TWST & ""
     .FLEX.TextMatrix(I, 17) = EDTDAT!RSRC & ""
     If IsNull(EDTDAT!RTYP) Or Trim(EDTDAT!RTYP) = "" Then
       .ALLOWEDITDEL = True
     Else
       .ALLOWEDITDEL = False
     End If
        
        
                If frmBundelDetails.mfgBndlDet.Cols < (I * 4) + 1 Then
                    frmBundelDetails.mfgBndlDet.Cols = frmBundelDetails.mfgBndlDet.Cols + 4
                End If
                
                If RS.State = 1 Then RS.Close
                RS.Open "SELECT * FROM TRDBOXREGISTER WHERE COMP='" & EDTDAT!COMP & "' AND UNIT='" & _
                    EDTDAT!unit & "' AND DVCD='" & EDTDAT!DVCD & "' AND DBCD='" & EDTDAT!dbcd & _
                    "' AND VTYP='" & EDTDAT!VTYP & "' AND GRNNO = '" & EDTDAT!VBNO & "' AND ICOD = '" & EDTDAT!ICOD & "'", CN, adOpenDynamic, adLockReadOnly
                J = 1
                Do While Not RS.EOF
                    
                    With frmBundelDetails.mfgBndlDet
                        .TextMatrix(J, (I * 4) - 3) = Trim(RS!VBNO)
                        .TextMatrix(J, (I * 4) - 2) = Trim(RS!GRSWGT)
                        .TextMatrix(J, (I * 4) - 1) = Trim(RS!TRWGT)
                        .TextMatrix(J, I * 4) = Format(Val(RS!NTWGT), "######.000")
                    End With
                    J = J + 1
                    RS.MoveNext
                Loop

     EDTDAT.MoveNext
     I = I + 1
     If Not EDTDAT.EOF Then
       If I > 1 Then
           .FLEX.Rows = .FLEX.Rows + 1
       End If
     End If
    Loop
    '-------------------------------------
    EDTDAT.MovePrevious
    
    Dim RETRS As ADODB.Recordset
    Set RETRS = New ADODB.Recordset
    
    Dim NOPLY As Long: NOPLY = 0
    
    If RETRS.State = 1 Then RETRS.Close
    SQL = "SELECT * FROM PKGSTK WHERE COMP='" & compPth & "' AND VTYP='IVR' AND DBCD='000001' AND UNIT='" & UNCD & _
          "' AND CHLN= '" & Trim(EDTDAT!VBNO & "") & "' "
    RETRS.Open SQL, CN, adOpenForwardOnly, adLockOptimistic
    If Not RETRS.EOF Then
       .FRMLRDTL.Tab = 1
       .TxtPallet = Val(RETRS!TOPPLY)
       .txtCops = Val(RETRS!QNTY)
              
       I = 0
       For I = 1 To .FLEXPLY.Cols - 1
         J = 0
            For J = 0 To RETRS.Fields.COUNT - 1
               If Trim(RETRS.Fields(J).NAME) = Trim(.FLEXPLY.TextMatrix(0, I)) Then
                  .FLEXPLY.TextMatrix(1, I) = Val(RETRS.Fields(J).Value)
                  NOPLY = NOPLY + Val(RETRS.Fields(J).Value)
               End If
            Next
       Next
      ' .ChkReturnable.Visible = False
       .txtPly = NOPLY
    Else
       .FRMBTRM.TabIndex = 1
     '  .ChkReturnable.Visible = True
    End If

    
    
    '-------------------------------------
    .M_SRNO = SEL_SRNO
    .TXTDBAC.Enabled = True
  End With
  Unload Me
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
  LBLDIVNAM.Caption = DIVNAM
  LBLDAYBOK.Caption = FRMBOXGRN.Caption
  txtFrDate = GetMinDate
  txtToDate = GetMaxDate
  Me.KeyPreview = True
  CMDOK.Enabled = False
  cmdCancel.Enabled = True
End Sub

Private Sub cmdGo_Click()
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Dim SQL As String
  Dim SQL1 As String
  SQL = Empty
 
 
 With FRMBOXGRN
  SQL = "SELECT DISTINCT CONVERT(VarChar(10), GRN.DATE, 103) As DATE,GRN.VBNO AS GRN,ACCMST.NAME AS PARTY," & _
        " GRN.TQTY AS TOTAL_QTY,GRN.BNET AS NET_AMOUNT," & _
        " GRN.VTYP AS VTYP,GRN.SRNO AS SRNO FROM GRN " & _
        " INNER JOIN ACCMST ON GRN.PCOD=ACCMST.CODE " & _
        " INNER JOIN STORETRAN ON GRN.COMP=STORETRAN.COMP " & _
        " AND GRN.VTYP=STORETRAN.VTYP AND GRN.SRNO=STORETRAN.SRNO " & _
        " WHERE GRN.COMP='" & compPth & "' AND GRN.VTYP='IVR' " & _
        " AND GRN.BSTS='P' AND GRN.DVCD='" & DIVCOD & "' AND GRN.DBCD='" & FRMBOXGRN.M_DBCD_DIRIVR & _
        "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.RECSTAT<>'D'  AND GRN.UNIT='" & UNCD & "' AND GRN.EXTRA1 = 'BOX'"
 End With
  
 ADOHELP.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
 ADOHELP.CommandType = adCmdText
 ADOHELP.RecordSource = SQL
 ADOHELP.Refresh
   
 Call SETFLEX
  
 CMDOK.Enabled = True
 CMDOK.Default = True
  
  If FrmBoxGrnList.Visible = True Then
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
Private Sub SETFLEX()
    With FLEX
        .Cols = 7
        If .Rows > 1 Then .FixedRows = 1
            .TextMatrix(0, 0) = "Date"
            .TextMatrix(0, 1) = "GRN"
            .TextMatrix(0, 2) = "Party"
            .TextMatrix(0, 3) = "Total NetWt."
            .TextMatrix(0, 4) = "Net Amt."
            .TextMatrix(0, 5) = "Vtyp"
            .TextMatrix(0, 6) = "Srno"
               
            .ColWidth(0) = 1200
            .ColWidth(1) = 1500
            .ColWidth(2) = 3500
            .ColWidth(3) = 1200
            .ColWidth(4) = 1200
            .ColWidth(5) = 0
            .ColWidth(6) = 0
          
    End With
End Sub



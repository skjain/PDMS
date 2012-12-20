VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_SPLISTIVR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRN List"
   ClientHeight    =   4545
   ClientLeft      =   2070
   ClientTop       =   2250
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7770
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   15
      TabIndex        =   2
      Top             =   225
      Width           =   7755
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         Height          =   330
         Left            =   6600
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4515
         TabIndex        =   6
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   54263809
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1410
         TabIndex        =   4
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   54263809
         CurrentDate     =   38429
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date: "
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date: "
         Height          =   195
         Left            =   3315
         TabIndex        =   5
         Top             =   285
         Width           =   810
      End
   End
   Begin VB.Frame FramCont 
      Height          =   2955
      Left            =   0
      TabIndex        =   8
      Top             =   930
      Width           =   7740
      Begin MSComctlLib.ListView lstBill 
         Height          =   2625
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   4630
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "G.R.N "
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party"
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total Quantity"
            Object.Width           =   2223
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Net Amount"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unique"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "VTYP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SRNO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   45
      TabIndex        =   10
      Top             =   3870
      Width           =   7710
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
         Left            =   5070
         TabIndex        =   11
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
         Left            =   6405
         TabIndex        =   12
         Top             =   195
         Width           =   1035
      End
   End
   Begin VB.Label LBLDAYBOK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   3495
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label LBLDIVNAM 
      BackColor       =   &H00FFFFC0&
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
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "FRM_SPLISTIVR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

  FRM_TRNIVR.M_SRNO = Empty
  Unload Me
End Sub

Private Sub CMDOK_Click()
  Dim SEL_SRNO As String
  SEL_SRNO = lstBill.SelectedItem.SubItems(7)
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
  If FRM_TRNIVR.EFFGRN = "Y" Then
    SQL = "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='IVR' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D'"
    SQL1 = "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'  AND VTYP='IVR' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D'"
   Else
    SQL = "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'  AND VTYP='IVR' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D'"
    SQL1 = "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='IVR' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D'"
  End If
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL1, CN, adOpenDynamic, adLockOptimistic
  End If
  If EDTDAT.EOF Then
     lstBill.SetFocus
     Exit Sub
  End If
  With FRM_TRNIVR
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ACCMST WHERE CODE='" & EDTDAT!DRAC & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTDBAC = MSTDAT!Name & ""
     Else
      .TXTDBAC.Text = Empty
    End If
    
    
    
    
    
    
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM REFMST WHERE CODE='" & EDTDAT!TRCD & "' AND CATA='R'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTTRNM = MSTDAT!Name & ""
     Else
      .TXTTRNM = Empty
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
    .TXTGATN = EDTDAT!GATN & ""
    If IsNull(EDTDAT!GATD) Then
     Else
      .TXTGATD = EDTDAT!GATD & ""
    End If
    .TXTPONO = EDTDAT!PONO & ""
    If IsNull(EDTDAT!PODT) Then
     Else
      .TXTPODT = EDTDAT!PODT & ""
    End If
    .TXTTQTY = Format(EDTDAT!TQTY, "######.000")
    .TXTITOT = Format(EDTDAT!ITOT, "##########.00")
    .TXTBNET = Format(EDTDAT!BNET, "##########.00")
    Dim I As Double
    Dim J As Double
    I = 0
    For I = 0 To .flexBTRM.Rows - 1
      J = 0
      For J = 0 To EDTDAT.Fields.Count - 1
        If Trim(EDTDAT.Fields(J).Name) = Trim(.flexBTRM.TextMatrix(I, 0)) Then
            .flexBTRM.TextMatrix(I, 2) = Format(EDTDAT.Fields(J).Value, "#########.00")
        End If
        If Trim(EDTDAT.Fields(J).Name) = "PER" & Trim(.flexBTRM.TextMatrix(I, 0)) Then
           .flexBTRM.TextMatrix(I, 1) = Format(EDTDAT.Fields(J).Value, "######.00")
        End If
      Next
    Next
    .FLEX.Rows = 2
    If FRM_TRNIVR.EFFGRN = "Y" Then
      SQL = "SELECT PURTRAN.*,ITMMST.NAME FROM PURTRAN INNER JOIN ITMMST ON PURTRAN.ICOD=ITMMST.CODE WHERE PURTRAN.COMP='" & compPth & "' AND PURTRAN.VTYP='IVR' AND PURTRAN.SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D' ORDER BY PURTRAN.SRCH"
      SQL1 = "SELECT STORETRAN.*,ITMMST.NAME FROM STORETRAN INNER JOIN ITMMST ON STORETRAN.ICOD=ITMMST.CODE WHERE STORETRAN.COMP='" & compPth & "' AND STORETRAN.VTYP='IVR' AND STORETRAN.SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D' ORDER BY STORETRAN.SRCH"
     Else
      SQL = "SELECT STORETRAN.*,ITMMST.NAME FROM STORETRAN INNER JOIN ITMMST ON STORETRAN.ICOD=ITMMST.CODE WHERE STORETRAN.COMP='" & compPth & "' AND STORETRAN.VTYP='IVR' AND STORETRAN.SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D' ORDER BY STORETRAN.SRCH"
      SQL1 = "SELECT PURTRAN.*,ITMMST.NAME FROM PURTRAN INNER JOIN ITMMST ON PURTRAN.ICOD=ITMMST.CODE WHERE PURTRAN.COMP='" & compPth & "' AND PURTRAN.VTYP='IVR' AND PURTRAN.SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D' ORDER BY PURTRAN.SRCH"
    End If
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL, CN, adOpenForwardOnly, adLockOptimistic
    If EDTDAT.EOF Then
      If EDTDAT.State = 1 Then EDTDAT.Close
      EDTDAT.Open SQL1, CN, adOpenDynamic, adLockOptimistic
    End If
    I = 1
    Do While Not EDTDAT.EOF
     .FLEX.TextMatrix(I, 0) = EDTDAT!SRCH
     .FLEX.TextMatrix(I, 1) = EDTDAT!chln & ""
     If Not IsNull(EDTDAT!CHDT) Then
       .FLEX.TextMatrix(I, 2) = EDTDAT!CHDT
      Else
       .FLEX.TextMatrix(I, 2) = ""
     End If
     
     .FLEX.TextMatrix(I, 3) = EDTDAT!Name & ""
     .FLEX.TextMatrix(I, 4) = EDTDAT!MRGN & ""
     .FLEX.TextMatrix(I, 5) = EDTDAT!grad & ""
     .FLEX.TextMatrix(I, 6) = EDTDAT!COPS
     .FLEX.TextMatrix(I, 7) = EDTDAT!PCES
     .FLEX.TextMatrix(I, 8) = Format(EDTDAT!QNTY, "########.000")
     .FLEX.TextMatrix(I, 9) = Format(EDTDAT!RATE, "######.0000")
     .FLEX.TextMatrix(I, 10) = Format(EDTDAT!AMNT, "#########.00")
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
     
     If RS.State = 1 Then RS.Close
     RS.Open "select * from itmmst where CODE='" & EDTDAT!ICOD & "'", CN, adOpenDynamic, adLockOptimistic
     If Not IsNull(RS!EXTRA5) Then
       Select Case RS!EXTRA5
        Case "B"
         
         If RS.State = 1 Then RS.Close
         RS.Open "SELECT * FROM BEMREG WHERE COMP='" & compPth & "' AND VTYP='IVR' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D' ORDER BY SRCH", CN, adOpenDynamic, adLockOptimistic
         With BeamDTL
         Do While Not RS.EOF
            If IsNull(RS!MCNO) Or Trim(RS!MCNO) = "" Then
              'O.k
             Else
              FRM_TRNIVR.ALLOWEDITDEL = False
            End If
            .LBLITEM = EDTDAT!Name
            .FLX.Rows = .FLX.Rows + 1
            .FLX.TextMatrix(.FLX.Rows - 1, 0) = .FLX.Rows - 1
            .FLX.TextMatrix(.FLX.Rows - 1, 1) = RS!BMNO
            .FLX.TextMatrix(.FLX.Rows - 1, 2) = RS!END1
            .FLX.TextMatrix(.FLX.Rows - 1, 3) = RS!lnth
            .FLX.TextMatrix(.FLX.Rows - 1, 4) = RS!nwgt
            RS.MoveNext
         Loop
         End With
        Case "T"
         If RS.State = 1 Then RS.Close
         RS.Open "SELECT * FROM TAKREG WHERE COMP='" & compPth & "' AND VTYP='PPF' AND DBCD='" & FRM_TRNIVR.M_DBCD_DIRIVR & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND ICLN='" & FRM_TRNIVR.TXTVBNO & "' AND RECSTAT<>'D' ORDER BY SRCH", CN, adOpenDynamic, adLockOptimistic
         With Taka_Dtl
         Do While Not RS.EOF
            If IsNull(RS!DTYP) Or Trim(RS!DTYP) = "" Then
              'O.k
             Else
              FRM_TRNIVR.ALLOWEDITDEL = False
            End If
            .LBLITEM = EDTDAT!Name
            .FLX.Rows = .FLX.Rows + 1
            .FLX.TextMatrix(.FLX.Rows - 1, 0) = .FLX.Rows - 1
            .FLX.TextMatrix(.FLX.Rows - 1, 1) = RS!TAKN
            .FLX.TextMatrix(.FLX.Rows - 1, 2) = RS!MTRS
            .FLX.TextMatrix(.FLX.Rows - 1, 3) = RS!WGHT
            .FLX.TextMatrix(.FLX.Rows - 1, 4) = RS!AVGW
            RS.MoveNext
         Loop
         End With
        Case "O"
       End Select
     End If
     
     EDTDAT.MoveNext
     I = I + 1
     If Not EDTDAT.EOF Then
       If I > 1 Then
           .FLEX.Rows = .FLEX.Rows + 1
       End If
     End If
    Loop
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
  LBLDIVNAM.Caption = DIVNAM
  LBLDAYBOK.Caption = FRM_TRNIVR.Caption
  txtFrDate = GetMinDate
  txtToDate = GetMaxDate
  Me.KeyPreview = True
  CMDOK.Enabled = False
  cmdCancel.Enabled = True
End Sub
Private Sub CMDGO_Click()
  lstBill.ListItems.Clear
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Dim SQL As String
  Dim SQL1 As String
  SQL = Empty
  SQL1 = Empty
  If FRM_TRNIVR.EFFGRN = "Y" Then
    SQL = "SELECT DISTINCT GRN.*,ACCMST.NAME FROM GRN INNER JOIN ACCMST ON GRN.PCOD=ACCMST.CODE INNER JOIN PURTRAN ON GRN.COMP=PURTRAN.COMP AND GRN.VTYP=PURTRAN.VTYP AND GRN.SRNO=PURTRAN.SRNO WHERE GRN.COMP='" & compPth & "' AND GRN.VTYP='IVR' AND GRN.DVCD='" & DIVCOD & "' AND GRN.DBCD='" & FRM_TRNIVR.M_DBCD_DIRIVR & "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "' AND GRN.RECSTAT<>'D' AND GRN.UNIT='" & UNCD & "' ORDER BY GRN.DATE,GRN.VBNO"
    SQL1 = "SELECT DISTINCT GRN.*,ACCMST.NAME FROM GRN INNER JOIN ACCMST ON GRN.PCOD=ACCMST.CODE INNER JOIN STORETRAN ON GRN.COMP=STORETRAN.COMP AND GRN.VTYP=STORETRAN.VTYP AND GRN.SRNO=STORETRAN.SRNO WHERE GRN.COMP='" & compPth & "' AND GRN.VTYP='IVR' AND GRN.DVCD='" & DIVCOD & "' AND GRN.DBCD='" & FRM_TRNIVR.M_DBCD_DIRIVR & "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "' AND GRN.RECSTAT<>'D'  AND GRN.UNIT='" & UNCD & "' ORDER BY GRN.DATE,GRN.VBNO"
   Else
    SQL = "SELECT DISTINCT GRN.*,ACCMST.NAME FROM GRN INNER JOIN ACCMST ON GRN.PCOD=ACCMST.CODE INNER JOIN STORETRAN ON GRN.COMP=STORETRAN.COMP AND GRN.VTYP=STORETRAN.VTYP AND GRN.SRNO=STORETRAN.SRNO WHERE GRN.COMP='" & compPth & "' AND GRN.VTYP='IVR' AND GRN.DVCD='" & DIVCOD & "' AND GRN.DBCD='" & FRM_TRNIVR.M_DBCD_DIRIVR & "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "' AND GRN.RECSTAT<>'D'  AND GRN.UNIT='" & UNCD & "' ORDER BY GRN.DATE,GRN.VBNO"
    SQL1 = "SELECT DISTINCT GRN.*,ACCMST.NAME FROM GRN INNER JOIN ACCMST ON GRN.PCOD=ACCMST.CODE INNER JOIN PURTRAN ON GRN.COMP=PURTRAN.COMP AND GRN.VTYP=PURTRAN.VTYP AND GRN.SRNO=PURTRAN.SRNO WHERE GRN.COMP='" & compPth & "' AND GRN.VTYP='IVR' AND GRN.DVCD='" & DIVCOD & "' AND GRN.DBCD='" & FRM_TRNIVR.M_DBCD_DIRIVR & "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "' AND GRN.RECSTAT<>'D' AND GRN.UNIT='" & UNCD & "' ORDER BY GRN.DATE,GRN.VBNO"
  End If
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL1, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
      MsgBox "No Record found for given criteria ", vbInformation
      txtToDate.SetFocus
      Exit Sub
    End If
  End If
  Do While Not EDTDAT.EOF
   Set lstitem = lstBill.ListItems.Add
   lstitem.Text = Format(EDTDAT![Date], "dd/MM/yyyy")
   lstitem.SubItems(1) = EDTDAT![VBNO]
   lstitem.SubItems(2) = EDTDAT![Name]
   lstitem.SubItems(3) = Format(EDTDAT!TQTY, "########.000")
   lstitem.SubItems(4) = Format(EDTDAT!BNET, "#############.00")
   lstitem.SubItems(6) = EDTDAT!VTYP
   lstitem.SubItems(7) = EDTDAT!SRNO
   EDTDAT.MoveNext
  Loop
  CMDOK.Enabled = True
  CMDOK.Default = True
  If FRM_SPLISTIVR.Visible = True Then lstBill.SetFocus
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

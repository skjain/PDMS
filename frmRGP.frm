VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRGP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gate Pass For Returnable Pallets  "
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7635
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   255
         Width           =   5805
      End
      Begin MSComCtl2.DTPicker dtOPDT 
         Height          =   330
         Left            =   1230
         TabIndex        =   4
         Top             =   720
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
         Format          =   56426497
         CurrentDate     =   39343
      End
      Begin MSComCtl2.DTPicker dtENDT 
         Height          =   330
         Left            =   4320
         TabIndex        =   6
         Top             =   720
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
         Format          =   56426497
         CurrentDate     =   39343
      End
      Begin WelchButton.lvButtons_H cmdSearch 
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmRGP.frx":0000
         cBack           =   -2147483633
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
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "To Date :"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "From Date :"
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
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3900
      Left            =   120
      TabIndex        =   15
      Top             =   1275
      Width           =   7455
      Begin MSComctlLib.ListView lstVoucher 
         Height          =   3570
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6297
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Gatepass"
            Object.Width           =   2311
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1958
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   5998
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Total . Pallets"
            Object.Width           =   2011
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   7455
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
         TabIndex        =   12
         Text            =   "100"
         Top             =   270
         Width           =   495
      End
      Begin VB.ComboBox cmbfmt 
         Height          =   315
         ItemData        =   "frmRGP.frx":039A
         Left            =   2880
         List            =   "frmRGP.frx":03A4
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   6720
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Image           =   "frmRGP.frx":03C4
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Image           =   "frmRGP.frx":0816
         cBack           =   -2147483633
      End
      Begin VB.Label Label13 
         Caption         =   "Report Zoom %"
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
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Format :"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_DBCD As String
Dim SER_VBN As String
Dim M_GPN As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
Dim I As Long
Dim GPNO As String

    CRPT.Reset
    crptConnect CRPT
        
    If M_COMPBILL = "CHK" Then
       ReportName = App.PATH & "\Reports\ReturnablePalletsGP_CHK.rpt"
    Else
       ReportName = App.PATH & "\Reports\ReturnablePalletsGatePass.rpt"
    End If
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
  
    If lstVoucher.ListItems.COUNT < 1 Then
        MsgBox "No Gatepass Found To Print !!", vbInformation
        Exit Sub
    End If
 
    'CHALLAN ARRAY
     M_GPN = Empty
     For I = 1 To lstVoucher.ListItems.COUNT
          If lstVoucher.ListItems(I).Checked = True Then
             If M_GPN <> Empty Then M_GPN = M_GPN & ","
              M_GPN = M_GPN & "'" & Trim(lstVoucher.ListItems(I)) & "'"
          End If
     Next
      
      If M_GPN = Empty Then
            MsgBox "No Item Selected !!", vbInformation, "No Information Found !!"
            Exit Sub
      End If
      
      rptsql = "{PKGSTK.COMP}='" & compPth & "' AND {PKGSTK.UNIT}='" & txtUNIT.Tag & _
               "'  And {PKGSTK.VTYP}='RET' AND {PKGSTK.RECSTAT} <>'D' " & _
               "AND {PKGSTK.CHLN} IN [" & M_GPN & "]"
    
    CRPT.ReportFileName = ReportName
    
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
       
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
         'txtUNIT.SetFocus
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
On Error GoTo errFillLst
Dim SQL As String
Dim Item As ListItem
    
If txtUNIT = Empty Then txtUNIT.SetFocus: Exit Sub

SQL = "SELECT PKGSTK.DATE AS DATE,PKGSTK.CHLN AS VBNO,ACCMST.NAME AS PARTY," & _
"PKGSTK.PALLETS AS TQTY FROM PKGSTK " & _
"INNER JOIN ACCMST ON PKGSTK.PCOD=ACCMST.CODE " & _
"WHERE PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & UNCD & _
"' AND PKGSTK.DATE >='" & Format(dtOPDT.Value, "mm/dd/yyyy") & _
"' AND PKGSTK.DATE <='" & Format(dtENDT.Value, "mm/dd/yyyy") & "' AND RECSTAT<>'D' AND VTYP='RET' AND OPER='-'"
    
    Set rsTemp = New Recordset
    rsTemp.Open SQL, CN
    
    lstVoucher.ListItems.Clear
    Screen.MousePointer = vbHourglass
    
    Do While Not rsTemp.EOF
        Set Item = lstVoucher.ListItems.ADD
        Item.Text = Trim(rsTemp!VBNO)
        
        Item.SubItems(1) = IIf(IsNull(rsTemp!Date), Date, rsTemp!Date)
        Item.SubItems(2) = Trim(rsTemp!PARTY & "")
        Item.SubItems(3) = rsTemp!TQTY
        
        rsTemp.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    rsTemp.Close
    Exit Sub
    
errFillLst:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If txtDVCD = Empty And ActiveControl.NAME = "txtDVCD" And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = 13 Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub Form_Load()
    Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtOPDT = GetMinDate
    dtENDT = Now
End Sub

Private Sub lstVoucher_GotFocus()
lstVoucher.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstVoucher_LostFocus()
 lstVoucher.BackColor = vbWhite
End Sub

Private Sub txtDVCD_GotFocus()
txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("SELECT TOP 20 Code,NAME From DIVMST Where COMP='" & compPth & "' and Unit='" & txtUNIT.Tag & "' AND RECSTAT='A'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
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
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("SELECT TOP 20 Code,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If
End Sub

Private Sub SER_CARTLIST()
  Dim SQL As String
  Dim MSTDAT As New ADODB.Recordset
  If MSTDAT.State = 1 Then MSTDAT.Close
  Dim PTYNM As String
  SQL = "SELECT BOXN,BOXREG.LTNO,ITMMST.NAME AS ITNM,REFMST.NAME AS BRNM,BOXREG.GRAD,LOCMST.NAME AS LOCA,ACCMST.NAME AS ACNM,boxreg.nwgt " & _
      " FROM BOXREG INNER JOIN ITMMST ON BOXREG.ICOD=ITMMST.CODE " & _
      " LEFT JOIN LOCMST ON LOCMST.CODE=BOXREG.LCOD " & _
      " LEFT JOIN REFMST ON REFMST.CODE=BOXREG.BRCD " & _
      " INNER JOIN ACCMST ON ACCMST.CODE=BOXREG.DCOD " & _
      " WHERE BOXREG.COMP='" & compPth & "' AND BOXREG.UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "' " & _
      " AND DDBC='" & M_DBCD & "' " & _
      " AND DCLN IN " + "(" + SER_VBN + ")" & _
      " ORDER BY ACCMST.NAME,LOCMST.NAME,BOXREG.BOXN"
  If RS.State = 1 Then RS.Close
  RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  Dim ROW As Double
  Dim PGL As Double
  ROW = 0
  PGL = 55
  Print #1, DWA + compNm + DWI
  Print #1, "-------------------------------------------------------------------------------"
  Print #1, "Sr..  Box No....  Location  Denier..............  Grade  Colour....  Net Weight"
  Print #1, "-------------------------------------------------------------------------------"
  ROW = ROW + 4
  Do While Not RS.EOF
   PTYNM = RS!ACNM
   Print #1, "Party : " + Mid(PTYNM + Space(40), 1, 40) + " Agent : " + Mid(RS!BRNM + Space(20), 1, 20)
   Print #1, "-------------------------------------------------------------------------------"
   ROW = ROW + 1
   Dim CNTR As Double
   Dim PTYSTG As String
   PTYSTG = Empty
   CNTR = 0
   Dim CLCD As String
   Dim CLNM As String
   CLNM = Space(10)
   Do While Not RS.EOF And RS!ACNM = PTYNM
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT SHCD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "' AND LTNO='" & RS!ltno & "'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       CLCD = MSTDAT!SHCD & ""
       If MSTDAT.State = 1 Then MSTDAT.Close
       MSTDAT.Open "SELECT * FROM REFMST WHERE CODE='" & CLCD & "'", CN, adOpenDynamic, adLockOptimistic
       If MSTDAT.EOF = False Then
         CLNM = MSTDAT!NAME & ""
       End If
     End If
     PTYSTG = Empty
     CNTR = CNTR + 1
     PTYSTG = PTYSTG + nstr(CNTR, 4, 0) + "  " + Mid(RS!BOXN + Space(10), 1, 10) + "  " + Mid(RS!LOCA & "" + Space(10), 1, 10)
     PTYSTG = PTYSTG + Mid(RS!ITNM + Space(20), 1, 20) + "  " + Mid(RS!grad + Space(5), 1, 5) + "  " + Mid(CLNM + Space(10), 1, 10) + "  " + nstr(RS!nwgt, 8, 3)
     Print #1, PTYSTG
     
     ROW = ROW + 1
     If ROW > PGL Then
       Print #1, Chr(12)
       Print #1, DWA + compNm + DWI
       Print #1, "-------------------------------------------------------------------------------"
       Print #1, "Sr..  Box No....  Location  Denier..............  Grade  Colour....  Net Weight"
       Print #1, "-------------------------------------------------------------------------------"
       ROW = 0
     End If
     RS.MoveNext
     If RS.EOF Then Exit Do
   Loop
   Print #1, "-------------------------------------------------------------------------------"
   ROW = ROW + 1
   
   If RS.EOF Then Exit Do
   If ROW > PGL Then
       Print #1, Chr(12)
       Print #1, DWA + compNm + DWI
       Print #1, "-------------------------------------------------------------------------------"
       Print #1, "Sr..  Box No....  Location  Denier..............  Grade  Colour....  Net Weight"
       Print #1, "-------------------------------------------------------------------------------"
       ROW = 0
   End If
 Loop
 
 
 Print #1, Chr(12)
End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub

Public Sub LoadingStatement_GEN(Optional unit As String, Optional GPNO As String)
'*******************************************************************************************
'Module Name : Loading Statement
'Company Name: General
'Developed Date : 30-12-2010
'********************************************************************************************

On Error GoTo errChallanPrint
Dim COUNTER As Long
COUNTER = 0
Dim spc_cmp As String
Dim spc_adr As String
Dim spc_tel As String

Dim SQLDATA As String
Dim RSNWDATA As New ADODB.Recordset
Dim RSDATA As Recordset
Dim RSSUBDATA As Recordset
Dim prt_stg As String
Dim PAGENUMBER As Long
Dim ROW_CTR As Long
Dim PAGELENGTH As Long

PAGELENGTH = 58
'Setting Up Company and Unit Detail
 Call CollectData(unit)
     
    spc_cmp = Space((79 - Len(compNm) * 2) / 2)
    spc_adr = Space((79 - Len(Mid(CMP_FADD, 1, 79))) / 2)
    spc_tel = Space((79 - Len(cmp_tel)) / 2)
    
    Dim RSNEW As New ADODB.Recordset
    Set RSNEW = New ADODB.Recordset
        
    SQLDATA = "SELECT *,FINITMMST.NAME AS DENIER,SHADEMST.NAME AS GRADE,TRANSPORTMST.NAME AS TRANSNM,SPTRAN.QNTY,SPTRAN.VBNO,GPMST.GPDT FROM SPTRAN " & _
              "INNER JOIN FINITMMST ON FINITMMST.COMP = SPTRAN.COMP AND FINITMMST.UNIT = SPTRAN.UNIT AND " & _
              "FINITMMST.DVCD = SPTRAN.DVCD AND FINITMMST.CODE = SPTRAN.ICOD " & _
              "INNER JOIN SHADEMST ON SHADEMST.CODE = SPTRAN.GRAD " & _
              "INNER JOIN TRANSPORTMST ON TRANSPORTMST.CODE = SPTRAN.TRCD " & _
              " INNER JOIN  GPMST ON GPMST.COMP = SPTRAN.COMP AND GPMST.UNIT = SPTRAN.UNIT AND GPMST.GPNO = SPTRAN.GATEPASSNO " & _
              "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & unit & _
              "' AND GATEPASSNO IN (" & GPNO & ") AND SPTRAN.RECSTAT<>'D'"
        
    If RS.State = adStateOpen Then RS.Close
    RS.Open SQLDATA, CN, adOpenKeyset, adLockPessimistic
    Do While Not RS.EOF
      COUNTER = COUNTER + 1
      If COUNTER = 1 Then
       'Call Header
       Print #1, CWA
       Print #1,
       Print #1, spc_cmp + DWA + DCA + compNm + DCI + DWI
       Print #1,
       Print #1, Space((79 - Len("GATE - PASS") * 2) / 2) + DWA + "GATE - PASS" + DWI
       Print #1, Space(1) + Replicate("-", 75)
       Print #1, Space(1) + Left(Trim(RS!TRANSNM & "") & Space(35), 35)
       Print #1, Space(49) + "Gate Pass No. : " + GPNO
       Print #1, Space(49) + "Date          : " + CStr(RS!GPDT)
       Print #1, Replicate("-", 75)
       Print #1, Space(1) + "Item Description ..." + Space(10) + "Shade" + Space(11) + "Qty. " + Space(2) + "Net Wt." + Space(7) + "Ch. No"
       Print #1, Replicate("-", 75)
   
     '  ROW_CTR = ROW_CTR + 6
      ' PAGENUMBER = 1
      End If
      Print #1, Space(1) + Left(Trim(RS!DENIER & "") & Space(25), 25) + Space(3) + Left(Trim(RS!GRADE & "") & Space(10), 10) + Space(7) + nstr(RS!PCES, 4, 0) + Space(2) + nstr(RS!QNTY, 9, 3) + Space(4) + RS!VBNO
      'Print #1, Replicate("-", 75)
      ROW_CTR = ROW_CTR + 1
      
      Dim TTL_QTY As Double: TTL_QTY = 0
      TTL_QTY = TTL_QTY + Val(RS!QNTY)
         
         RS.MoveNext
        Loop
        Do While ROW_CTR < 16
            Print #1,
            ROW_CTR = ROW_CTR + 1
        Loop
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, CWI
        RS.Close
        
 Exit Sub
errChallanPrint:
    Close #1
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show
 End Sub




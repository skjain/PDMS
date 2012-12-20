VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_ORDLIST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order List"
   ClientHeight    =   6675
   ClientLeft      =   405
   ClientTop       =   1440
   ClientWidth     =   11355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11355
   Begin VB.Frame FRM3 
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   11175
      Begin WelchButton.lvButtons_H CMDOK 
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&O.k"
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
         Image           =   "frm_ORDLIST.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDCLOSE 
         Height          =   375
         Left            =   8040
         TabIndex        =   15
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Close"
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
         Image           =   "frm_ORDLIST.frx":059A
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame FRM2 
      Height          =   4575
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   11175
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   4215
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7435
         _Version        =   393216
         Cols            =   8
         BackColor       =   -2147483634
         ForeColorFixed  =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FRM1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.TextBox M_ORDN 
         Height          =   285
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox M_PNAM 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox M_INAM 
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker m_stdt 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16384001
         CurrentDate     =   39340
      End
      Begin MSComCtl2.DTPicker m_eddt 
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16384001
         CurrentDate     =   39340
      End
      Begin WelchButton.lvButtons_H BTNSEARCH 
         Height          =   375
         Left            =   9120
         TabIndex        =   11
         Top             =   600
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
         Image           =   "frm_ORDLIST.frx":0B34
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         Caption         =   "Order No."
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
         Left            =   5640
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name of A/c Party"
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
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Name of Item"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "To Date"
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
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "From Date"
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
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_ORDLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ORDERNO As String
Dim ORDBOK As String
Dim ORDDBC As String
Dim ORDUNITCOD As String
Public EXPORTREQ As String

Private Sub btnSearch_Click()
On Error GoTo errSearch
Dim Ctrl As Control

    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
    
    ORDERNO = Empty
        
    Dim SQL As String
    Dim M_PCOD As String
    Dim M_ICOD As String
    
    SQL = Empty
    
    SQL = "select * from ordMAN where COMP='" & compPth & "' AND DBCD='" & frm_ORDERBOOK.ORDDBCD & _
          "' AND ORDT>='" & Format(m_stdt.Value, "MM/DD/YYYY") & _
          "' AND ORDT<='" & Format(m_eddt.Value, "MM/DD/YYYY") & "' AND DCOD='" & frm_ORDERBOOK.DIVCODE & _
          "' AND RECSTAT='A' "
          
    If Not EXP_REQ Then
       SQL = SQL & " AND FIN_USER IS Null "
    Else
       'SQL = SQL & " AND ORDMAN.ORDN NOT IN (SELECT DISTINCT ORDN FROM ORDMAN where COMP='" & compPth & _
                   "' AND DBCD='" & frm_ORDERBOOK.ORDDBCD & "' AND DCOD='" & DIVCOD & _
                   "' AND RECSTAT='A' AND DISPATCHQTY > 0) "
    End If
        
    If Not M_PNAM = Empty Then
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM ACCMST WHERE NAME='" & M_PNAM.Text & "'", CN, adOpenKeyset, adLockPessimistic
        
        If Not RS.EOF Then
            M_PCOD = RS!CODE
        Else
            M_PCOD = Empty
        End If
    End If
    
    If Not M_INAM = Empty Then
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & ORDUNITCOD & _
                "' AND DVCD='" & frm_ORDERBOOK.DIVCODE & "' AND NAME='" & M_INAM.Text & "'", CN, adOpenKeyset, adLockPessimistic
        If Not RS.EOF Then
            M_ICOD = RS!CODE
        Else
            M_ICOD = Empty
        End If
    End If
    
    If Not M_PCOD = Empty Then
        SQL = SQL & " AND PCOD='" & M_PCOD & "'"
    End If
    
    If Not M_ICOD = Empty Then
        SQL = SQL & " AND ICOD='" & M_ICOD & "'"
    End If
           
    If Not Trim(M_ORDN) = Empty Then
        SQL = SQL & " AND ORDN='" & M_ORDN & "'"
    End If
    
    SQL = SQL & " AND  RECSTAT<>'D'"
    
    Dim MSTRS As New ADODB.Recordset
    
    SQL = SQL & " ORDER BY ORDN "
    
    If RS.State = 1 Then RS.Close
    RS.Open SQL, CN, adOpenKeyset, adLockPessimistic
    
    If RS.EOF Then
        MsgBox "Record Not Found For a Period", vbCritical
        m_stdt.SetFocus
        Exit Sub
    End If
    
    Dim I As Double
    I = 0
    Call HEDFLEX
    'Loop through Order Items
    Do While Not RS.EOF
        I = I + 1
    
        If I > FLEX.Rows - 1 Then
            FLEX.Rows = FLEX.Rows + 1
        End If
    
        If MSTRS.State = 1 Then MSTRS.Close
    
        MSTRS.Open "SELECT * FROM ACCMST WHERE CODE='" & RS!PCOD & "'", CN, adOpenKeyset, adLockPessimistic
    
        FLEX.TextMatrix(I, 0) = RS!ORDN
    
        If Not MSTRS.EOF Then
            FLEX.TextMatrix(I, 1) = Trim(MSTRS!NAME)
        Else
            FLEX.TextMatrix(I, 1) = ""
        End If
    
        If MSTRS.State = 1 Then MSTRS.Close
    
        MSTRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & ORDUNITCOD & _
                   "' AND DVCD='" & frm_ORDERBOOK.DIVCODE & "' AND CODE='" & RS!ICOD & "'", CN, adOpenKeyset, adLockPessimistic

        If Not MSTRS.EOF Then
            FLEX.TextMatrix(I, 2) = Trim(MSTRS!NAME & "")
        Else
            FLEX.TextMatrix(I, 2) = ""
        End If

        MSTRS.Close
        
        If MSTRS.State = 1 Then MSTRS.Close
    
        MSTRS.Open "SELECT * FROM GRDMST WHERE CODE='" & Trim(RS!TRCD & "") & "'", CN, adOpenKeyset, adLockPessimistic

        If Not MSTRS.EOF Then
            FLEX.TextMatrix(I, 3) = Trim(MSTRS!grad & "")
        Else
            FLEX.TextMatrix(I, 3) = ""
        End If

        MSTRS.Close
        
        FLEX.TextMatrix(I, 4) = Trim(nstr(RS!QNTY, 12, 3))
        FLEX.TextMatrix(I, 5) = Trim(nstr(RS!ARAT, 10, 3))
        FLEX.TextMatrix(I, 6) = Trim(nstr(RS!RATE, 10, 2))
        FLEX.TextMatrix(I, 7) = Trim(RS!RMRK & "")
                
        RS.MoveNext
    Loop
        
    RS.Close
    
    FRM2.Enabled = True
    CMDOK.Enabled = True
    If FLEX.Rows >= 2 And FLEX.TextMatrix(1, 1) <> "" Then FLEX.SetFocus
    
    Exit Sub
    
errSearch:
Resume
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdClose_Click()
  frm_ORDERBOOK.cmdCancel_Click
  frm_ORDERBOOK.cmdAdd.Enabled = False
  frm_ORDERBOOK.cmdEdit.Enabled = False
  frm_ORDERBOOK.cmdDelete.Enabled = False
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim SQL As String
  
  If ORDERNO = Empty Then
    MsgBox "Please Select From Order List", vbInformation
    FLEX.SetFocus
    Exit Sub
  End If
  
  SQL = Empty
  SQL = "SELECT * FROM ORDMAN WHERE COMP='" & compPth & "' AND ORDN='" & ORDERNO & _
        "'  AND RECSTAT<>'D' "
  
      
  If Not EXP_REQ Then
     
   'IsAllowForEdit
    Dim ALLOWRS As ADODB.Recordset
    Set ALLOWRS = New ADODB.Recordset
    
    If ALLOWRS.State = 1 Then ALLOWRS.Close
    ALLOWRS.Open SQL & " AND FIN_USER IS NOT Null ", CN, adOpenDynamic, adLockOptimistic
    If Not ALLOWRS.EOF Then
       MsgBox "Partial Order Cann't Be Edit.", vbCritical, "Approval Exist"
       Exit Sub
    End If
    '========================================================================================
     
     SQL = SQL & " AND FIN_USER IS Null "
  Else
  
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT ISNULL(SUM(QNTY),0) AS ORDQTY FROM ORDMAN WHERE COMP='" & compPth & _
             "' AND ORDN='" & ORDERNO & "' AND RECSTAT<>'D' ", CN, adOpenDynamic, adLockOptimistic
     If Not RS.EOF Then
        frm_ORDERBOOK.EDIT_ORDQTY = Val(RS!ORDQTY)
     End If
     RS.Close
     
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open SQL, CN, adOpenKeyset, adLockPessimistic
  
  If RS.EOF Then
    m_stdt.SetFocus
    Exit Sub
  End If
  
  Dim M_PCOD As String
  Dim M_ICOD As String
  Dim M_BRCD As String
  Dim mstrst As New ADODB.Recordset
  Dim ORDER_NO As String
  With frm_ORDERBOOK
    .M_ORDN = RS!ORDN
    ORDER_NO = RS!ORDN
    If mstrst.State = 1 Then mstrst.Close
    mstrst.Open "SELECT * FROM ACCMST WHERE CODE='" & RS!PCOD & "'", CN, adOpenKeyset, adLockPessimistic
    If Not mstrst.EOF Then
      .M_PNAM = mstrst!NAME
      .M_PCOD = RS!PCOD
    End If
    
    If mstrst.State = 1 Then mstrst.Close
    mstrst.Open "SELECT * FROM REFMST WHERE CODE='" & RS!BRCD & "'", CN, adOpenKeyset, adLockPessimistic
    If Not mstrst.EOF Then
      .M_BRNM = mstrst!NAME
    End If
    
    .TXTFREIGHT = Val(Trim(RS!FREIGHT_PERKG & ""))
    
    If mstrst.State = 1 Then mstrst.Close
    mstrst.Open "SELECT * FROM TAXMST WHERE CODE='" & RS!TXCD & "'", CN, adOpenKeyset, adLockPessimistic
    If Not mstrst.EOF Then
      .M_TXNM = mstrst!NAME
    End If
               
        
    .M_ORDT = RS!ORDT
    .M_CRDS = RS!CRDS
    .M_RMRK = RS!ORDRMRK & ""
    .M_PORD = RS!PORD & ""
    
    .TXTGRAD = GetCode("GRDMST", Trim(RS!TRCD) & "", "CODE", "GRAD")
    .ITMFLEX.Clear
    .ITMFLEX.ColWidth(0) = 400
    .ITMFLEX.ColWidth(1) = 1700
    .ITMFLEX.ColWidth(2) = 1900
    .ITMFLEX.ColWidth(3) = 1250
    .ITMFLEX.ColWidth(4) = 1000
    .ITMFLEX.ColWidth(5) = 1250
    .ITMFLEX.ColWidth(6) = 1600
    .ITMFLEX.ColWidth(7) = 0
    
    .ITMFLEX.Clear
    .ITMFLEX.TextMatrix(0, 0) = "Sr."
    .ITMFLEX.TextMatrix(0, 1) = "Item Description"
    .ITMFLEX.TextMatrix(0, 2) = "Grade"
    .ITMFLEX.TextMatrix(0, 3) = "Quantity"
    .ITMFLEX.TextMatrix(0, 4) = "Ass.Rate"
    .ITMFLEX.TextMatrix(0, 5) = "Net Rate"
    .ITMFLEX.TextMatrix(0, 6) = "Remarks"
    .ITMFLEX.TextMatrix(0, 7) = "UsedQty" 'Specially for export editing
  
    Dim I As Double
    I = 0
    Do While Not RS.EOF
     I = I + 1
     If I >= .ITMFLEX.Rows - 1 Then
       .ITMFLEX.Rows = .ITMFLEX.Rows + 1
     End If
     
     If mstrst.State = 1 Then mstrst.Close
     mstrst.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & ORDUNITCOD & _
                "' AND DVCD='" & frm_ORDERBOOK.DIVCODE & "' AND CODE='" & RS!ICOD & "'", CN, adOpenKeyset, adLockPessimistic
     .ITMFLEX.TextMatrix(I, 0) = Trim(RS!OSRC)
     .ITMFLEX.TextMatrix(I, 1) = Trim(mstrst!NAME)
     .ITMFLEX.TextMatrix(I, 2) = GetCode("GRDMST", Trim(RS!TRCD) & "", "CODE", "GRAD")
     .ITMFLEX.TextMatrix(I, 3) = Trim(nstr(RS!QNTY, 12, 3))
     .ITMFLEX.TextMatrix(I, 4) = Trim(nstr(RS!ARAT, 10, 3))
     .ITMFLEX.TextMatrix(I, 5) = Trim(nstr(RS!RATE, 10, 3))
     .ITMFLEX.TextMatrix(I, 6) = Trim(RS!RMRK & "")
     .ITMFLEX.TextMatrix(I, 7) = Val(RS!DOQTY & "") + Val(RS!DISPATCHQTY & "") + Val(RS!CANCELQTY & "")
     
     RS.MoveNext
     If .ITMFLEX.Rows > 6 Then .ITMFLEX.TopRow = .ITMFLEX.TopRow + 2
    Loop
    .ITMFLEX.Rows = .ITMFLEX.Rows - 1
    .btn_sts (True)
    .Frm1.Enabled = True
    .FRM2.Enabled = True
    .FRM3.Enabled = True
    .M_SRCH.Enabled = False
    .cmdCancel.Cancel = True
    
    If EXPORTREQ = "Y" Then
       Call EXPORT_DETAIL
       .cmdExport.Enabled = True
    Else
      .cmdExport.Enabled = False
    End If
    
  End With
  
  If EXPORTREQ = "Y" Then Unload Me: Exit Sub
  
  Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ORDTRN WHERE ORDN='" & ORDER_NO & "' AND COMP='" & compPth & "' AND DBCD='" & frm_ORDERBOOK.ORDDBCD & "'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
     Unload Me
     With frm_ORDERBOOK
        .cmdSave.Enabled = False
     End With
     MsgBox "Further Transaction Exists. Record Can Not Be Edited.", vbCritical
     Exit Sub
  End If
  
  Unload Me
  
End Sub

Private Sub Flex_Click()
  ORDERNO = FLEX.TextMatrix(FLEX.ROW, 0)
End Sub

Private Sub FLEX_EnterCell()
 FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub FLEX_GotFocus()
    CMDOK.Default = False
    If FLEX.TextMatrix(1, 1) <> "" Then Call Flex_Click
End Sub

Private Sub Flex_LeaveCell()
 FLEX.CellBackColor = vbWhite
End Sub

Private Sub FLEX_RowColChange()
    ORDERNO = FLEX.TextMatrix(FLEX.ROW, 0)
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
  
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
  Me.Caption = "Order List : "
  m_stdt.MinDate = FSDT: m_stdt.MaxDate = FEDT
  m_eddt.MinDate = FSDT: m_eddt.MaxDate = FEDT
  
  m_stdt.Value = GetMinDate
  m_eddt.Value = GetMaxDate
  FRM2.Enabled = False
  CMDOK.Enabled = False
  Call HEDFLEX
  Me.KeyPreview = True
  ORDERNO = Empty
  
  Me.Caption = "LIST OF ORDER Booked By : " + frm_ORDERBOOK.ORDBOK
  ORDUNITCOD = frm_ORDERBOOK.FindUnit
End Sub
Private Sub HEDFLEX()
  FLEX.Clear
  FLEX.Rows = 2
  FLEX.ColWidth(0) = 1200
  FLEX.ColWidth(1) = 1700
  FLEX.ColWidth(2) = 1900
  FLEX.ColWidth(3) = 1200
  FLEX.ColWidth(4) = 1000
  FLEX.ColWidth(5) = 1000
  FLEX.ColWidth(6) = 1000
  FLEX.ColWidth(7) = 1300
    
  
  FLEX.Clear
  FLEX.TextMatrix(0, 0) = "OrderNo."
  FLEX.TextMatrix(0, 1) = "Name of Party"
  FLEX.TextMatrix(0, 2) = "Item Description"
  FLEX.TextMatrix(0, 3) = "Grade/Shade"
  FLEX.TextMatrix(0, 4) = "Quantity"
  FLEX.TextMatrix(0, 5) = "Ass.Rate"
  FLEX.TextMatrix(0, 6) = "Net Rate"
  FLEX.TextMatrix(0, 7) = "Remarks"

End Sub

Private Sub m_eddt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub M_INAM_GotFocus()
M_INAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_INAM_LostFocus()
 M_INAM.BackColor = vbWhite
End Sub

Private Sub M_ORDN_GotFocus()
 M_ORDN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_ORDN_LostFocus()
 M_ORDN.BackColor = vbWhite
End Sub

Private Sub M_PNAM_GotFocus()
M_PNAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_PNAM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        M_PNAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST", 0, M_PNAM.Text, "SELECT A/C PARTY")
        
    End If
    Me.KeyPreview = True
End Sub
Private Sub M_INAM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        M_INAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM FINITMMST WHERE COMP='" & compPth & _
                                  "' AND UNIT='" & ORDUNITCOD & "' AND DVCD='" & frm_ORDERBOOK.DIVCODE & _
                                  "'", 0, M_INAM.Text, "SELECT ITEM FROM LIST")
    End If
    Me.KeyPreview = True
End Sub

Private Sub M_PNAM_LostFocus()
M_PNAM.BackColor = vbWhite
End Sub

Private Sub m_stdt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub EXPORT_DETAIL()
With frm_ORDERBOOK
.cmdExport.Enabled = True
 If RS.State = 1 Then RS.Close
 RS.Open "SELECT * FROM EXPORD WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & frm_ORDERBOOK.ORDDBCD & "' AND ORDN='" & .M_ORDN & "'", CN, adOpenDynamic
 If Not RS.EOF Then
    FRM_TRNEXPORD.TXTEXPORTREF = RS!EXPORTREFNO
    FRM_TRNEXPORD.TXTCNTRYOFORIGIN = RS!CNTRYOFORGIN
    FRM_TRNEXPORD.TXTCNTRYFNLDES = RS!CNTRYOFFINALDES
    FRM_TRNEXPORD.TXTTERMS = RS!TRMSOFDLRY
    FRM_TRNEXPORD.TXTPAYMENT = RS!TRMSOFPYMT
    FRM_TRNEXPORD.TXTPRECARIAGE = RS!PRECARIGBY
    FRM_TRNEXPORD.TXTPLACEOFRCPT = RS!PLACEOFRCPT
    FRM_TRNEXPORD.TXTVSLNO = RS!VSLFLTNO
    FRM_TRNEXPORD.TXTPORTOFLOD = RS!PORTOFLOAD
    FRM_TRNEXPORD.TXTPORTOFDIS = RS!PORTOFDISCHARG
    FRM_TRNEXPORD.TXTFNLDES = RS!FINALDEST
    FRM_TRNEXPORD.TXTREMARK1 = RS!REMARK1
    FRM_TRNEXPORD.TXTREMARK2 = RS!REMARK2
    FRM_TRNEXPORD.TXTREMARK3 = RS!REMARK3
    FRM_TRNEXPORD.TXTMARKS = RS!MARKNO
    FRM_TRNEXPORD.TXTPKGTYP = RS!PKGDESC
    FRM_TRNEXPORD.TXTPAYMENTIRR = RS!PAYMENTBYIRRLC
    FRM_TRNEXPORD.TXTCIFFOB = RS!CIFFOB
    FRM_TRNEXPORD.txtEXRate = Val(RS!EXRAT)
    Dim BNKCOD As String
    Dim PKGCOD As String
    BNKCOD = RS!BANKCODE
    PKGCOD = RS!PKGTYPE
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM PKGNGMST WHERE CODE='" & PKGCOD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      FRM_TRNEXPORD.TXTPKGTYP = RS!NAME
    End If
       
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM REFMST WHERE CODE='" & BNKCOD & "' AND CATA='L'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       FRM_TRNEXPORD.TXTBANKDTL = RS!NAME
    End If
End If
End With
End Sub

Private Function EXP_REQ() As Boolean
EXP_REQ = False
Dim ISEXPORT As String
Dim EXPRS As ADODB.Recordset
Set EXPRS = New ADODB.Recordset

If EXPRS.State = 1 Then EXPRS.Close
EXPRS.Open "SELECT * FROM SALMANMST WHERE CODE='" & frm_ORDERBOOK.ORDDBCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not EXPRS.EOF Then
   ISEXPORT = Trim(EXPRS!ISEXPORTORDER & "")
End If

If ISEXPORT = "1" Then
  EXP_REQ = True
Else
  EXP_REQ = False
End If
End Function



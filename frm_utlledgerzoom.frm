VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_utlledgerzoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/c Ledger Zoom"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   1440
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   11355
   Begin VB.Frame Frame1 
      Caption         =   "Ledger Zoom"
      Height          =   1080
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.TextBox txtParty 
         Height          =   330
         Left            =   5310
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   5835
      End
      Begin VB.TextBox TXTUNIT 
         Height          =   330
         Left            =   5310
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   645
         Visible         =   0   'False
         Width           =   5835
      End
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   405
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox dtTo 
         Height          =   315
         Left            =   2955
         TabIndex        =   4
         Top             =   405
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "From :"
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "To :"
         Height          =   255
         Left            =   2295
         TabIndex        =   3
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Name :"
         Height          =   255
         Left            =   4410
         TabIndex        =   5
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lblCSCD 
         Caption         =   "UNIT :"
         Height          =   255
         Left            =   4395
         TabIndex        =   7
         Top             =   675
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Total Transactions Found :"
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
         TabIndex        =   9
         Top             =   750
         Width           =   2355
      End
      Begin VB.Label lblTotRec 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA03A5&
         Height          =   255
         Left            =   2490
         TabIndex        =   10
         Top             =   750
         Width           =   120
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A/c Transaction Detail"
      Height          =   5520
      Left            =   0
      TabIndex        =   11
      Top             =   1140
      Width           =   11325
      Begin VB.Frame framCSCD 
         Height          =   630
         Left            =   120
         TabIndex        =   13
         Top             =   4770
         Visible         =   0   'False
         Width           =   4125
         Begin VB.Label LBLUNIT 
            Caption         =   "Selected UNIT"
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   585
            TabIndex        =   15
            Top             =   240
            Width           =   3450
         End
         Begin VB.Label lbldiv 
            Caption         =   "Unit :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   45
            TabIndex        =   14
            Top             =   255
            Width           =   465
         End
      End
      Begin VB.Frame Frame4 
         Height          =   630
         Left            =   6330
         TabIndex        =   18
         Top             =   4770
         Width           =   2970
         Begin VB.Label LBLTOTCRBAL 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000000000.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label LBLTOTDRBAL 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000000000.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   0
            TabIndex        =   22
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   45
            TabIndex        =   19
            Top             =   135
            Width           =   2880
         End
      End
      Begin VB.Frame Frame5 
         Height          =   630
         Left            =   4260
         TabIndex        =   16
         Top             =   4770
         Width           =   2040
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   120
            TabIndex        =   17
            Top             =   150
            Width           =   1755
         End
      End
      Begin MSComctlLib.ListView lstBill 
         Height          =   4500
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   7938
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Ledger Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   270
         Left            =   9360
         TabIndex        =   20
         Top             =   4830
         Width           =   1800
      End
      Begin VB.Label lblLgrBal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000000000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   9360
         TabIndex        =   21
         Top             =   5100
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frm_utlledgerzoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim ALLOW_EDIT As Boolean
Dim TDAMT As Double
Dim TCAMT As Double
Dim IsFlagChange As Boolean
Dim ISLEDGERZOOM As Boolean
Dim rsZoom As Recordset
Dim upAcc As Recordset
Dim M_SELINDX As Double
Dim M_SYSR As String
Dim ctr As Double
Dim CTRIND As Long
Dim RSMST As New ADODB.Recordset
Dim rszom As New ADODB.Recordset
Dim SEL_DVCD As String
Dim salwithchaln As String
Dim purwithgrn As String
Private Sub dtFrom_Change()
    IsFlagChange = True
End Sub

Private Sub dtFrom_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub dtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtFrom_LostFocus()
    
    If Not IsDate(dtFrom) Then
        MsgBox "Please Check Date Format !!", vbInformation, "Wrong Date Format Found !!"
        dtFrom.SetFocus
        Exit Sub
    End If
    
    If txtParty <> Empty And IsFlagChange Then
        Call GenPtyDtl
    End If
    
End Sub

Private Sub dtTo_Change()
    IsFlagChange = True
End Sub

Private Sub dtTo_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtTo_LostFocus()
    
    If Not IsDate(dtTo) Then
        MsgBox "Please Check Date Format !!", vbInformation, "Wrong Date Format Found !!"
        dtTo.SetFocus
        Exit Sub
    End If

    If txtParty <> Empty And IsFlagChange Then
        GenPtyDtl
    End If
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
    If M_USRSECLEVL = "1" Then
      If ReadConfigMaster("0034", 4, "M") = False Then
         ModuleDeniedMessage
         Unload Me
         Exit Sub
      End If
    End If
    If M_SELINDX <> 0 And lstBill.ListItems.COUNT > M_SELINDX Then
        
        lstBill.ListItems(M_SELINDX).Selected = True
        lstBill.ListItems(M_SELINDX).EnsureVisible
        lstBill.SetFocus
    End If
    If txtParty <> Empty Then
      Call GenPtyDtl
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    
    Call CenterChild(frm_Main, Me)
    
    lstBill.ListItems.Clear
    lstBill.ColumnHeaders.Clear
    
    Call CreatelstBillCols
    Me.KeyPreview = True
    M_SELINDX = 1
    
    lblCSCD.Visible = True
    TXTUNIT.Visible = True
    framCSCD.Visible = True

    IsFlagChange = False
    lblLgrBal.Caption = ".00"
    dtFrom = FSDT
    dtTo = FEDT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    M_SELINDX = 0
    ISLEDGERZOOM = False
    'Unload frm_RCPTLST
    'Unload frm_SPLIST
    'Unload frm_CRDBRecLst
    zoomflag = False
    Set frm_utlledgerzoom = Nothing
    
End Sub

Private Sub lstbill_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static SBIT As Byte
    lstBill.Sorted = True
    If SBIT = lvwDescending Then
        lstBill.SortOrder = lvwAscending
        SBIT = lvwAscending
    Else
        lstBill.SortOrder = lvwDescending
        SBIT = lvwDescending
    End If
    lstBill.SortKey = ColumnHeader.Index - 1
End Sub
Private Sub lstBill_DblClick()
    Call LSTBILL_KeyDown(13, 0)
End Sub

Private Sub LSTBILL_GotFocus()
lstBill.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstBill_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lstBill.SelectedItem.Text = "OPN" Or lstBill.SelectedItem.Text = Empty Then Exit Sub
    LBLUNIT.Caption = GETUNT(lstBill.SelectedItem.SubItems(9))
End Sub

Private Sub LSTBILL_LostFocus()
lstBill.BackColor = vbWhite
End Sub

Private Sub txtParty_GotFocus()
txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = (vbKeyBack Or vbKeyDelete) Then
        txtParty = Empty
        txtParty.Tag = Empty
    ElseIf (KeyCode = vbKeyReturn And txtParty = Empty) Or KeyCode = vbKeyF2 Then
        txtParty = Empty
        lstBill.ListItems.Clear
        txtParty = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/C")
        txtParty.Tag = Key
        IsFlagChange = True
        
    ElseIf KeyCode = vbKeyReturn Then
        
    End If
    
    If txtParty = Empty Then Exit Sub
    SendKeys "{TAB}"
    CTRIND = 0
End Sub
Private Sub CreatelstBillCols()
    lstBill.ColumnHeaders.Add 1, "VTYP", "$$$", 599.811
    lstBill.ColumnHeaders.Add 2, "DATE", "V.Date", 1184.882, lvwColumnLeft
    lstBill.ColumnHeaders.Add 3, "VBNO", "V. No.", 1000.063, lvwColumnLeft
    lstBill.ColumnHeaders.Add 4, "PCOD", "Party Name", 3300.095, lvwColumnLeft
    lstBill.ColumnHeaders.Add 5, "DAMNT", "Debit Amount", 1560.189, lvwColumnRight
    lstBill.ColumnHeaders.Add 6, "CAMNT", "Credit Amount", 1560.189, lvwColumnRight
    lstBill.ColumnHeaders.Add 7, "CDNO", "Chq. No", 1124.787, lvwColumnRight
    lstBill.ColumnHeaders.Add 8, "SRNO", "", 0
    lstBill.ColumnHeaders.Add 9, "OnAccount", "Recon", 345.2599
    lstBill.ColumnHeaders.Add 10, "UNIT", "", 0, lvwColumnRight
End Sub
Public Function GenPtyDtl()
    If sel_untcod = Empty Then Exit Function
    Dim DSPDAT As New ADODB.Recordset
    Set DSPDAT = New ADODB.Recordset
    Dim OPENING As Double
    Dim BALANCE As Double
    Dim TOTCRBAL As Double
    Dim TOTDRBAL As Double
    OPENING = 0
    TOTCRBAL = 0
    TOTDRBAL = 0
    BALANCE = 0
    TDAMT = 0
    TCAMT = 0
    Dim LSTKEY
    lstBill.ListItems.Clear
    Dim lst As ListItem
    SQL = Empty
    SQL = "SELECT TRNMAN.VTYP,TRNMAN.DATE,TRNMAN.VBNO,TRNMAN.RCOD,TRNMAN.DAMT,TRNMAN.CAMT,ISNULL(TRNMAN.CDNO,'') AS CDNO,TRNMAN.COMP,TRNMAN.SRNO,TRNMAN.UNIT,ISNULL(ACCMST.NAME,'Multiple Entries') AS NAME FROM TRNMAN LEFT JOIN ACCMST ON ACCMST.CODE=TRNMAN.RCOD WHERE TRNMAN.DATE<='" & Format(CDate(dtTo), "MM/DD/YYYY") & "' and trnman.Acod='" & txtParty.Tag & "' AND RECSTAT<>'D' AND COMP='" & compPth & "' AND TRNMAN.UNIT IN (" & sel_untcod & ") AND TRNMAN.VTYP <>'BNK' ORDER BY DATE"
    If DSPDAT.State = 1 Then DSPDAT.Close
    DSPDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If DSPDAT.EOF = True Then
      If txtParty.Enabled = True Then
      
        txtParty.SetFocus
      End If
      MsgBox "No Data Exist"
      Exit Function
    End If
    OPENING = 0
    TOTCRBAL = 0
    TOTDRBAL = 0
    Do While Not DSPDAT.EOF And DSPDAT!Date < CDate(dtFrom)
     OPENING = OPENING + DSPDAT!damt - DSPDAT!camt
     DSPDAT.MoveNext
     If DSPDAT.EOF Then
       Exit Do
     End If
    Loop
    'Add Opening to lst box
    Set lst = lstBill.ListItems.Add(, "OPN", "OPN")
    lst.ForeColor = vbMagenta
    lst.Bold = True
    lst.ListSubItems.Add 1, "OD", dtFrom, , "DATE"
    lst.ListSubItems.Add 2, "OV", "XXXXXX", , "Bill No"
    lst.ListSubItems.Add 3, "OA", "Opening Balance", , "Account Description"
    lst.ListSubItems(3).Bold = True
    If OPENING > 0 Then
       lst.ListSubItems.Add 4, "ODA", Format(Abs(OPENING), "#.00"), , "Debit Amount"
       TDAMT = TDAMT + Abs(OPENING)
       TOTDRBAL = TOTDRBAL + Abs(OPENING)
      Else
       lst.ListSubItems.Add 4, "ODA", "0", , "Debit Amount"
       lst.ListSubItems.Add 5, "OCA", Format(Abs(OPENING), "#.00"), , "Credit Amount"
       TCAMT = TCAMT + Abs(OPENING)
       TOTCRBAL = TOTCRBAL + Abs(OPENING)
    End If
    TDAMT = 0
    TCAMT = 0
    'Add Trancasation to the list
    If Not DSPDAT.EOF Then
      Do While Not DSPDAT.EOF And DSPDAT!Date <= CDate(dtTo)
        Set lst = lstBill.ListItems.Add(, , DSPDAT!VTYP)
        LSTKEY = Trim(DSPDAT!VBNO) + "_" + CStr(ctr)
        lst.ListSubItems.Add 1, "D" + LSTKEY, Trim(DSPDAT!Date), , "DATE"
        lst.ListSubItems.Add 2, "V" + LSTKEY, Trim(DSPDAT!VBNO), , "Bill No"
        lst.ListSubItems.Add 3, "A" + LSTKEY, Trim(DSPDAT!Name), , "Account Description"
        lst.ListSubItems.Add 4, "DA" + LSTKEY, Format(Trim(DSPDAT!damt), "##############.00"), , "Debit Amount"
        TDAMT = TDAMT + DSPDAT!damt
        TOTDRBAL = TOTDRBAL + DSPDAT!damt
        lst.ListSubItems.Add 5, "CA" + LSTKEY, Format(Trim(DSPDAT!camt), "#############.00"), , "Credit Amount"
        TCAMT = TCAMT + DSPDAT!camt
        TOTCRBAL = TOTCRBAL + DSPDAT!camt
        lst.ListSubItems.Add 6, "C" + LSTKEY, IIf(IsNull(DSPDAT!CDNO), Empty, Trim(DSPDAT!CDNO)), , ""
        lst.ListSubItems.Add 7, "S" + LSTKEY, Trim(DSPDAT!SRNO), , ""
        Dim RPTDAT As ADODB.Recordset
        Set RPTDAT = New ADODB.Recordset
        If DSPDAT!VTYP = "REC" Or DSPDAT!VTYP = "PAY" Or DSPDAT!VTYP = "RSL" Or DSPDAT!VTYP = "RPR" Or DSPDAT!VTYP = "JDN" Or DSPDAT!VTYP = "JCN" Then
          If RPTDAT.State = 1 Then RPTDAT.Close
          RPTDAT.Open "SELECT * FROM RPTRAN WHERE COMP='" & compPth & "' AND VTYP='" & DSPDAT!VTYP & "' AND SRNO='" & DSPDAT!SRNO & "'", CN, adOpenDynamic, adLockOptimistic
          If Not RPTDAT.EOF Then
            If RPTDAT!ONAC = "U" Then
              lst.ListSubItems.Add 8, "" + LSTKEY, "U", , ""
             Else
              lst.ListSubItems.Add 8, "" + LSTKEY, "A", , ""
            End If
           Else
            lst.ListSubItems.Add 8, "" + LSTKEY, "U", , ""
          End If
         Else
          If RPTDAT.State = 1 Then RPTDAT.Close
          RPTDAT.Open "SELECT * FROM RPTRAN WHERE COMP='" & compPth & "' AND BSR1='" & DSPDAT!VTYP & "' AND BSR2='" & DSPDAT!SRNO & "'", CN, adOpenDynamic, adLockOptimistic
          If RPTDAT.EOF Then
             lst.ListSubItems.Add 8, "" + LSTKEY, "N", , ""
            Else
             lst.ListSubItems.Add 8, "" + LSTKEY, "Y", , ""
          End If
        End If
        'lst.ListSubItems.Add 8, "O" + LSTKEY, "", , ""
        lst.ListSubItems.Add lst.ListSubItems.COUNT + 1, "O" + LSTKEY + DSPDAT!unit, Trim(DSPDAT!unit), , ""
        DSPDAT.MoveNext
        If DSPDAT.EOF Then
          Exit Do
        End If
      Loop
    End If
    BALANCE = OPENING + TDAMT - TCAMT
    LBLTOTDRBAL = Format(TOTDRBAL, "#########.00")
    LBLTOTCRBAL = Format(TOTCRBAL, "#########.00")
    lblLgrBal = Format(BALANCE, "############.00")
    If CTRIND > 0 Then
      If lstBill.ListItems.COUNT > CTRIND Then
        lstBill.ListItems(CTRIND).Selected = True
      End If
    End If
End Function
Private Sub txtParty_LostFocus()
    txtParty.BackColor = vbWhite
    If TXTUNIT <> Empty And IsFlagChange Then
        IsFlagChange = False
        Call GenPtyDtl
    End If
End Sub

Private Sub TXTUNIT_GotFocus()
 TXTUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTUNIT = Empty) Then
        NEW_VISIBLE = False
        Key = Empty
        LOAD frm_askunit
        If frm_askunit.LSTUNIT.ListCount > 0 Then
            frm_askunit.Show 1
        End If
        TXTUNIT = sel_untnam
        If TXTUNIT = Empty Then Exit Sub
        Unload frm_askunit
        'MUNIT = sel_untcod
        Call GenPtyDtl
        If lstBill.ListItems.COUNT > 0 Then lstBill.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        lstBill.SetFocus
    ElseIf KeyCode = vbKeyDelete Then
        TXTUNIT = Empty
    End If
    If TXTUNIT.Text = cUName Then
      ALLOW_EDIT = True
     Else
      ALLOW_EDIT = False
    End If
End Sub
Private Sub LSTBILL_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If TXTUNIT.Text = UntNm Then
       ALLOW_EDIT = True
      Else
       ALLOW_EDIT = False
    End If
    If ALLOW_EDIT = True Then
      'O.k
     Else
      MsgBox "SELECT CURRENT UNIT ONLY TO EDIT"
      lstBill.SetFocus
      Exit Sub
    End If
  End If
  Dim sel_vtyp As String
  Dim SEL_SRNO As String
  Dim SEL_COMP As String
  Dim SEL_UNIT As String
  Dim lstitem As ListItem
  Dim DSPDAT As New ADODB.Recordset
  Dim DEPO_PCOD As String
  If lstBill.ListItems.COUNT < 1 Then Exit Sub
  sel_vtyp = lstBill.SelectedItem.Text
  If sel_vtyp = "OPN" Or Trim(sel_vtyp) = Empty Then Exit Sub
  zoomflag = True
  CTRIND = lstBill.SelectedItem.Index
  M_SELINDX = 0
  Set DSPDAT = New ADODB.Recordset
  If DSPDAT.State = 1 Then DSPDAT.Close
  SEL_COMP = compPth
  SEL_SRNO = lstBill.SelectedItem.SubItems(7)
  SEL_UNIT = lstBill.SelectedItem.SubItems(9)
  If KeyCode = vbKeyDelete Then
    If Not cUName = "ADMIN" Then Exit Sub
    If UCase(M_COMPBILL) = "GSS" Or UCase(M_COMPBILL) = "GSL" Then Exit Sub
    Select Case sel_vtyp
    Case "SAL"
    
      'Check Further Entry Exist or not
      If RS.State = 1 Then RS.Close
      RS.Open "select * from rptran where comp='" & compPth & "' and bsr1='" & sel_vtyp & "' and bsr2='" & SEL_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        MsgBox "Further Entry Exist Can Not Delete !!! "
        Exit Sub
      End If
      
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        DEPO_PCOD = RS!PCOD
       Else
        DEPO_PCOD = Empty
      End If
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM DEPOTMST WHERE COMP='" & compPth & "' AND CODE='" & DEPO_PCOD & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        MsgBox "Depot transfer entry can not delete !!! "
        lstBill.SetFocus
        Exit Sub
      End If
      'Delete records from trnman
      CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      
      'Delete records from billmain
      CN.Execute "DELETE FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      'Update records from sptran
      CN.Execute "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "DELETE FROM PURTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "UPDATE SPTRAN SET RTYP=NULL, RSRN=NULL, RSRC=NULL WHERE COMP='" & compPth & "' AND RTYP='" & sel_vtyp & "' AND RSRN='" & SEL_SRNO & "'"
      lstBill.SetFocus
      MsgBox "Delete Succefuly Referesh it."
      Call GenPtyDtl
      Exit Sub
    Case "RSL"
      
      'Delete records from trnman
      CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      'Delete records from billmain
      CN.Execute "DELETE FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      'Update records from sptran
      CN.Execute "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "DELETE FROM PURTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "DELETE FROM RPTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      lstBill.SetFocus
      MsgBox "Delete Succefuly Referesh it."
      Call GenPtyDtl
      Exit Sub
    Case "PRM"
      'Check Further Entry Exist or not
      
      If RS.State = 1 Then RS.Close
      RS.Open "select * from rptran where comp='" & compPth & "' and bsr1='" & sel_vtyp & "' and bsr2='" & SEL_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        MsgBox "Further Entry Exist Can Not Delete !!! "
        lstBill.SetFocus
        Exit Sub
      End If
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM PURMAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        DEPO_PCOD = RS!PCOD
       Else
        DEPO_PCOD = Empty
      End If
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM DEPOTMST WHERE COMP='" & compPth & "' AND CODE='" & DEPO_PCOD & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        MsgBox "Depot transfer entry can not delete !!! "
        lstBill.SetFocus
        Exit Sub
      End If
      'Delete records from trnman
      CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      'Delete records from billmain
      CN.Execute "DELETE FROM PURMAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      'Update records from sptran
      Dim PURNO As Double
      PURNO = 1
      Dim GRNRS As New ADODB.Recordset
      Set GRNRS = New ADODB.Recordset
      If GRNRS.State = 1 Then GRNRS.Close
      GRNRS.Open "SELECT * FROM PURTRAN WHERE COMP='" & compPth & "' AND RTYP='" & sel_vtyp & "' AND RSRN='" & SEL_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
      Do While Not GRNRS.EOF
        CN.Execute "UPDATE GRN SET PSNO=NULL WHERE COMP='" & compPth & "' AND VTYP='" & GRNRS!VTYP & "' AND SRNO='" & GRNRS!SRNO & "'"
        GRNRS.MoveNext
      Loop
      If GRNRS.State = 1 Then GRNRS.Close
      GRNRS.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND RTYP='" & sel_vtyp & "' AND RSRN='" & SEL_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
      Do While Not GRNRS.EOF
        CN.Execute "UPDATE GRN SET PSNO=NULL WHERE COMP='" & compPth & "' AND VTYP='" & GRNRS!VTYP & "' AND SRNO='" & GRNRS!SRNO & "'"
        GRNRS.MoveNext
      Loop
      If GRNRS.State = 1 Then GRNRS.Close
      CN.Execute "DELETE FROM PURTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "UPDATE PURTRAN SET RTYP=NULL, RSRN=NULL, RSRC=NULL WHERE COMP='" & compPth & "' AND RTYP='" & sel_vtyp & "' AND RSRN='" & SEL_SRNO & "'"
      CN.Execute "UPDATE STORETRAN SET RTYP=NULL, RSRN=NULL, RSRC=NULL WHERE COMP='" & compPth & "' AND RTYP='" & sel_vtyp & "' AND RSRN='" & SEL_SRNO & "'"
      
      MsgBox "Delete Succefuly Referesh it."
      Call GenPtyDtl
      Exit Sub
     Case "RPR"
      
      'Delete records from trnman
      CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      'Delete records from billmain
      CN.Execute "DELETE FROM PURMAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      'Update records from sptran
      CN.Execute "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "DELETE FROM PURTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      CN.Execute "DELETE FROM RPTRAN WHERE COMP='" & compPth & "' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
      lstBill.SetFocus
      MsgBox "Delete Succefuly Referesh it."
      Call GenPtyDtl
      Exit Sub
    End Select
  End If
  If Not KeyCode = vbKeyReturn Then Exit Sub
  Select Case sel_vtyp
   
   Case "EXP"
     frm_trnexp.UNTNAM = LBLUNIT
     frm_trnexp.UNTCOD = lstBill.SelectedItem.SubItems(9)
     LOAD FRM_TRNLSTEXP
     With FRM_TRNLSTEXP
       SQL = Empty
       SQL = "SELECT TRNMAN.*,ACCMST.NAME AS ACNM FROM TRNMAN INNER JOIN ACCMST ON (ACCMST.CODE=TRNMAN.RCOD) WHERE COMP='" & compPth & "' AND UNIT='" & SEL_UNIT & "' AND RECSTAT<>'D' AND VTYP='EXP' AND SRNO='" & SEL_SRNO & "'"
       Set RS = New ADODB.Recordset
       If RS.State = 1 Then RS.Close
       RS.Open SQL, CN, adOpenKeyset, adLockOptimistic
       If RS.EOF Then
           MsgBox "No Data Found"
           Exit Sub
       End If
       Do While Not RS.EOF
        Set lstitem = .lstBill.ListItems.Add
        lstitem.Text = Format(RS![Date], "dd/MM/yyyy")
        lstitem.SubItems(1) = RS![VBNO]
        lstitem.SubItems(2) = RS!CDNO & ""
        lstitem.SubItems(3) = RS![ACNM]
        lstitem.SubItems(4) = Format(RS!AMNT, "###########.00")
        lstitem.SubItems(6) = RS!VTYP
        lstitem.SubItems(7) = RS!SRNO
        RS.MoveNext
       Loop
       .cmdOk.Enabled = True
       .cmdOk.Default = True
       .CMDOK_Click
     End With
   Case "JPY"
     FRM_BANKPAY.UNTNAM = LBLUNIT
     FRM_BANKPAY.UNTCOD = lstBill.SelectedItem.SubItems(9)
     LOAD FRM_LSTBANKPAY
     With FRM_LSTBANKPAY
        SQL = Empty
        SQL = "SELECT TRNMAN.*,ACCMST.NAME AS ACNM FROM TRNMAN INNER JOIN ACCMST ON (ACCMST.CODE=TRNMAN.RCOD) WHERE COMP='" & compPth & "' AND UNIT='" & SEL_UNIT & "' AND RECSTAT<>'D' AND VTYP='PAY' AND SRNO='" & SEL_SRNO & "'"
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
        RS.Open SQL, CN, adOpenKeyset, adLockOptimistic
        If RS.EOF Then
            MsgBox "No Data Found"
            Exit Sub
        End If
        Do While Not RS.EOF
         Set lstitem = .lstBill.ListItems.Add
         lstitem.Text = Format(RS![Date], "dd/MM/yyyy")
         lstitem.SubItems(1) = RS![VBNO]
         lstitem.SubItems(2) = RS!CDNO & ""
         lstitem.SubItems(3) = RS![ACNM]
         lstitem.SubItems(4) = Format(RS!AMNT, "###########.00")
         lstitem.SubItems(6) = RS!VTYP
         lstitem.SubItems(7) = RS!SRNO
         RS.MoveNext
        Loop
        .cmdOk.Enabled = True
        .cmdOk.Default = True
       .CMDOK_Click
     End With
   Case "JRC"
     Frm_BankRec.UNTNAM = LBLUNIT
     Frm_BankRec.UNTCOD = lstBill.SelectedItem.SubItems(9)
     LOAD FRM_LSTBANKREC
     With FRM_LSTBANKREC
        SQL = Empty
        SQL = "SELECT TRNMAN.*,ACCMST.NAME AS ACNM FROM TRNMAN INNER JOIN ACCMST ON (ACCMST.CODE=TRNMAN.RCOD) WHERE COMP='" & compPth & "' AND UNIT='" & SEL_UNIT & "' AND RECSTAT<>'D' AND VTYP='REC' AND SRNO='" & SEL_SRNO & "'"
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
        RS.Open SQL, CN, adOpenKeyset, adLockOptimistic
        If RS.EOF Then
            MsgBox "No Data Found"
            Exit Sub
        End If
        Do While Not RS.EOF
         Set lstitem = .lstBill.ListItems.Add
         lstitem.Text = Format(RS![Date], "dd/MM/yyyy")
         lstitem.SubItems(1) = RS![VBNO]
         lstitem.SubItems(2) = RS!CDNO & ""
         lstitem.SubItems(3) = RS![ACNM]
         lstitem.SubItems(4) = Format(RS!AMNT, "###########.00")
         lstitem.SubItems(6) = RS!VTYP
         lstitem.SubItems(7) = RS!SRNO
         RS.MoveNext
        Loop
        .cmdOk.Enabled = True
        .cmdOk.Default = True
        .CMDOK_Click
     End With
   Case "PAY"
     FRM_BANKPAY.UNTNAM = LBLUNIT
     FRM_BANKPAY.UNTCOD = lstBill.SelectedItem.SubItems(9)
     LOAD FRM_LSTBANKPAY
     With FRM_LSTBANKPAY
        SQL = Empty
        SQL = "SELECT TRNMAN.*,ACCMST.NAME AS ACNM FROM TRNMAN INNER JOIN ACCMST ON (ACCMST.CODE=TRNMAN.RCOD) WHERE COMP='" & compPth & "' AND UNIT='" & SEL_UNIT & "' AND ACOD='" & txtParty.Tag & "' AND RECSTAT<>'D' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
        RS.Open SQL, CN, adOpenKeyset, adLockOptimistic
        If RS.EOF Then
            MsgBox "No Data Found"
            Exit Sub
        End If
        Do While Not RS.EOF
         Set lstitem = .lstBill.ListItems.Add
         lstitem.Text = Format(RS![Date], "dd/MM/yyyy")
         lstitem.SubItems(1) = RS![VBNO]
         lstitem.SubItems(2) = RS!CDNO & ""
         lstitem.SubItems(3) = RS![ACNM]
         lstitem.SubItems(4) = Format(RS!AMNT, "###########.00")
         lstitem.SubItems(6) = RS!VTYP
         lstitem.SubItems(7) = RS!SRNO
         RS.MoveNext
        Loop
        .cmdOk.Enabled = True
        .cmdOk.Default = True
       .CMDOK_Click
     End With
   
   Case "PSR"

   Case "REC"
     Frm_BankRec.UNTNAM = LBLUNIT
     Frm_BankRec.UNTCOD = lstBill.SelectedItem.SubItems(9)
     LOAD FRM_LSTBANKREC
     With FRM_LSTBANKREC
        SQL = Empty
        SQL = "SELECT TRNMAN.*,ACCMST.NAME AS ACNM FROM TRNMAN INNER JOIN ACCMST ON (ACCMST.CODE=TRNMAN.RCOD) WHERE COMP='" & compPth & "' AND UNIT='" & SEL_UNIT & "' AND ACOD='" & txtParty.Tag & "' AND RECSTAT<>'D' AND VTYP='" & sel_vtyp & "' AND SRNO='" & SEL_SRNO & "'"
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
        RS.Open SQL, CN, adOpenKeyset, adLockOptimistic
        If RS.EOF Then
            MsgBox "No Data Found"
            Exit Sub
        End If
        Do While Not RS.EOF
         Set lstitem = .lstBill.ListItems.Add
         lstitem.Text = Format(RS![Date], "dd/MM/yyyy")
         lstitem.SubItems(1) = RS![VBNO]
         lstitem.SubItems(2) = RS!CDNO & ""
         lstitem.SubItems(3) = RS![ACNM]
         lstitem.SubItems(4) = Format(RS!AMNT, "###########.00")
         lstitem.SubItems(6) = RS!VTYP
         lstitem.SubItems(7) = RS!SRNO
         RS.MoveNext
        Loop
        .cmdOk.Enabled = True
        .cmdOk.Default = True
        .CMDOK_Click
     End With
   Case "RPR"
     
   Case "RSL"
     'MUNIT = UNCD
     'UNCD = lstBill.SelectedItem.SubItems(9)
     
     
     'UNCD = MUNIT
   Case "SAL"
     'Find wether direct sale or challan sale
     'Data of sptran with vtyp ='dpf' then sale through challan and
     'if vtyp='sal' then direct sale
    
     salwithchaln = "N"
     Set rszom = New ADODB.Recordset
     If rszom.State = 1 Then rszom.Close
     rszom.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND RTYP='SAL' AND RSRN='" & SEL_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
     If Not rszom.EOF Then
       If rszom!VTYP = "DPF" Then
         salwithchaln = "Y"
        Else
         salwithchaln = "N"
       End If
      Else
       salwithchaln = "N"
     End If
     If salwithchaln = "N" Then
     
        'frm_Directsal.UNTNAM = LBLUNIT
        'frm_Directsal.UNTCOD = lstBill.SelectedItem.SubItems(9)
        LOAD FRM_SPLISTDIRSAL
        'MUNIT = UNCD
        'UNCD = lstBill.SelectedItem.SubItems(9)
        
        With FRM_SPLISTDIRSAL
          SQL = Empty
          SQL = "SELECT DISTINCT BILLMAIN.*,ACCMST.NAME FROM BILLMAIN INNER JOIN ACCMST ON BILLMAIN.PCOD=ACCMST.CODE WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.VTYP='SAL' AND BILLMAIN.SRNO='" & SEL_SRNO & "' AND BILLMAIN.RECSTAT<>'D' ORDER BY BILLMAIN.DATE,BILLMAIN.VBNO"
          Set rszom = New ADODB.Recordset
          If rszom.State = 1 Then RS.Close
          rszom.Open SQL, CN, adOpenKeyset, adLockOptimistic
          If rszom.EOF Then
              MsgBox "No Data Found"
              Exit Sub
          End If
          Do While Not rszom.EOF
           'frm_Directsal.M_DBCD = rszom!DBCD
           DIVCOD = rszom!DVCD & ""
           SEL_DVCD = rszom!DVCD & ""
           frm_Directsal.Tag = SEL_DVCD
           
           Set RSMST = New ADODB.Recordset
           If RSMST.State = 1 Then RSMST.Close
           RSMST.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & lstBill.SelectedItem.SubItems(9) & "' AND CODE='" & SEL_DVCD & "'", CN, adOpenDynamic, adLockOptimistic
           If Not RSMST.EOF Then
             frm_Directsal.LBLDIV = RSMST!Name & ""
           End If
           If RSMST.State = 1 Then RSMST.Close
           RSMST.Open "SELECT * FROM DAYBOK WHERE COMP='" & compPth & "' AND UNIT='" & lstBill.SelectedItem.SubItems(9) & "' AND VTYP='SAL' AND DVCD='" & SEL_DVCD & "' AND DBCD='" & rszom!DBCD & "'", CN, adOpenDynamic, adLockOptimistic
           If Not RSMST.EOF Then
             'frm_Directsal.LBLDAYBOK = RSMST!Name & ""
            Else
             'frm_Directsal.LBLDAYBOK = "??????"
           End If
           'frm_Directsal.Caption = frm_Directsal.LBLDAYBOK
           'frm_Directsal.SALBOK_DIRSAL = frm_Directsal.LBLDAYBOK
           Set lstitem = .lstBill.ListItems.Add
           lstitem.Text = Format(rszom![Date], "dd/MM/yyyy")
           lstitem.SubItems(1) = rszom![VBNO]
           lstitem.SubItems(2) = rszom![Name]
           lstitem.SubItems(3) = Format(rszom!TQTY, "########.000")
           lstitem.SubItems(4) = Format(rszom!BNET, "#############.00")
           lstitem.SubItems(6) = rszom!VTYP
           lstitem.SubItems(7) = rszom!SRNO
           rszom.MoveNext
          Loop
          'frm_Directsal.FIL_Billingterm
          .cmdOk.Enabled = True
          .cmdOk.Default = True
          .CMDOK_Click
        End With
       Else
        'Sale with challan
        frm_transale.UNTNAM = LBLUNIT
        frm_transale.UNTCOD = lstBill.SelectedItem.SubItems(9)
        LOAD frm_SPLIST
        'MUNIT = UNCD
        'UNCD = lstBill.SelectedItem.SubItems(9)
        
        With frm_SPLIST
          SQL = Empty
          SQL = "SELECT DISTINCT BILLMAIN.*,ACCMST.NAME FROM BILLMAIN INNER JOIN ACCMST ON BILLMAIN.PCOD=ACCMST.CODE WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.VTYP='SAL' AND BILLMAIN.SRNO='" & SEL_SRNO & "' AND BILLMAIN.RECSTAT<>'D' ORDER BY BILLMAIN.DATE,BILLMAIN.VBNO"
          Set rszom = New ADODB.Recordset
          If rszom.State = 1 Then RS.Close
          rszom.Open SQL, CN, adOpenKeyset, adLockOptimistic
          If rszom.EOF Then
              MsgBox "No Data Found"
              Exit Sub
          End If
          Do While Not rszom.EOF
           frm_transale.M_DBCD = rszom!DBCD
           DIVCOD = rszom!DVCD & ""
           SEL_DVCD = rszom!DVCD & ""
           frm_transale.Tag = SEL_DVCD
           Set RSMST = New ADODB.Recordset
           If RSMST.State = 1 Then RSMST.Close
           
           RSMST.Open "SELECT * FROM DAYBOK WHERE DBCD='" & frm_transale.M_DBCD & "' AND COMP='" & compPth & "' AND UNIT='" & lstBill.SelectedItem.SubItems(9) & "' AND VTYP='SAL' AND DVCD='" & SEL_DVCD & "'", CN, adOpenDynamic, adLockOptimistic
           If Not RSMST.EOF Then
             frm_transale.LBLDAYBOK = RSMST!Name & ""
            Else
             frm_transale.LBLDAYBOK = "??????"
           End If
           frm_transale.Caption = frm_transale.LBLDAYBOK
           frm_transale.SALBOK = frm_transale.LBLDAYBOK
           Set lstitem = .lstBill.ListItems.Add
           lstitem.Text = Format(rszom![Date], "dd/MM/yyyy")
           lstitem.SubItems(1) = rszom![VBNO]
           lstitem.SubItems(2) = rszom![Name]
           lstitem.SubItems(3) = Format(rszom!TQTY, "########.000")
           lstitem.SubItems(4) = Format(rszom!BNET, "#############.00")
           lstitem.SubItems(6) = rszom!VTYP
           lstitem.SubItems(7) = rszom!SRNO
           rszom.MoveNext
          Loop
          frm_transale.FIL_Billingterm
          .cmdOk.Enabled = True
          .cmdOk.Default = True
          .CMDOK_Click
        End With
     End If
   Case "PRM"
    
    
    
  End Select
  Call GenPtyDtl
End Sub

Private Sub TXTUNIT_LostFocus()
TXTUNIT.BackColor = vbWhite
End Sub

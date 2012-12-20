VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_RECONCILIATIONLST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Sale Order"
   ClientHeight    =   5340
   ClientLeft      =   390
   ClientTop       =   1740
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   11445
   Begin VB.TextBox TXTPCOD 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4800
      Width           =   4695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   8520
      TabIndex        =   3
      Top             =   4920
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   9600
      TabIndex        =   4
      Top             =   4920
      Width           =   960
   End
   Begin VB.TextBox SOBNO 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame frmIVR 
      Height          =   4245
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   11205
      Begin MSComctlLib.ListView lst 
         Height          =   3945
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6959
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sr."
            Object.Width           =   707
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order No."
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agent Name"
            Object.Width           =   2648
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Party Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Name"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Order Qnty"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Disp. Qty"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Bal Qnty"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "PCOD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "BRCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "ICOD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "TXCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "BRCD"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Agent"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label LBLORD 
      Caption         =   "Order No."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Press Enter To See the List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Choose Sales Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "FRM_RECONCILIATIONLST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AGC_NAM As String
Dim agc_cod As String

Private Sub cmdCancel_Click()
  'Call frmOrderReconcile.cmdCancel_Click
  Unload Me
End Sub

Private Sub cmdOk_Click()
   If lst.ListItems.Count > 0 Then
    frmOrderReconcile.lblORDNO = Trim(lst.SelectedItem.SubItems(2))
    frmOrderReconcile.BALQTY = nstr(lst.SelectedItem.SubItems(8), 10, 3)
    frmOrderReconcile.lblORDQTY = nstr(lst.SelectedItem.SubItems(6), 10, 3)
    frmOrderReconcile.Party = Trim(lst.SelectedItem.SubItems(10))
    frmOrderReconcile.AGENT = Trim(lst.SelectedItem.SubItems(11))
    frmOrderReconcile.ITEM = Trim(lst.SelectedItem.SubItems(12))
    frmOrderReconcile.tax = Trim(lst.SelectedItem.SubItems(13))
    
    frmOrderReconcile.lblAGENT = Trim(lst.SelectedItem.SubItems(3))
    frmOrderReconcile.lblITEM = Trim(lst.SelectedItem.SubItems(5))
    frmOrderReconcile.LBLPARTY = Trim(lst.SelectedItem.SubItems(4))
    frmOrderReconcile.LBLORDRAT = Trim(lst.SelectedItem.SubItems(9))
  End If
    If frmOrderReconcile.DISPQTY.Enabled Then frmOrderReconcile.DISPQTY.SetFocus
  Unload Me
End Sub

Private Sub FIL_LST(Optional PCOD As String)
  Dim TEMPRS As New ADODB.Recordset
  Set TEMPRS = New ADODB.Recordset
  Dim lstItm As ListItem
  Dim DSPDATA As New ADODB.Recordset
  Dim DSPVTYP As String
  Dim DSPSRNO As String
  Dim DSPOSRC As String
  Dim SYSDAT As Date
  Dim I As Integer
  Dim SQL As String
  
  If TEMPRS.State = adStateOpen Then TEMPRS.Close
  TEMPRS.Filter = adFilterNone
  SYSDAT = Now
  
  SQL = "select * from SALORDTRN where COMP='" & compPth & "' AND vtyp='DOS' and DFLG ='N' AND QNTY>0 AND RECSTAT<>'D' AND DODT <='" & Format(SYSDAT, "MM/DD/YYYY") & "' AND DOSTAT='Y' "
  
  If Trim(SOBNO) <> Empty Then
     SQL = SQL & "AND DONO='" & SOBNO & "' "
  End If
  
  If Trim(PCOD) <> Empty Then
     SQL = SQL & "AND BRCD='" & PCOD & "' "
  End If
  
  SQL = SQL & "ORDER BY DODT"
    
  TEMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    
    Dim accrs As New ADODB.Recordset
    Dim DOS_DONO As String
    Dim DOS_ARCD As String
    Dim DOS_DPTY As String
    Dim DOS_DENI As String
    Dim DOS_RATE As Double
    Dim DOS_PNAM As String
    Dim DOS_BRCD As String
    Dim DOS_BRNM As String
        
    If Not TEMPRS.EOF Then
        I = 1
        Do While Not TEMPRS.EOF
            DOS_DONO = Trim(TEMPRS!dono & "")
            DOS_DENI = TEMPRS!ICOD
            DOS_RATE = TEMPRS!ARAT
            DOS_PNAM = TEMPRS!PCOD
            DOS_BRCD = TEMPRS!BRCD
            
            If accrs.State = 1 Then accrs.Close
            accrs.Open "SELECT * FROM ITMMST WHERE CODE='" & DOS_DENI & "'", CN, adOpenKeyset, adLockPessimistic
            If Not accrs.EOF Then
              DOS_DENI = accrs!Name
             Else
              DOS_DENI = Empty
            End If
            accrs.Close
            If accrs.State = 1 Then accrs.Close
            accrs.Open "SELECT * FROM ACCMST WHERE CODE='" & DOS_PNAM & "'", CN, adOpenDynamic, adLockOptimistic
            If Not accrs.EOF Then
               DOS_PNAM = accrs!Name
              Else
               DOS_PNAM = Empty
            End If
            
            If accrs.State = 1 Then accrs.Close
            accrs.Open "SELECT * FROM REFMST WHERE CODE='" & DOS_BRCD & "' AND CATA='B'", CN, adOpenDynamic, adLockOptimistic
            If Not accrs.EOF Then
               DOS_BRNM = accrs!Name
              Else
               DOS_BRNM = Empty
            End If
            
            If DSPDATA.State = adStateOpen Then DSPDATA.Close
            DSPDATA.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM SALORDTRN WHERE COMP='" & compPth & "' AND  VTYP<>'DOS' AND DONO='" & DOS_DONO & "' AND RECSTAT<>'D'", CN, adOpenKeyset, adLockPessimistic
            Dim DO_BALQTY As Double
            DO_BALQTY = 0
            If Not DSPDATA.EOF Then
              DO_BALQTY = TEMPRS!QNTY - DSPDATA!QNTY
             Else
              DO_BALQTY = TEMPRS!QNTY
            End If
            If DO_BALQTY > 0 Then
              Set lstItm = lst.ListItems.Add()
              lstItm.Text = lst.ListItems.Count
              lstItm.SubItems(1) = CStr(Format(TEMPRS![DODT], "dd/mm/yyyy"))
              lstItm.SubItems(2) = Trim(TEMPRS![dono])
              lstItm.SubItems(3) = DOS_BRNM
              lstItm.SubItems(4) = DOS_PNAM
              lstItm.SubItems(5) = DOS_DENI
              lstItm.SubItems(6) = TEMPRS!QNTY
              lstItm.SubItems(7) = DSPDATA!QNTY
                           
              'TEMPRS!dono
              If Not DSPDATA.EOF Then
                lstItm.SubItems(8) = TEMPRS!QNTY - DSPDATA!QNTY
               Else
                lstItm.SubItems(8) = TEMPRS!QNTY
              End If
                lstItm.SubItems(9) = DOS_RATE
                
              lstItm.SubItems(10) = Trim(TEMPRS!PCOD & "")
              lstItm.SubItems(11) = Trim(TEMPRS!BRCD & "")
              lstItm.SubItems(12) = Trim(TEMPRS!ICOD & "")
              lstItm.SubItems(13) = Trim(TEMPRS!TXCD & "")
              lstItm.SubItems(14) = Trim(DOS_BRCD)
            End If
            TEMPRS.MoveNext
        Loop
        
      Else
       If Not Trim(SOBNO) = Empty Then
         MsgBox "Order Not Exist"
         SOBNO.SetFocus
         Exit Sub
       End If
    End If
    
    If lst.ListItems.Count > 0 Then
      lst.ListItems(1).Selected = True
      lst.SetFocus
      cmdOk.Default = True
     Else
      cmdOk.Default = False
    End If
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
  Unload Me
End Sub

Private Sub SOBNO_GotFocus()
 SOBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SOBNO_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then Call FIL_LST
End Sub


Private Sub SOBNO_LostFocus()
SOBNO.BackColor = vbWhite
End Sub

Private Sub txtPCOD_GotFocus()
   TXTPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTPCOD.Text) = Empty Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTPCOD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM REFMST WHERE CATA='B'", 0, TXTPCOD, "Select Agent From List")
        TXTPCOD.Tag = Key
    ElseIf KeyCode = vbKeyDelete Then
        TXTPCOD = Empty
    End If
    Me.KeyPreview = True
    
    If KeyCode = vbKeyReturn And TXTPCOD <> Empty Then
       Call FIL_LST(TXTPCOD.Tag)
    End If
End Sub

Private Sub txtPCOD_LostFocus()
TXTPCOD.BackColor = vbWhite
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCancelOrderList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Sales Order"
   ClientHeight    =   6930
   ClientLeft      =   450
   ClientTop       =   1185
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmDTRNGE 
      Height          =   1200
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   11130
      Begin VB.TextBox ORPTY 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   9015
      End
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
         Height          =   375
         Left            =   9600
         TabIndex        =   4
         Top             =   240
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4200
         TabIndex        =   2
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56426497
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56426497
         CurrentDate     =   38429
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Party :"
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
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date : "
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date : "
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
         Left            =   3360
         TabIndex        =   8
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5070
      Left            =   120
      TabIndex        =   10
      Top             =   1185
      Width           =   11160
      Begin MSComctlLib.ListView lstBill 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8493
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Agent Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Item Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Grade"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Qnty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Rate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Party Order no"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TAX CATEGORY"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Rate Factor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Freight/KG"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   11160
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Enabled         =   0   'False
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
         Left            =   8400
         TabIndex        =   6
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
         Left            =   9720
         TabIndex        =   7
         Top             =   195
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmCancelOrderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VTYP As String
Dim STRSQL As String
Dim M_DBCD As String
Dim TEMPRS As ADODB.Recordset

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdGo_Click()
Dim lstITEM As ListItem
Dim PTYCOD As String
Dim SHDRQ As String
Dim O As Integer
        
    Screen.MousePointer = vbHourglass
    cmdGo.Default = False
    lstBill.ListItems.Clear
    
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT CODE FROM ACCMST WHERE NAME='" & ORPTY & "'", CN, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
          PTYCOD = RS!CODE
         Else
          PTYCOD = Empty
        End If
        If RS.State = 1 Then RS.Close
        
        STRSQL = "Select ORDMAN.*,ACCMST.NAME,FINITMMST.NAME AS ITEM,REFMST.NAME AS BRNM,GRDMST.GRAD AS GRADE from ORDMAN " & _
        "INNER JOIN ACCMST on ORDMAN.PCOD = ACCMST.CODE " & _
        "INNER JOIN FINITMMST on ORDMAN.COMP = FINITMMST.COMP AND ORDMAN.UNIT = FINITMMST.UNIT AND " & _
        "ORDMAN.DCOD = FINITMMST.DVCD AND ORDMAN.ICOD = FINITMMST.CODE " & _
        "INNER JOIN GRDMST on ORDMAN.TRCD = GRDMST.CODE " & _
        "INNER JOIN REFMST on ORDMAN.BRCD = REFMST.CODE where ORDMAN.DBCD='" & frmOrderReconcile.ORDDBCD & _
        "' AND ORDMAN.RECSTAT<>'D' AND ORDMAN.ORDT >= '" & Format(txtFrDate.Value, "MM/dd/YYYY") & _
        "' and ORDMAN.ORDT <= '" & Format(txtToDate.Value, "MM/dd/YYYY") & "' AND ORDMAN.COMP='" & compPth & _
        "' and oflg<>'Y' AND FIN_APRV = 'O' AND (ORDMAN.QNTY - ORDMAN.DOQTY - ORDMAN.DISPATCHQTY - ORDMAN.CANCELQTY) > 0 " 'FOR BALANCE
                
        If Not PTYCOD = Empty Then
          STRSQL = STRSQL & " AND ACCMST.NAME='" & Trim(ORPTY) & "'"
        End If
        
        STRSQL = STRSQL & " ORDER BY ORDT,ORDMAN.ORDN ASC"
        
        If RS.State = 1 Then RS.Close
        RS.Open STRSQL, CN, adOpenDynamic, adLockOptimistic
        STRSQL = Empty
        If RS.EOF = True Then
            MsgBox "There are no Record found.", vbInformation, App.Title
            cmdOk.Enabled = False
        Else
            Do While Not RS.EOF
                Set lstITEM = lstBill.ListItems.ADD
                lstITEM.Text = Format(RS!ORDT, "dd/MM/yyyy")
                lstITEM.SubItems(1) = RS!ORDN & ""
                lstITEM.SubItems(2) = RS![NAME] & ""
                lstITEM.SubItems(3) = RS![BRNM] & ""
                lstITEM.SubItems(4) = RS!Item & ""
                lstITEM.SubItems(5) = RS![GRADE] & ""
                If Not IsNull(RS!QNTY) Then lstITEM.SubItems(6) = RS!QNTY Else lstITEM.SubItems(6) = 0
                lstITEM.SubItems(7) = Format(RS!RATE, "#####.00")
                lstITEM.SubItems(8) = Trim(RS!PORD & "")
                lstITEM.SubItems(9) = RS!RMRK & ""
                lstITEM.SubItems(10) = GetCode("TAXMST", RS!TXCD, "CODE", "NAME")
                lstITEM.SubItems(11) = GetCode("RATEMST", RS!RTCD, "CODE", "NAME")
                lstITEM.SubItems(12) = RS!FREIGHT_PERKG & ""
                RS.MoveNext
            Loop
            cmdOk.Enabled = True
            cmdOk.Default = True
            lstBill.SetFocus
        End If
        Screen.MousePointer = vbNormal
    
End Sub

Private Sub CMDOK_Click()
    With frmOrderReconcile
      .TXTVBNO = Empty: .txtPCOD = Empty: .TXTBRCD = Empty: .txtTXCD = Empty: .TXTICOD = Empty: .TXTOGRD = Empty: .txtTTQty = Empty: .TXTRATE = Empty
      .TXTVBNO = Trim(lstBill.SelectedItem.SubItems(1)): .txtPCOD = Trim(lstBill.SelectedItem.SubItems(2))
      .TXTBRCD = Trim(lstBill.SelectedItem.SubItems(3)): .TXTICOD = Trim(lstBill.SelectedItem.SubItems(4))
      .txtTXCD = Trim(lstBill.SelectedItem.SubItems(10)): .TXTOGRD = Trim(lstBill.SelectedItem.SubItems(5))
      .txtTTQty = Trim(lstBill.SelectedItem.SubItems(6)): .TXTRATE = Trim(lstBill.SelectedItem.SubItems(7))
      .TXTVBDT = Trim(lstBill.SelectedItem.Text)
      .TXTRATEFACTOR = Trim(lstBill.SelectedItem.SubItems(11)): .TXTFREIGHT = Trim(lstBill.SelectedItem.SubItems(12))
    End With
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    
    txtFrDate.Value = GetMinDate
    txtToDate.Value = GetMaxDate
    Set TEMPRS = New Recordset
    Me.KeyPreview = True
    Me.Caption = "List of Order FOR " + frmOrderReconcile.ORDBOK
End Sub

Private Sub lstBill_GotFocus()
lstBill.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstBill_LostFocus()
 lstBill.BackColor = vbWhite
End Sub

Private Sub ORPTY_GotFocus()
 ORPTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ORPTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        ORPTY.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST", 0, ORPTY.Text, "SELECT A/C PARTY")
    End If
    If KeyCode = vbKeyDelete Then
       ORPTY.Text = Empty
    End If
    Me.KeyPreview = True
End Sub

Private Sub ORPTY_LostFocus()
ORPTY.BackColor = vbWhite
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_excisedblist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Excise Db Entry "
   ClientHeight    =   6555
   ClientLeft      =   4140
   ClientTop       =   2850
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9750
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   0
      TabIndex        =   8
      Top             =   5880
      Width           =   9720
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
         Left            =   8520
         TabIndex        =   10
         Top             =   195
         Width           =   1035
      End
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
         Left            =   7200
         TabIndex        =   9
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5070
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   9720
      Begin MSComctlLib.ListView lstBill 
         Height          =   4815
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   9495
         _ExtentX        =   16748
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Voucher No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Doc. No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Supplier"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Manufacturer"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Desc."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Excise Reg."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Supllier Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Remarks"
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9690
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
         Left            =   7320
         TabIndex        =   5
         Top             =   240
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4320
         TabIndex        =   4
         Top             =   285
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56033281
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   255
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56033281
         CurrentDate     =   38429
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
         TabIndex        =   3
         Top             =   285
         Width           =   885
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
         TabIndex        =   1
         Top             =   240
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frm_excisedblist"
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

Private Sub CMDGO_Click()
Dim lstItem As ListItem
Dim PTYCOD As String
Dim SHDRQ As String
Dim O As Integer
        
    Screen.MousePointer = vbHourglass
    cmdGo.Default = False
    lstBill.ListItems.Clear
    
        If RS.State = 1 Then RS.Close
        STRSQL = "Select *,EGPMAN.TTYP AS TTYP1,EGPMAN.EXTRA1 AS SUPTYP,EGPMAN.EXTRA5 AS RMRK," & _
        "SUPMST.NAME AS SUPMST,MFGMST.NAME AS MFGMST,ITMMST.NAME AS ITEM from EGPMAN " & _
        "LEFT JOIN ACCMST SUPMST ON SUPMST.CODE=EGPMAN.CRAC " & _
        "LEFT JOIN ACCMST MFGMST ON MFGMST.CODE=EGPMAN.DRAC " & _
        "LEFT JOIN ITMMST ON ITMMST.CODE=EGPMAN.ICOD " & _
        "WHERE EGPMAN.COMP='" & compPth & "' AND EGPMAN.UNIT='" & UNCD & "' AND VTYP='EXD' AND RECSTAT<>'D' AND " & _
        "DATE >= '" & Format(txtFrDate.Value, "MM/dd/YYYY") & _
        "' and DATE <= '" & Format(txtToDate.Value, "MM/dd/YYYY") & "' ORDER BY DATE,SRNO" 'FOR BALANCE
                  
        If RS.State = 1 Then RS.Close
        RS.Open STRSQL, CN, adOpenDynamic, adLockOptimistic
        STRSQL = Empty
        If RS.EOF = True Then
            MsgBox "There are no Record found.", vbInformation, App.Title
            cmdOk.Enabled = False
        Else
            Do While Not RS.EOF
                Set lstItem = lstBill.ListItems.Add
                lstItem.Text = Format(RS!Date, "dd/MM/yyyy")
                lstItem.SubItems(1) = Trim(RS!SRNO & "")
                lstItem.SubItems(2) = Trim(RS!VBNO & "")
                lstItem.SubItems(3) = Trim(RS!SUPMST & "")
                lstItem.SubItems(4) = Trim(RS!MFGMST & "")
                lstItem.SubItems(5) = Trim(RS!Item & "")
                lstItem.SubItems(6) = Trim(nstr(RS!CENVAT + RS!EDUCESS + RS!H_ED_CESS + RS!A_DUTY, 12, 2))
                lstItem.SubItems(7) = Trim(RS![TTYP1] & "")
                lstItem.SubItems(8) = Trim(RS![SUPTYP] & "")
                lstItem.SubItems(9) = Trim(RS![RMRK] & "")
                RS.MoveNext
            Loop
            cmdOk.Enabled = True
            cmdOk.Default = True
            lstBill.SetFocus
        End If
        Screen.MousePointer = vbNormal
End Sub

Private Sub CMDOK_Click()
    With frm_servicetaxdb
    
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='EXD' AND " & _
              "SRNO='" & Trim(lstBill.SelectedItem.SubItems(1)) & _
              "'  AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
      
         'Store All The Data In Respective Box
         .M_VBNO = Trim(lstBill.SelectedItem.SubItems(1))
         
         Select Case RS!TTYP
         Case "RG23-A"
           .EXCREG.ListIndex = 0
         Case "RG23-C"
           .EXCREG.ListIndex = 1
         Case Else
           .EXCREG.ListIndex = 2
         End Select
         
         
        .ENTRYNO = Trim(RS!SRNO)
        .ENTDAT.Value = RS!Date
        
        .VBNO.Text = Trim(RS!VBNO & "")
        

        .CENVAT.Text = RS!CENVAT
        .TXTCESS.Text = RS!CESS
        .EDUCESS.Text = RS!EDUCESS
        .HEDUCESS.Text = RS!H_ED_CESS

        .ADUTY.Text = RS!A_DUTY
        
        .TXTEXTRA5.Text = RS!EXTRA5 & ""
        .txtpurac.Tag = RS!EXTRA1 & ""
        .TXTVBNO = RS!EXTRA2 & ""
        If IsNull(RS!CHDT) Then
          .txtodat = RS!Date
         Else
          .txtodat = RS!CHDT
        End If
        
        Dim SUP_COD As String, MFG_COD As String, ITM_COD As String
        
        SUP_COD = RS!CRAC
        MFG_COD = RS!DRAC
        ITM_COD = RS!ICOD & ""
        
        
        
        
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & .txtpurac.Tag & "'", CN, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
          .txtpurac = RS!NAME & ""
         Else
          .txtpurac = Empty
        End If
        
        
        
                
        
        '.EXCREG.SetFocus
    End If
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
    txtFrDate.MaxDate = FEDT
    txtFrDate.MinDate = FSDT
  
    txtToDate.MaxDate = FEDT
    txtToDate.MinDate = FSDT
    Set TEMPRS = New Recordset
    Me.KeyPreview = True
End Sub

Private Sub lstBill_GotFocus()
  lstBill.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstBill_LostFocus()
 lstBill.BackColor = vbWhite
End Sub





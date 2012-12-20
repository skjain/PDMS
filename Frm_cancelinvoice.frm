VERSION 5.00
Begin VB.Form Frm_cancelinvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancel of Invoice Challan and Packing"
   ClientHeight    =   3330
   ClientLeft      =   5295
   ClientTop       =   2205
   ClientWidth     =   6255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6255
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame FRMOPT 
      Caption         =   "Select Option"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton OPTINV 
         Caption         =   "Invoice Challan and Packing Slip"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   6015
      Begin VB.TextBox LBLNET 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox LBLQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtDVCD 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4365
      End
      Begin VB.TextBox TXTDAYBOK 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4365
      End
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdgen 
         Caption         =   "&Show"
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Sale Day Book"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Sale Bill No."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Division Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Frm_cancelinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CUR_DVCD As String
Dim CUR_DBCD As String
Dim CUR_PCOD As String
Dim CUR_CRAC As String
Dim CUR_DLPT As String
Dim CHG_PCOD As String
Dim CHG_CRAC As String
Dim CHG_DLAC As String
Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub cmddelete_Click()
  On Error GoTo LAST
  If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM DIVMST WHERE NAME='" & txtDVCD & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
      If RS.EOF Then
        MsgBox "Invalid Division Name", vbCritical
        txtDVCD.SetFocus
        Exit Sub
      End If
      CUR_DVCD = RS!code
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM DAYBOK WHERE NAME='" & TXTDAYBOK & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & CUR_DVCD & "' AND VTYP='SAL'", CN, adOpenDynamic, adLockOptimistic
      If RS.EOF Then
        MsgBox "Invalid DAYBOOK Name", vbCritical
        TXTDAYBOK.SetFocus
        Exit Sub
      End If
      CUR_DBCD = RS!dbcd
  Dim AYS
  AYS = MsgBox("Are You sure to delete the data ? ", vbYesNo)
  If AYS = vbYes Then
      On Error GoTo LAST
      
      Dim SQL As String
      Dim SAL_COMP
      Dim sal_vtyp
      Dim sal_srno
      SAL_COMP = Empty
      sal_vtyp = Empty
      sal_srno = Empty
      Dim mstrst As New ADODB.Recordset
      SQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND UNIT='" & UNCD & "' AND DVCD='" & CUR_DVCD & "' AND DBCD='" & CUR_DBCD & "' AND VBNO='" & TXTVBNO.Text & "'"
      Set RS = New ADODB.Recordset
      Set mstrst = New ADODB.Recordset
      If RS.State = 1 Then RS.Close
      RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
      If RS.EOF Then
        
        MsgBox "Invalid Bill No."
        TXTVBNO.SetFocus
        Exit Sub
      End If
      LBLQTY = RS!TQTY
      LBLNET = RS!BNET
      SAL_COMP = RS!COMP
      sal_vtyp = RS!VTYP
      sal_srno = RS!SRNO
      CN.BeginTrans
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND RTYP='" & sal_vtyp & "' AND RSRN='" & sal_srno & "'", CN, adOpenDynamic, adLockOptimistic
      Do While Not RS.EOF
       Dim BOXREG As New ADODB.Recordset
       Set BOXREG = New ADODB.Recordset
       Dim DSP_VTYP As String
       Dim DSP_SRNO As String
       DSP_VTYP = RS!VTYP
       DSP_SRNO = RS!SRNO
       If BOXREG.State = 1 Then BOXREG.Close
       BOXREG.Open "SELECT * FROM BOXREG WHERE COMP='" & compPth & "' AND DTYP='" & DSP_VTYP & "' AND DSRN='" & DSP_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
       Do While Not BOXREG.EOF
        Dim pack_vtyp As String
        Dim pack_srno As String
        Dim PACK_ICOD As String
        Dim PACK_LTNO As String
        Dim PACK_GRAD As String
        Dim PACK_TWST As String
        Dim pack_nwgt As Double
        
        pack_vtyp = BOXREG!VTYP
        pack_srno = BOXREG!SRNO
        PACK_ICOD = BOXREG!ICOD
        PACK_LTNO = BOXREG!LTNO
        PACK_GRAD = BOXREG!grad
        PACK_TWST = BOXREG!twst
        pack_nwgt = BOXREG!nwgt
        Dim SPT As New ADODB.Recordset
        Set SPT = New ADODB.Recordset
        If SPT.State = 1 Then SPT.Close
        SPT.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND VTYP='" & pack_vtyp & "' AND SRNO='" & pack_srno & "' AND ICOD='" & PACK_ICOD & "' AND LTNO='" & PACK_LTNO & "' AND GRAD='" & PACK_GRAD & "' AND TWST='" & PACK_TWST & "'", CN, adOpenDynamic, adLockOptimistic
        If Not SPT.EOF Then
          SPT!QNTY = SPT!QNTY - pack_nwgt
          SPT!PCES = SPT!PCES - 1
          SPT.Update
        End If
        If SPT.State = 1 Then SPT.Close
        SPT.Open "SELECT * FROM PURTRAN WHERE COMP='" & compPth & "' AND VTYP='" & pack_vtyp & "' AND SRNO='" & pack_srno & "' AND ICOD='" & PACK_ICOD & "' AND LTNO='" & PACK_LTNO & "' AND GRAD='" & PACK_GRAD & "' AND TWST='" & PACK_TWST & "'", CN, adOpenDynamic, adLockOptimistic
        If Not SPT.EOF Then
          SPT!QNTY = SPT!QNTY - pack_nwgt
          SPT!PCES = SPT!PCES - 1
          SPT.Update
        End If
        If SPT.State = 1 Then SPT.Close
        SPT.Open "SELECT * FROM SPMAIN WHERE COMP='" & compPth & "' AND VTYP='" & pack_vtyp & "' AND SRNO='" & pack_srno & "'", CN, adOpenDynamic, adLockOptimistic
        If Not SPT.EOF Then
          SPT!TQTY = SPT!TQTY - pack_nwgt
          SPT!TPCS = SPT!TPCS - 1
          SPT.Update
        End If
        BOXREG!RECSTAT = "D"
        BOXREG.Update
        BOXREG.MoveNext
        
       Loop
       CN.Execute "DELETE FROM BOXREG WHERE RECSTAT='D'"
       CN.Execute "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND VTYP='PPF' AND QNTY=0 AND PCES=0"
       CN.Execute "DELETE FROM PURTRAN WHERE COMP='" & compPth & "' AND VTYP='PPF' AND QNTY=0 AND PCES=0"
       CN.Execute "DELETE FROM SPMAIN WHERE COMP='" & compPth & "' AND VTYP='PPF' AND TQTY=0 AND TPCS=0"
       CN.Execute "DELETE FROM SPMAIN WHERE COMP='" & compPth & "' AND VTYP='" & DSP_VTYP & "' AND SRNO='" & DSP_SRNO & "'"
       CN.Execute "DELETE FROM ORDTRN WHERE COMP='" & compPth & "' AND VTYP='" & DSP_VTYP & "' AND SRNO='" & DSP_SRNO & "'"
       RS.MoveNext
      Loop
      CN.Execute "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND RTYP='" & sal_vtyp & "' AND RSRN='" & sal_srno & "'"
      CN.Execute "DELETE FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='" & sal_vtyp & "' AND SRNO='" & sal_srno & "'"
      CN.Execute "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='" & sal_vtyp & "' AND SRNO='" & sal_srno & "'"
      CN.CommitTrans
  End If
  MsgBox "Record Delete Successfuly"
  Call ClsData(Frm_cancelinvoice)
  txtDVCD.SetFocus
  Exit Sub
LAST:
  MsgBox Err.Description

  CN.RollbackTrans
End Sub

Private Sub cmdgen_Click()
  On Error GoTo LAST
  Dim SQL As String
  Dim mstrst As New ADODB.Recordset
  SQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND UNIT='" & UNCD & "' AND DVCD='" & CUR_DVCD & "' AND DBCD='" & CUR_DBCD & "' AND VBNO='" & TXTVBNO.Text & "'"
  Set RS = New ADODB.Recordset
  Set mstrst = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Bill No."
    TXTVBNO.SetFocus
    Exit Sub
  End If
  LBLQTY = RS!TQTY
  LBLNET = RS!BNET
  If mstrst.State = 1 Then mstrst.Close
  mstrst.Open "select * from accmst where code='" & RS!PCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If mstrst.EOF Then
    MsgBox "Invalid Party Name"
    TXTVBNO.SetFocus
    Exit Sub
  End If
  cmddelete.SetFocus
  Exit Sub
LAST:
  MsgBox Err.Description
  
  Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If ActiveControl.Name = "txtDVCD" Or ActiveControl.Name = "TXTDAYBOK" Then Exit Sub
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
End Sub

Private Sub TXTDAYBOK_GotFocus()
TXTDAYBOK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDAYBOK_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And Trim(TXTDAYBOK) = Empty) Or KeyCode = vbKeyF2 Then
    TXTDAYBOK.Text = SearchList1("SELECT DBCD,NAME FROM DAYBOK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & CUR_DVCD & "' AND VTYP='SAL'", 0, TXTDAYBOK.Text, "SELECT SALE DAYBOOK FROM LIST")
    CUR_DBCD = Key
  End If
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TXTDAYBOK_LostFocus()
 TXTDAYBOK.BackColor = vbWhite
End Sub

Private Sub txtDVCD_GotFocus()
txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And Trim(txtDVCD) = Empty) Or KeyCode = vbKeyF2 Then
    txtDVCD.Text = SearchList1("SELECT CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, txtDVCD.Text, "SELECT DIVISION FROM LIST")
    CUR_DVCD = Key
  End If
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDVCD_LostFocus()
txtDVCD.BackColor = vbWhite
End Sub


Private Sub TXTVBNO_GotFocus()
 TXTVBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTVBNO_LostFocus()
TXTVBNO.BackColor = vbWhite

End Sub

Private Sub TXTVBNO_Validate(cancel As Boolean)
  Call cmdgen_Click
End Sub

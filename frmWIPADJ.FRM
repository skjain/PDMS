VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "welchbutton.ocx"
Begin VB.Form FRMWIPADJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W.I.P. Adjustment"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8805
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   5280
      Width           =   8655
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Add"
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
         Image           =   "frmWIPADJ.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "E&dit"
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
         Image           =   "frmWIPADJ.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   3720
         TabIndex        =   22
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Delete"
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
         Image           =   "frmWIPADJ.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Save"
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
         Image           =   "frmWIPADJ.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frmWIPADJ.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         Image           =   "frmWIPADJ.frx":1CAA
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   8655
      Begin VB.TextBox TXTMACHINE 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox TXTRMRK 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   3600
         Width           =   6975
      End
      Begin VB.TextBox TXTQTY 
         Height          =   285
         Left            =   5760
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TXTITEM 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   3135
      End
      Begin WelchButton.lvButtons_H CMDOK 
         Height          =   375
         Left            =   7320
         TabIndex        =   7
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Add"
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
         Image           =   "frmWIPADJ.frx":20FC
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDITMDEL 
         Height          =   375
         Left            =   7320
         TabIndex        =   17
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Remove"
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
         Image           =   "frmWIPADJ.frx":2496
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid ITMFLEX 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Machine :"
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
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Line Line4 
         X1              =   2280
         X2              =   2280
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7200
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label6 
         Caption         =   "Remark :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Line Line3 
         X1              =   7200
         X2              =   7200
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line2 
         X1              =   5640
         X2              =   5640
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Shape Shape1 
         Height          =   1095
         Left            =   120
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TXTDVCD 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6495
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
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
         Format          =   88801281
         CurrentDate     =   39339
      End
      Begin VB.Label Label3 
         Caption         =   "Adjustment No. :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Date      :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Division :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FRMWIPADJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public M_DBCD As String
Public M_DVCD As String
Public M_DVNM As String
Public M_SRNO As String
Dim SAVEFLAG As Boolean
Dim ROWNO As Long
Dim SWITCH As Boolean
'-------------------------------------------------------------------------------------------
' FORM EVENTS
'-------------------------------------------------------------------------------------------

Private Sub CMDCANCEL_Click()
 
  ClsData (Me)
  txtDVCD = Empty
  ITMFLEX.Clear
  ITMFLEX.Rows = 2
  btn_sts (True)
  Call SetFlex
  cmdAdd.SetFocus
  M_SRNO = Empty
  cmdOk.Caption = "&Add"
  SWITCH = False
End Sub


Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("0019", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If

  SAVEFLAG = False
  M_SRNO = Empty
  
  M_DESC = Empty
  Key = Empty
  NEW_VISIBLE = False
  btn_sts (False)
  FRMWIPADJLIST.Show 1
  If Not M_SRNO = Empty Then
     Dim AYS
     AYS = MsgBox("Are you sure to delete the Adjustment Slip ", vbYesNo)
     If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='WIP' AND SRNO='" & M_SRNO & "'"
        CN.CommitTrans
     End If
  End If
    
  Call CMDCANCEL_Click

End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("0019", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  SAVEFLAG = False
  M_SRNO = Empty
  
  M_DESC = Empty
  Key = Empty
  NEW_VISIBLE = False

  'TXTTODIV = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
     
 ' frmStoreIssList.Show 1
   FRMWIPADJLIST.Show 1
  
  If ITMFLEX.TextMatrix(1, 2) <> Empty Then Call ITMFLEX_Click
    
  If cmdSave.Enabled Then
    cmdDelete.Enabled = True
  End If
  
  If ITMFLEX.TextMatrix(1, 2) <> Empty Then Call ITMFLEX_Click
  Call btn_sts(False)
  
  If TXTMACHINE.Enabled = True Then
    TXTMACHINE.SetFocus
  End If
  
  Exit Sub
  
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub CMDITMDEL_Click()
Dim CURSOR As Long
Dim J As Long

For J = ROWNO To ITMFLEX.Rows - 2
 ITMFLEX.TextMatrix(J, 0) = ITMFLEX.TextMatrix(J + 1, 0)
 ITMFLEX.TextMatrix(J, 1) = ITMFLEX.TextMatrix(J + 1, 1)
 ITMFLEX.TextMatrix(J, 2) = ITMFLEX.TextMatrix(J + 1, 2)
 ITMFLEX.TextMatrix(J, 3) = ITMFLEX.TextMatrix(J + 1, 3)
 ITMFLEX.TextMatrix(J, 4) = ITMFLEX.TextMatrix(J + 1, 4)
Next J

ITMFLEX.Rows = ITMFLEX.Rows - 1
Call CLEARDATA

If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
        TXTMACHINE.SetFocus
    Else
        
    End If

SWITCH = False
If TXTMACHINE.Enabled Then TXTMACHINE.SetFocus
cmdOk.Caption = "&Add"
TXTMACHINE.Enabled = False

End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

 If CHKSAVEDATA = True Then
    Exit Sub
 End If
  
'GENERATE SR. NO
 If M_SRNO = Empty Then
    M_SRNO = pubGenSrNoSTR(TXTVBDT, "WIP")
 End If
    
 If SAVEFLAG = True Then
    txtvbno = GenVNO("WIP", M_DBCD)
 End If
    
 Call SAVEREC
 
 If SAVEFLAG = True Then
    MsgBox "Your Adjustment Slip No. is " + txtvbno.Text
 End If
    Call CMDCANCEL_Click
 Exit Sub
    
LAST:
    MsgBox ERR.Description
    If RS.State = 1 Then
        RS.CancelUpdate
    End If
    CN.RollbackTrans
    Exit Sub

End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 Me.BackColor = RGB(RED, GREEN, BLUE)
 
End Sub

Private Sub Form_Load()
 Call CenterChild(frm_Main, Me)
 FIFOREQ = "Y"
 Me.KeyPreview = True
 Me.Tag = zoomflag
 M_DBCD = "000001"
 If Not zoomflag = True Then
    M_SRNO = Empty
 End If
 'M_DVCD = "000001"
 TXTVBDT = Date
 TXTVBDT.MaxDate = FEDT
 TXTVBDT.MinDate = FSDT
 Call SetFlex
 btn_sts (True)
' CMDITMDEL.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If UCase(ActiveControl.NAME) = "TXTRMRK" And KeyAscii = vbKeyReturn Then cmdSave.SetFocus: Exit Sub
 If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
 End If
End Sub
'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
' BUTTON EVENTS
'-------------------------------------------------------------------------------------------
Private Sub cmdAdd_Click()
    zoomflag = False
    btn_sts (False)
    M_SRNO = Empty
    txtvbno = GenVNO("WIP", M_DBCD)
    SAVEFLAG = True
    txtDVCD.SetFocus
    End Sub

Private Sub CMDOK_Click()
 Dim INDEX As Long
 
 If Not SWITCH Then
      ROWNO = ITMFLEX.Rows - 1
 End If
 
 If CheckData(ROWNO) Then Exit Sub
 
    ITMFLEX.TextMatrix(ROWNO, 0) = Trim(TXTMACHINE)
    ITMFLEX.TextMatrix(ROWNO, 1) = Trim(txtitem)
    ITMFLEX.TextMatrix(ROWNO, 2) = Trim(txtQty)
    ITMFLEX.TextMatrix(ROWNO, 3) = GetMachineCode(txtDVCD.Tag, TXTMACHINE)
    ITMFLEX.TextMatrix(ROWNO, 4) = GetCode("ITMMST", txtitem, "NAME", "CODE")
    
    If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
            TXTMACHINE.SetFocus
    Else
    TXTRMRK.SetFocus
   
    End If
    Call CLEARDATA
    cmdOk.Caption = "&Add"
    SWITCH = False
End Sub
'-------------------------------------------------------------------------------------------
' LOCAL PROCEDURE
'-------------------------------------------------------------------------------------------
Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdOk.Enabled = Not Yes
    CMDITMDEL.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    TXTMACHINE.Enabled = Not Yes
    TXTVBDT.Enabled = Not Yes
    txtitem.Enabled = Not Yes
    TXTMACHINE.Enabled = Not Yes
    txtQty.Enabled = Not Yes
    TXTRMRK.Enabled = Not Yes
    txtDVCD.Enabled = Not Yes
    txtvbno.Enabled = Not Yes
End Sub
'-------------------------------------------------------------------------------------------

Private Sub ITMFLEX_Click()
   If ITMFLEX.Rows > 1 Then
    cmdOk.Caption = "Upd&ate"
    CMDITMDEL.Enabled = True
    ROWNO = ITMFLEX.ROW
    TXTMACHINE = Trim(ITMFLEX.TextMatrix(ROWNO, 0))
    txtitem = Trim(ITMFLEX.TextMatrix(ROWNO, 1))
    txtQty = Val(ITMFLEX.TextMatrix(ROWNO, 2))
    ITMFLEX.TextMatrix(ROWNO, 3) = Empty
    ITMFLEX.TextMatrix(ROWNO, 4) = Empty
    SWITCH = True
  End If
    
   If Val(ITMFLEX.ROW) > 0 Then
      If TXTMACHINE.Enabled Then TXTMACHINE.SetFocus
   End If
   
End Sub

'-------------------------------------------------------------------------------------------
' CODE FOR CURSOR POSITION ON MODULE
'-------------------------------------------------------------------------------------------

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}":
End Sub


Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
If KeyCode = vbKeyF2 Or (Trim(txtitem) = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False
   If ItemSearchField = 0 Then
      M_DESC = txtitem.Text
      
            
   Else
      M_DESC = txtitem.Text
   End If
   Key = Empty
   If SAVEFLAG Then
      txtitem.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ITMMST", 0, txtitem.Text, "SELECT ITEM FROM LIST")
   Else
      txtitem.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ITMMST", 0, txtitem.Text, "SELECT ITEM FROM LIST")
   End If
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            txtitem.Text = ""
            frm_Item.Show
        Else
            txtitem.Tag = Key
          
        End If
    Else
    End If
    Me.KeyPreview = True
End Sub
Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & UNCD & "' AND CODE <> '000001' AND RECSTAT='A' ", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
        
        
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
    If KeyCode = vbKeyReturn And txtDVCD <> Empty Then txtDVCD.SetFocus
End Sub

Private Sub txtDVCD_LostFocus()
txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
 TXTMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub txtMACHINE_LostFocus()
 TXTMACHINE.BackColor = vbWhite
End Sub
Private Sub txtItem_GotFocus()
 txtitem.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub txtItem_LostFocus()
 txtitem.BackColor = vbWhite
End Sub

Private Sub TXTQTY_GotFocus()
 txtQty.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTQTY_LostFocus()
 txtQty.BackColor = vbWhite
End Sub


Private Sub TXTRMRK_GotFocus()
 TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTRMRK_LostFocus()
 TXTRMRK.BackColor = vbWhite
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTMACHINE = Empty) Then
        If txtDVCD = Empty Then txtDVCD.Enabled = True: txtDVCD.SetFocus: Exit Sub
        NEW_VISIBLE = False
        
        M_DESC = Empty
        TXTMACHINE = SearchList1("Select TOP 20 Code,Name From MACMST WHERE COMP='" & compPth & _
        "' AND UNIT='" & UNCD & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Machine")
        TXTMACHINE.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        TXTMACHINE = Empty
        TXTMACHINE.Tag = Empty
    End If
    
     If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub SetFlex()
  ITMFLEX.Clear
  ITMFLEX.ColWidth(0) = 2500
  ITMFLEX.ColWidth(1) = 4000
  ITMFLEX.ColWidth(2) = 1500
  ITMFLEX.ColWidth(3) = 0
  ITMFLEX.ColWidth(4) = 0

  
  ITMFLEX.Clear
  ITMFLEX.TextMatrix(0, 0) = "Machine Name"
  ITMFLEX.TextMatrix(0, 1) = "Item Description"
  ITMFLEX.TextMatrix(0, 2) = "Item Qty."
  ITMFLEX.TextMatrix(0, 3) = "Machine Code"
  ITMFLEX.TextMatrix(0, 4) = "Itemcode"
  
  
  ITMFLEX.ColAlignment(0) = vbLeftJustify
  ITMFLEX.ColAlignment(1) = vbLeftJustify
  ITMFLEX.ColAlignment(2) = vbRightJustify
  ITMFLEX.ColAlignment(3) = vbRightJustify
  ITMFLEX.ColAlignment(4) = vbRightJustify
End Sub

Private Sub CLEARDATA()
        txtitem = Empty
        TXTMACHINE = Empty
        txtQty = Empty
        
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
    If Trim(txtitem) = Empty Then
        MsgBox "Please Select Items From List !!", vbInformation
        If txtitem.Enabled Then txtitem.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If Not IsNumeric(txtQty) Then
        MsgBox "Please Enter Valid Quantity !!", vbInformation, "Quantity Missing !!"
        If txtQty.Enabled Then txtQty.SetFocus
        CheckData = True
        Exit Function
    End If
                
    
    For INDEX = 1 To ITMFLEX.Rows - 1
        If Trim(ITMFLEX.TextMatrix(INDEX, 0)) = TXTMACHINE And Trim(ITMFLEX.TextMatrix(INDEX, 1)) = txtitem And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
           MsgBox "Invalid Item Detail"
           If txtitem.Enabled Then TXTMACHINE.SetFocus
           CheckData = True
           Exit Function
        End If
    Next INDEX
    
End Function


Private Function CHKSAVEDATA() As Boolean
If txtDVCD = Empty Then
  MsgBox "Enter Source Division then Save"
  CHKSAVEDATA = True
  If txtDVCD.Enabled Then txtDVCD.SetFocus
  Exit Function
End If


If ITMFLEX.TextMatrix(1, 0) = Empty Then
  MsgBox "Enter Data then Save"
  CHKSAVEDATA = True
  txtitem.Enabled = True
  txtitem.SetFocus
  Exit Function
End If

End Function

Private Sub SAVEREC()
  
  On Error GoTo LAST
  Dim SQL As String
  Dim CSHD As String
  
  Dim SAVDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  
  Set SAVDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
      
  
  CN.BeginTrans
  Call DELETEISS
  SQL = Empty
  
    
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND VTYP='WIP' AND SRNO='" & M_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
  
  Dim AI As String
  Dim BQ As Double
  Dim I As Long
  Dim DVCOD As String
  DVCOD = GetDivCode(txtDVCD)
    
  I = 1
  Dim FIFORATE As Double
  
  For I = 1 To ITMFLEX.Rows - 1
    If ITMFLEX.TextMatrix(I, 0) <> Empty And ITMFLEX.TextMatrix(I, 1) <> Empty Then
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "WIP"
    SAVDAT!SRNO = M_SRNO
    SAVDAT!SRCH = I
    SAVDAT!VBNO = Trim(txtvbno.Text)
    SAVDAT!chln = txtvbno.Text
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!ICOD = GetCode("ITMMST", ITMFLEX.TextMatrix(I, 1), "NAME", "CODE")
    SAVDAT!CSHD = ""
    SAVDAT!CHEAD = ""
    SAVDAT!PCES = 0
    SAVDAT!QNTY = Abs(Val(ITMFLEX.TextMatrix(I, 2)))
    SAVDAT!QORP = "Q"
    SAVDAT![User] = cUName
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    If Val(ITMFLEX.TextMatrix(I, 2)) > 0 Then
    SAVDAT!OPER = "-"
     Else
     SAVDAT!OPER = "+"
     End If
    SAVDAT!PCOD = GetMachineCode(DVCOD, ITMFLEX.TextMatrix(I, 0))
    SAVDAT!DVCD = DVCOD
    
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
    SAVDAT!ITEMRMRK = Trim(TXTRMRK)
    SAVDAT!EXTRA1 = nstr(Val(ITMFLEX.TextMatrix(I, 2)), 12, 3)
    SAVDAT.Update
  End If
  Next
 
 'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
    Call SetSRNO(txtvbno, "WIP", M_DBCD)
  End If
  
  
 
  '----------------------
  
  CN.CommitTrans
  Exit Sub
LAST:
 MsgBox ERR.Description
 Resume
 If SAVDAT.State = 1 Then
   SAVDAT.CancelUpdate
   SAVDAT.Close
 End If
 CN.RollbackTrans
End Sub


Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "ISS"
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = 0
  DLYSTA!VBNO = txtvbno & ""
  DLYSTA!AMNT = 0
  DLYSTA!CUSR = cUName
  If SAVEFLAG = True Then
    DLYSTA!ACTN = "E"
   Else
    DLYSTA!ACTN = "M"
  End If
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub DELETEISS()
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND VTYP='WIP' AND SRNO='" & M_SRNO & "'"
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

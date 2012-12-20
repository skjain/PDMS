VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmPalletDeletion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pallet Deletion"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6794.673
   ScaleMode       =   0  'User
   ScaleWidth      =   16939.7
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   6795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11986
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   12632319
      BackColor       =   16777215
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
      Begin MSComctlLib.ListView lstOrdn 
         Height          =   6015
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   10485760
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order No."
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pallet No."
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qnty."
            Object.Width           =   2540
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   5160
         TabIndex        =   5
         Top             =   5520
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
         Image           =   "frmPalletDeletion.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5160
         TabIndex        =   6
         Top             =   6120
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
         Image           =   "frmPalletDeletion.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         Height          =   780
         Left            =   4920
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Deletion"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pallet "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carton Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label LBLDIV 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmPalletDeletion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DIVCODE As String

Private Sub Form_Activate()
If DIVCODE = Empty Or LBLDIV = Empty Then
  Unload Me
End If

If lstOrdn.ListItems.COUNT > 0 Then
  lstOrdn.SetFocus
End If

Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
Me.Left = 50: Me.KeyPreview = True
  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  DIVCODE = Empty
  If DIVCODE = Empty Then
    LBLDIV.Caption = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
  
  Call FillDetails
  
End Sub

Private Sub FillDetails()

Dim PKGRS As ADODB.Recordset
Set PKGRS = New ADODB.Recordset

Dim SRCRS As ADODB.Recordset
Set SRCRS = New ADODB.Recordset
If SRCRS.State = 1 Then SRCRS.Close
SRCRS.Open "SELECT DISTINCT ORDN FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND RECSTAT<>'D' AND OFLG<>'Y' AND FIN_APRV = 'O' AND DOQTY > 0", CN, adOpenDynamic, adLockOptimistic
           
lstOrdn.ListItems.Clear
Do While Not SRCRS.EOF
   
   If PKGRS.State = 1 Then PKGRS.Close
   PKGRS.Open "SELECT ORDN,PLTNO,SUM(NTWGT) AS QTY FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND (VTYP='PPF' OR VTYP='OPN') AND ORDN='" & SRCRS!ORDN & _
              "' AND RECSTAT<>'D' GROUP BY ORDN,PLTNO", CN, adOpenDynamic, adLockOptimistic
   Do While Not PKGRS.EOF
        Set Item = lstOrdn.ListItems.ADD
        Item.Text = Trim(PKGRS!ORDN & "")
        Item.SubItems(1) = Trim(PKGRS!PLTNO & "")
        Item.SubItems(2) = nstr(Val(PKGRS!QTY), 8, 3)
   PKGRS.MoveNext
   Loop
   PKGRS.Close
   
SRCRS.MoveNext
Loop
SRCRS.Close

End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo LAST
Dim Index As Long
Dim FLAG As Boolean  'OK
Dim SQL As String

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset

Dim PALLETNO As String, ORDN As String


'TO FIND ATLEAST ONE SELECTION
FLAG = False
For Index = 1 To lstOrdn.ListItems.COUNT
  If lstOrdn.ListItems(Index).Checked = True Then: FLAG = True: Exit For
Next

If FLAG = False Then Exit Sub

If MsgBox("Are You Sure ? Want To Delete Record ?", vbQuestion + vbYesNo, "Delete Location") = vbYes Then

    CN.BeginTrans
    
    'FIRST SET ALL BOX AS DISPATCH ON BASIS OF PALLET NO WHICH IS SELECTED BY USER
    For Index = 1 To lstOrdn.ListItems.COUNT
      If lstOrdn.ListItems(Index).Checked = True Then
         PALLETNO = lstOrdn.ListItems(Index).SubItems(1)
         
         If CHKRS.State = 1 Then CHKRS.Close
         CHKRS.Open "SELECT DISTINCT CHLN,DBCD FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND (VTYP='PPF' OR VTYP='OPN') AND PLTNO ='" & PALLETNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
         Do While Not CHKRS.EOF
            CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                       "' AND VTYP='PPF' AND CHLN ='" & Trim(CHKRS!chln & "") & _
                       "' AND DBCD ='" & Trim(CHKRS!dbcd & "") & "' AND RECSTAT<>'D'"
         CHKRS.MoveNext
         Loop
         CHKRS.Close
               
         SQL = "UPDATE BOXREGISTER SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND (VTYP='PPF' OR VTYP='OPN') AND PLTNO ='" & PALLETNO & "' AND RECSTAT<>'D'"
         
         CN.Execute SQL
         
      End If
    Next
    
    
    'SET ORDER DISPATCH QTY REVERSAL
    For Index = 1 To lstOrdn.ListItems.COUNT
      If lstOrdn.ListItems(Index).Checked = True Then
         PALLETNO = lstOrdn.ListItems(Index).SubItems(1)
         ORDN = lstOrdn.ListItems(Index).Text
         
         If CHKRS.State = 1 Then CHKRS.Close
         CHKRS.Open "SELECT ICOD,GRAD,SUM(NTWGT) AS QNTY FROM BOXREGISTER WHERE COMP='" & compPth & _
                    "' AND UNIT='" & UNCD & "' AND (VTYP='PPF' OR VTYP='OPN') AND PLTNO ='" & PALLETNO & _
                    "' AND ORDN='" & ORDN & "' AND RECSTAT='D' GROUP BY ICOD,GRAD", CN, adOpenDynamic, adLockOptimistic
         Do While Not CHKRS.EOF
                                
          SQL = "UPDATE ORDMAN SET DOQTY = DOQTY - " & Val(CHKRS!QNTY & "") & _
               " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DCOD='" & DIVCODE & _
               "' AND ORDN = '" & ORDN & "' AND ICOD = '" & Trim(CHKRS!ICOD & "") & _
               "' AND TRCD='" & Trim(CHKRS!grad & "") & "'"
                    
         CN.Execute SQL
         
         CHKRS.MoveNext
         Loop
         CHKRS.Close
                  
         CN.Execute "DELETE FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                    "' AND (VTYP='PPF' OR VTYP='OPN') AND PLTNO ='" & PALLETNO & _
                    "' AND ORDN='" & ORDN & "' AND RECSTAT='D' "
         
      End If
    Next
    
    Call DAILYSTATUS("PPF", "", "", 0, PALLETNO, 0, cUName, "D", Now, Now)
    
    CN.CommitTrans
    
    Call FillDetails

End If

Exit Sub
LAST:
MsgBox ERR.Description
CN.RollbackTrans
Exit Sub
End Sub


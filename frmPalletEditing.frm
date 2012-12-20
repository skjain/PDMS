VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmPalletEditing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Updation of Pallet Against Order"
   ClientHeight    =   6765
   ClientLeft      =   375
   ClientTop       =   1110
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Packing"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6794.673
   ScaleMode       =   0  'User
   ScaleWidth      =   11535.64
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrTool 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1920
         Top             =   0
      End
      Begin VB.Label lblToolTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   6840
   End
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   6795
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11986
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   12632256
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
      Begin VB.TextBox TXTPALLET 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox NETWGT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   9
         Tag             =   "0"
         Top             =   6360
         Width           =   975
      End
      Begin VB.TextBox NETCOPS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   8760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   7
         Tag             =   "0"
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox NETBOXES 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   5
         Tag             =   "0"
         Top             =   6360
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   5205
         Left            =   45
         TabIndex        =   2
         Top             =   960
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9181
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColor       =   -2147483628
         BackColorBkg    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Update"
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
         Image           =   "frmPalletEditing.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Pallet No."
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
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Weight"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9960
         TabIndex        =   8
         Top             =   6195
         Width           =   1065
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total BOXES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7320
         TabIndex        =   4
         Top             =   6195
         Width           =   1170
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total COPS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8640
         TabIndex        =   6
         Top             =   6195
         Width           =   1050
      End
      Begin VB.Label LBLDESC1 
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
         Left            =   1920
         TabIndex        =   17
         Top             =   120
         Width           =   3375
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name :"
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
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label LBLDESC2 
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
         Left            =   7560
         TabIndex        =   15
         Top             =   120
         Width           =   3615
      End
      Begin VB.Shape BORDER2 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label LBLHEADING2 
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Station :"
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
         Left            =   5760
         TabIndex        =   14
         Top             =   120
         Width           =   1695
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
         TabIndex        =   11
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmPalletEditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ERROROCCUR As Boolean
Dim LOAD As String
Dim DIVCODE As String
Dim LSPKGCOD As String
Dim M_DBCD As String
Dim PKGNGCD As String
Dim MCCD As String
Dim LOCCOD As String
Dim RETURNABLE As String
Dim GRADE As String
Dim SUBGRADE As String
Dim CHALLAN As String
Dim PALETNO As String
'---
Dim SAVEFLAG As Boolean
Dim ROWNO As Long
Dim SWITCH As Boolean
Dim SQL As String
Dim COUNTER As Long
Dim M_PCOD As String
Dim LASTBOXN As String
Dim FINITMCOD As String
Dim Emptycell As Boolean
Dim NETQTY As Double

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       FLEX.SetFocus
       FLEX.COL = 0
       FLEX.ROW = 1
    End If
End Sub

Private Sub cmdSave_Click()

CN.BeginTrans

Call ReduceOrder

If Not IsDataOK Then
   CN.RollbackTrans
   Exit Sub
End If

Call SavePallet
             
'ERROROCCUR = False

If ERROROCCUR Then
   CN.RollbackTrans
   Exit Sub
End If

End Sub

Private Sub FLEX_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case FLEX.COL
   
Case 3
    Dim FILTER As String
    Dim FILTER2 As String
    If Len(Trim(FLEX.TextMatrix(FLEX.ROW, 0))) = 10 Then
       If RS.State = 1 Then RS.Close
       RS.Open "SELECT DISTINCT TRCD FROM ORDMAN WHERE COMP='" & compPth & _
               "' AND UNIT='" & UNCD & "' AND ORDN ='" & Trim(FLEX.TextMatrix(FLEX.ROW, 0)) & "'", CN, adOpenDynamic, adLockOptimistic
       Do While Not RS.EOF
          If FILTER <> Empty Then FILTER = FILTER & ","
          FILTER = FILTER & "'" & Trim(RS!TRCD & "") & "'"
          
          If FILTER2 <> Empty Then FILTER2 = FILTER2 & ","
          FILTER2 = FILTER2 & "'" & Trim(RS!ICOD & "") & "'"
                    
       RS.MoveNext
       Loop
       RS.Close
    End If
            
    If KeyCode = vbKeyF2 Or (KeyCode = 13 And FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty) Then
        NEW_VISIBLE = False: M_DESC = Empty: Key = Empty
        SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND DVCD='" & DIVCODE & "' AND SHCD IN (" & FILTER & ") AND FICD IN (" & FILTER2 & ") "
        FLEX.TextMatrix(FLEX.ROW, 3) = SearchList(SQL)
        FLEX.TextMatrix(FLEX.ROW, 4) = FindFinishItem(FLEX.TextMatrix(FLEX.ROW, 3))
        'FIND SHADE
        If FLEX.TextMatrix(FLEX.ROW, 3) <> Empty Then
            If RS.State = 1 Then RS.Close
            RS.Open "SELECT SHCD FROM TXULOT WHERE COMP='" & compPth & _
                   "' AND UNIT='" & UNCD & "' AND DVCD ='" & DIVCODE & "' AND LTNO='" & FLEX.TextMatrix(FLEX.ROW, 3) & "'", CN, adOpenDynamic, adLockOptimistic
            If Not RS.EOF Then
               FLEX.TextMatrix(FLEX.ROW, 5) = GetCode("GRDMST", RS!SHCD & "", "CODE", "GRAD")
            End If
        End If
        '-------------------------------
    End If
    
 Case 6
 
   If KeyCode = vbKeyF2 Or (KeyCode = 13 And FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty) Then
      NEW_VISIBLE = False:  M_DESC = Empty:   Key = Empty
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = SearchList1("SELECT  TOP 20 CODE,NAME FROM MACMST WHERE COMP='" & compPth & _
                      "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'", 0, "", "List of Machine Name")
   ElseIf KeyCode = vbKeyDelete Then
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty
   End If
    
    
End Select

End Sub

Private Sub Flex_LeaveCell()
  FLEX.CellBackColor = vbWhite
  Dim i As Long
  
  For i = 1 To FLEX.Rows - 1
  'FIND SHADE
   If FLEX.TextMatrix(i, 3) <> Empty Then
      FLEX.TextMatrix(i, 4) = FindFinishItem(FLEX.TextMatrix(i, 3))
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT SHCD FROM TXULOT WHERE COMP='" & compPth & _
             "' AND UNIT='" & UNCD & "' AND DVCD ='" & DIVCODE & "' AND LTNO='" & FLEX.TextMatrix(i, 3) & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
         FLEX.TextMatrix(i, 5) = GetCode("GRDMST", RS!SHCD & "", "CODE", "GRAD")
      End If
   End If
   '-------------------------------
  Next i
  
End Sub

Private Sub FLEX_LostFocus()
FLEX.CellBackColor = vbWhite
End Sub

Private Sub Form_Activate()
  If DIVCODE = Empty Or Trim(LBLDESC1.Caption) = "XXXXXXXXXX" Then
     MsgBox "Select Division For Packing."
     Unload Me
  End If
  
  If LSPKGCOD = Empty Or Trim(LBLDESC2.Caption) = "XXXXXXXXXX" Then
     MsgBox "Select Packing Station For Packing."
     Unload Me
  End If
  
  'For Raw Material Consumption Slip
  If CHALLAN = Empty Or PALETNO = Empty Then
     Unload Me
  End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me): Call ColorComponent(Me)
  
  ERROROCCUR = False
  
  Me.Left = 50: Me.KeyPreview = True
  SAVEFLAG = True
'-------DIVISION NAME
  M_DESC = Empty: Key = Empty:  NEW_VISIBLE = False:  DIVCODE = Empty
  LBLDESC1.Caption = Empty
  If DIVCODE = Empty Then
    LBLDESC1 = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A' AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
    If LBLDESC1 <> Empty Then DIVCODE = Key Else LBLDESC1 = "???????": Unload Me
  End If
    
 If PackingType(Key) = "L" Then MsgBox "Division Not Allowed Carton Packing.Check Configuration": LOAD = "N": GoTo JUMP
  
'-------PACKING STATION MASTER
M_DESC = Empty:  Key = Empty:  NEW_VISIBLE = False: LSPKGCOD = Empty
LBLDESC2 = SearchList1("SELECT TOP 20 CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, "", "SELECT PACKING STATION FROM MASTER LIST")
If Key = Empty Then Exit Sub
LSPKGCOD = Key
'---------------------------
Call setflexhead

'For Raw Material Consumption Slip
CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
'FOR PALLET NO.
PALETNO = GenPackSlipNo(LSPKGCOD, "LPNO")

'For Box No.
'BOXNO.Caption = GenPackSlipNo(LSPKGCOD)

COUNTER = 0
       
'TXTVBDT.MinDate = FSDT
Call SetLastDateForPacking


Call SetPackingType


JUMP:
End Sub


Private Function FindFinishItem(txtLTNo As String) As String
FindFinishItem = Empty
Dim RSITM As ADODB.Recordset: Set RSITM = New ADODB.Recordset
Dim FICD As String

If RSITM.State = 1 Then RSITM.Close
RSITM.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSITM.EOF Then FindFinishItem = RSITM!FICD
RSITM.Close

End Function

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
    
    PKTYPRS.MoveNext
Loop
    PKTYPRS.Close
End Sub

Private Function FindFinItemCode(INAM As String) As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND NAME ='" & INAM & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   FindFinItemCode = GRRS!CODE
Else
   FindFinItemCode = Empty
End If
GRRS.Close
End Function

Private Sub SetLastDateForPacking()
Dim DTRS As ADODB.Recordset
Set DTRS = New ADODB.Recordset

If DTRS.State = 1 Then DTRS.Close
DTRS.Open "SELECT IsNull(LSTPCKDT,'" & FSDT & "') AS LSTPCKDT FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & LSPKGCOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DTRS.EOF Then
   
End If
DTRS.Close
End Sub

Private Function IsBoxExistInUnit(BOXNUM As String) As Boolean
IsBoxExistInUnit = False

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND VBNO='" & BOXNUM & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
   IsBoxExistInUnit = True
End If
CHKRS.Close
End Function

Private Sub setflexhead()

    FLEX.TextMatrix(0, 0) = "Order No."
    FLEX.TextMatrix(0, 1) = "Pallet No."
    FLEX.TextMatrix(0, 2) = "Carton No."
    FLEX.TextMatrix(0, 3) = "Lot No."
    FLEX.TextMatrix(0, 4) = "Item Name"
    FLEX.TextMatrix(0, 5) = "Shade"
    FLEX.TextMatrix(0, 6) = "Machine Name"
    FLEX.TextMatrix(0, 7) = "Cops"
    FLEX.TextMatrix(0, 8) = "Grs Wgt"
    FLEX.TextMatrix(0, 9) = "Tare Wgt"
    FLEX.TextMatrix(0, 10) = "Net Wgt"
    
    FLEX.ColWidth(0) = 1300
    FLEX.ColWidth(1) = 1300
    FLEX.ColWidth(2) = 1300
    FLEX.ColWidth(3) = 1100
    FLEX.ColWidth(4) = 0
    FLEX.ColWidth(5) = 2100
    FLEX.ColWidth(6) = 2100
    FLEX.ColWidth(7) = 600
    FLEX.ColWidth(8) = 1100
    FLEX.ColWidth(9) = 1100
    FLEX.ColWidth(10) = 1100
    
    FLEX.ColAlignment(8) = 1
    FLEX.ColAlignment(9) = 1
    FLEX.ColAlignment(10) = 1
    
End Sub

Private Sub FLEX_EnterCell()
  FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
  Emptycell = True
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
  On Error GoTo LAST
  
  Dim ALLOW_KEY As Boolean
  Dim FWD_COL As Boolean
  Dim ENTER_PRESS As Boolean
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  FWD_COL = False
  ALLOW_KEY = False
  
  If FLEX.COL = 8 Or FLEX.COL = 9 Then
    If InStr(1, FLEX.TextMatrix(FLEX.ROW, FLEX.COL), ".") > 0 And KeyAscii = 46 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
   
  If Emptycell = True And (Not KeyAscii = 13) Then
     If FLEX.COL <> 10 Then
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty
        Call CalculateTotal
     End If
     Emptycell = False
  End If
  
  Select Case FLEX.COL
   Case 0
    ALLOW_KEY = False
   Case 3
    ALLOW_KEY = False
   Case 5
    ALLOW_KEY = False
   Case 7
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
    
   Case 8
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
    
   Case 9
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If

  End Select
  
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
   Else
    ENTER_PRESS = False
  End If
  
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))
    If lnth > 0 Then
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Mid(FLEX.TextMatrix(FLEX.ROW, FLEX.COL), 1, lnth - 1)
      Call CalculateTotal
      Exit Sub
    End If
  End If
  
  If ALLOW_KEY = False Then
    If ENTER_PRESS = True Then
    Else
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  If ALLOW_KEY = True Then
    If ENTER_PRESS = False Then
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) + Chr(KeyAscii)
      Call CalculateTotal
    End If
  End If
  
  'ENTER PRESS : FORWARD COLUMN
  FWD_COL = False
  If ENTER_PRESS = True Then
    Select Case FLEX.COL
    Case 7
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
    Case 8
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
    Case 9
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL + 1) = Val(FLEX.TextMatrix(FLEX.ROW, FLEX.COL - 1)) - Val(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL + 1) = nstr(FLEX.TextMatrix(FLEX.ROW, FLEX.COL + 1), 7, 3)
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL + 1) = Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL + 1))
       Else
        FWD_COL = False
      End If
    Case 10
        FWD_COL = True
    Case 11
        FWD_COL = False
    Case Else
        FWD_COL = True
    End Select
  End If
  '=======================================================================
    
    If FWD_COL = True Then
      If FLEX.COL = 3 Then
         FLEX.COL = FLEX.COL + 2
      ElseIf FLEX.COL = 10 Then
         If FLEX.Rows - 1 <> FLEX.ROW Then
            FLEX.ROW = FLEX.ROW + 1
            FLEX.COL = FLEX.COL - 3
         End If
      Else
         FLEX.COL = FLEX.COL + 1
      End If
      
      Emptycell = True
    End If
    
  Exit Sub
  
LAST:
  MsgBox "Error In Item Detail"
  FLEX.SetFocus
  Exit Sub
End Sub

Private Sub SavePallet()
On Error GoTo ERRDESC

Dim i As Long, BOXNO As String, PALET As String
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset

Dim ITMCODE As String, ITMQTY As Double, ITMRATE As Double
Dim SRCH As Long: SRCH = 0

For i = 1 To FLEX.Rows - 1
     
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
            "' AND NAME='" & FLEX.TextMatrix(i, 6) & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       MCCD = Trim(RS!CODE & "")
    Else
       MCCD = ""
    End If
    RS.Close
          
    SQL = "UPDATE BOXREGISTER SET MCCD='" & MCCD & "',GRSWGT='" & Val(FLEX.TextMatrix(i, 8)) & _
          "',TRWGT='" & Val(FLEX.TextMatrix(i, 9)) & "',NTWGT='" & Val(FLEX.TextMatrix(i, 10)) & _
          "',COPS='" & Val(FLEX.TextMatrix(i, 7)) & "',LOTNO='" & FLEX.TextMatrix(i, 3) & _
          "',ICOD='" & FindFinishItem(FLEX.TextMatrix(i, 3)) & _
          "',GRAD='" & GetCode("GRDMST", FLEX.TextMatrix(i, 5), "GRAD", "CODE") & "' WHERE COMP = '" & compPth & _
          "' AND UNIT ='" & UNCD & "' AND DVCD ='" & DIVCODE & "' AND VBNO='" & FLEX.TextMatrix(i, 2) & _
          "' AND VTYP='PPF' AND RECSTAT<>'D' AND PKG_STCOD='" & LSPKGCOD & "'"
    
    CN.Execute SQL
    
    If IsAccessOrder(Trim(FLEX.TextMatrix(i, 0))) Then
       ERROROCCUR = True
       Exit Sub
    End If
    
Next

'CODE FOR GENERATE SUMARY AND UPDATE IN ORDMAN
Dim SETRS As ADODB.Recordset
Set SETRS = New ADODB.Recordset

If SETRS.State = 1 Then SETRS.Close
SETRS.Open "SELECT ORDN,ICOD,GRAD,SUM(NTWGT) AS NWGT FROM BOXREGISTER " & _
             "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
             "' AND PKG_STCOD='" & LSPKGCOD & "' AND PLTNO='" & FLEX.TextMatrix(1, 1) & _
             "' AND RECSTAT<>'D' GROUP BY ORDN,ICOD,GRAD", CN, adOpenDynamic, adLockOptimistic
Do While Not SETRS.EOF
    CN.Execute "UPDATE ORDMAN SET DOQTY = DOQTY + " & Val(SETRS!nwgt) & " WHERE COMP='" & compPth & _
               "' AND UNIT='" & UNCD & "' AND ORDN='" & SETRS!ORDN & "' AND ICOD='" & SETRS!ICOD & _
               "' AND TRCD='" & SETRS!grad & "' AND RECSTAT<>'D'"
SETRS.MoveNext
Loop
SETRS.Close

'==============================================================================================

'WORK IN PROCESS
Dim MAINRS As ADODB.Recordset
Set MAINRS = New ADODB.Recordset
Dim L As Long

PALET = FLEX.TextMatrix(1, 1)

If MAINRS.State = 1 Then MAINRS.Close
MAINRS.Open "SELECT DISTINCT DBCD,CHLN FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
            "' AND PKG_STCOD='" & LSPKGCOD & "' AND PLTNO='" & PALET & "'", CN, adOpenDynamic, adLockOptimistic
If Not MAINRS.EOF Then
    CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
               "' AND VTYP='PPF' AND DBCD='" & Trim(MAINRS!dbcd & "") & "' AND VBNO='" & Trim(MAINRS!chln & "") & "'", L
               
    If L = 0 Then
       MsgBox "ILLEGAL OPERATION HAPPEN : CHECKING PURPOSE MESSAGE", vbCritical
       CN.RollbackTrans
       End
    End If
               
    CHALLAN = Trim(MAINRS!chln & "")
End If
MAINRS.Close
    
    'TO FIND TTL QNTY USING GROUP OF LOTNO AND GRAD
    If MAINRS.State = 1 Then MAINRS.Close
    MAINRS.Open "SELECT BOXREGISTER.DBCD,BOXREGISTER.VBDT,BOXREGISTER.LOTNO,BOXREGISTER.GRAD,SUM(GRSWGT) AS GWGT,SUM(TRWGT) AS TWGT,SUM(NTWGT) AS NWGT,SUM(COPS) AS COPS FROM BOXREGISTER " & _
                "INNER JOIN MACMST ON MACMST.COMP = BOXREGISTER.COMP AND MACMST.UNIT = BOXREGISTER.UNIT " & _
                "AND MACMST.DVCD = BOXREGISTER.DVCD AND MACMST.CODE = BOXREGISTER.MCCD AND MACMST.WIPEFFECT='Y' " & _
                "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
                "' AND BOXREGISTER.DVCD='" & DIVCODE & "' AND BOXREGISTER.PKG_STCOD='" & LSPKGCOD & _
                "' AND BOXREGISTER.CHLN='" & CHALLAN & "' GROUP BY BOXREGISTER.DBCD,BOXREGISTER.VBDT,BOXREGISTER.LOTNO,BOXREGISTER.GRAD", CN, adOpenDynamic, adLockOptimistic
    Do While Not MAINRS.EOF
       SRCH = SRCH + 1
       If TEMPRS.State = 1 Then TEMPRS.Close
       TEMPRS.Open "SELECT RICD,PERC,SRCH FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND DVCD='" & DIVCODE & "' AND LTNO='" & MAINRS!LOTNO & "' ORDER BY SRCH", CN, adOpenDynamic, adLockOptimistic
       Do While Not TEMPRS.EOF
          SRCH = SRCH + 1
          ITMCODE = Trim(TEMPRS!RICD & "")
          ITMQTY = Val(nstr((Val(TEMPRS!PERC) * Val(MAINRS!nwgt)) / 100, 10, 3))
          ITMRATE = 0
                    
          CN.Execute "INSERT INTO STORETRAN(COMP,UNIT,DVCD,VTYP,DBCD,SRNO,SRCH,VBNO,CHLN,DATE,CHDT,PCOD,ICOD," & _
                     "PCES,QNTY,GWGT,TWGT,RATE,AMNT,QORP,[USER],[SYSR],OPER,GRAD,LTNO,SUBGRD,COPS,RECSTAT) " & _
                     "VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & "','PPF','" & MAINRS!dbcd & _
                     "','" & LSPKGCOD & "','" & SRCH & "','" & CHALLAN & "','" & CHALLAN & _
                     "','" & Format(MAINRS!VBDT, "YYYY/MM/DD") & "','" & Format(MAINRS!VBDT, "YYYY/MM/DD") & _
                     "','" & MCCD & "','" & ITMCODE & "','0','" & ITMQTY & "','" & Val(MAINRS!GWGT) & _
                     "','" & Val(MAINRS!TWGT) & "','" & ITMRATE & "','" & Val(MAINRS!nwgt) * Val(ITMRATE) & _
                     "','Q','" & cUName & "','N','-','" & MAINRS!grad & "','" & MAINRS!LOTNO & _
                     "','0','" & MAINRS!COPS & "','A')"
                                                                    
        TEMPRS.MoveNext
        Loop
        TEMPRS.Close
    MAINRS.MoveNext
    Loop
    MAINRS.Close
    
    '==============================================================================================
                  
    MsgBox "Pallet No. " & PALET & " Edit Successfully."
    CN.CommitTrans
    
    FLEX.Rows = 1
    FLEX.Rows = 2
    TxtPallet.SetFocus
       
Exit Sub
ERRDESC:
ERROROCCUR = True
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Function FindDivAccessQty() As Double
Dim divrs As ADODB.Recordset
Set divrs = New ADODB.Recordset

If divrs.State = 1 Then divrs.Close
divrs.Open "SELECT DOQTYLIMIT FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & DIVCODE & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not divrs.EOF Then
   FindDivAccessQty = Val(divrs!DOQTYLIMIT)
Else
   FindDivAccessQty = 0
End If

End Function

Private Function IsDataOK() As Boolean
Dim i As Long
IsDataOK = True

For i = 1 To FLEX.Rows - 1
  If Trim(FLEX.TextMatrix(i, 0)) = Empty Then
     IsDataOK = False
     FLEX.ROW = i
     FLEX.COL = 0
     FLEX.SetFocus
     MsgBox "Invalid Data Entered"
     Exit Function
  End If
  
  If Trim(FLEX.TextMatrix(i, 3)) = Empty Then
     IsDataOK = False
     FLEX.ROW = i
     FLEX.COL = 3
     FLEX.SetFocus
     MsgBox "Invalid Data Entered"
     Exit Function
  End If
  
  If Trim(FLEX.TextMatrix(i, 5)) = Empty Then
     IsDataOK = False
     FLEX.ROW = i
     FLEX.COL = 5
     FLEX.SetFocus
     MsgBox "Invalid Data Entered"
     Exit Function
  End If
  
  If Val(FLEX.TextMatrix(i, 7)) <= 0 Then
     IsDataOK = False
     FLEX.ROW = i
     FLEX.COL = 6
     FLEX.SetFocus
     MsgBox "Invalid Data Entered"
     Exit Function
  End If
  
  If Val(FLEX.TextMatrix(i, 10)) <= 0 Then
     IsDataOK = False
     FLEX.ROW = i
     FLEX.COL = 9
     FLEX.SetFocus
     MsgBox "Invalid Data Entered"
     Exit Function
  End If
Next

Dim FILTER As String
FILTER = Empty

For i = 1 To FLEX.Rows - 1
    If Len(Trim(FLEX.TextMatrix(i, 0))) = 10 Then
       If FILTER <> Empty Then FILTER = FILTER & ","
       FILTER = FILTER & "'" & Trim(FLEX.TextMatrix(i, 0)) & "'"
    End If
Next

If RS.State = 1 Then RS.Close
RS.Open "SELECT DISTINCT PCOD FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND ORDN IN (" & FILTER & ") ", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   If RS.RecordCount > 1 Then
      MsgBox "Pallet Must be Contain orders of Single Party "
      IsDataOK = False
      Exit Function
   End If
End If

End Function

Private Sub CalculateTotal()
NETBOXES.Text = 0
    NETCOPS.Text = 0
    NETWGT.Text = 0
    
    Dim i As Double
    i = 1
    For i = 1 To FLEX.Rows - 1
      NETCOPS.Text = Format(Val(NETCOPS.Text) + Val(FLEX.TextMatrix(i, 7)), "######")
      FLEX.TextMatrix(FLEX.ROW, 10) = Val(FLEX.TextMatrix(FLEX.ROW, 8)) - Val(FLEX.TextMatrix(FLEX.ROW, 9))
      NETWGT.Text = Format(Val(NETWGT.Text) + Val(FLEX.TextMatrix(i, 10)), "########.00")
    Next
    NETBOXES.Text = FLEX.Rows - 1
End Sub

Private Sub TXTPALLET_Change()
If Len(Trim(TxtPallet)) = 10 Then   '1
   Call TxtPallet_KeyPress(13)
   Call CalculateTotal
End If
End Sub

Private Sub TXTPALLET_GotFocus()
   TxtPallet.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TxtPallet_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 58
Case 97 To 122
     KeyAscii = KeyAscii - 32
Case 65 To 90
Case 8
Case 13
Case Else
     KeyAscii = 0
     Exit Sub
End Select

If KeyAscii = 32 Or KeyAscii = 95 Then KeyAscii = 0: Exit Sub

If Len(TxtPallet) = 10 And KeyAscii = 13 Then
   If IsPalletDispatchExist(DIVCODE, LSPKGCOD, TxtPallet) Then 'IsPalletDispatchExist
      MsgBox "Pallet No. " & TxtPallet & " has been Dispatched."
      TxtPallet = Empty
      If TxtPallet.Enabled Then TxtPallet.SetFocus
      Exit Sub
   End If
End If

Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset
Dim INDEX As Long

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT *,FINITMMST.NAME AS ITEM,GRDMST.GRAD AS SHADE,MACMST.NAME AS MACHINE FROM BOXREGISTER " & _
            "INNER JOIN FINITMMST ON FINITMMST.COMP=BOXREGISTER.COMP AND FINITMMST.UNIT=BOXREGISTER.UNIT AND " & _
            "FINITMMST.DVCD=BOXREGISTER.DVCD AND FINITMMST.CODE=BOXREGISTER.ICOD " & _
            "INNER JOIN MACMST ON MACMST.COMP=BOXREGISTER.COMP AND MACMST.UNIT=BOXREGISTER.UNIT AND " & _
            "MACMST.DVCD=BOXREGISTER.DVCD AND MACMST.CODE=BOXREGISTER.MCCD " & _
            "INNER JOIN GRDMST ON GRDMST.CODE=BOXREGISTER.GRAD " & _
            "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
            "' AND BOXREGISTER.DVCD='" & DIVCODE & "' AND (VTYP='PPF' OR VTYP='OPN') AND PLTNO='" & TxtPallet & _
            "' AND PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
            
FLEX.Rows = 1
FLEX.Rows = 2
     
If RSDATA.EOF Then
   Exit Sub
End If
     
Do While Not RSDATA.EOF '2
With FLEX
    INDEX = INDEX + 1
    .Rows = FLEX.Rows + 1
    .ROW = FLEX.Rows - 1
    .TextMatrix(INDEX, 0) = Trim(RSDATA!ORDN & "")
    .TextMatrix(INDEX, 1) = Trim(RSDATA!PLTNO & "")
    .TextMatrix(INDEX, 2) = Trim(RSDATA!VBNO & "")
    .TextMatrix(INDEX, 3) = Trim(RSDATA!LOTNO & "")
    .TextMatrix(INDEX, 4) = Trim(RSDATA!Item & "")
    .TextMatrix(INDEX, 5) = Trim(RSDATA!SHADE & "")
    .TextMatrix(INDEX, 6) = Trim(RSDATA!MACHINE & "")
    .TextMatrix(INDEX, 7) = Trim(RSDATA!COPS & "")
    .TextMatrix(INDEX, 8) = Trim(nstr(RSDATA!GRSWGT, 10, 3))
    .TextMatrix(INDEX, 9) = Trim(nstr(RSDATA!TRWGT, 10, 3))
    .TextMatrix(INDEX, 10) = Trim(nstr(RSDATA!NTWGT, 10, 3))
End With
RSDATA.MoveNext
Loop
RSDATA.Close
FLEX.Rows = FLEX.Rows - 1

FLEX.COL = 0
FLEX.ROW = 1
FLEX.SetFocus

End Sub

Private Sub TxtPallet_LostFocus()
  TxtPallet.BackColor = vbWhite
End Sub

Private Sub ReduceOrder()
Dim PALLETNUM As String
PALLETNUM = FLEX.TextMatrix(1, 1)

Dim VALIDRS As ADODB.Recordset
Set VALIDRS = New ADODB.Recordset

'REDUCE DO QTY FROM ORDER WHICH IS RESIDE BEFORE PALLET UPDATION
If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT ORDN,ICOD,GRAD,SUM(NTWGT) AS NWGT FROM BOXREGISTER " & _
             "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
             "' AND PKG_STCOD='" & LSPKGCOD & "' AND PLTNO='" & PALLETNUM & _
             "' AND RECSTAT<>'D' GROUP BY ORDN,ICOD,GRAD", CN, adOpenDynamic, adLockOptimistic
Do While Not VALIDRS.EOF
   SQL = "UPDATE ORDMAN SET DOQTY = DOQTY - " & Val(VALIDRS!nwgt) & " WHERE COMP='" & compPth & _
         "' AND UNIT='" & UNCD & "' AND ORDN='" & VALIDRS!ORDN & "' AND ICOD='" & VALIDRS!ICOD & _
         "' AND TRCD='" & VALIDRS!grad & "' AND RECSTAT<>'D'"
   
   CN.Execute SQL
   VALIDRS.MoveNext
Loop
VALIDRS.Close
'-----------------------------------------------------------------------------------
End Sub

Private Function IsAccessOrder(CHKORDN As String) As Boolean
IsAccessOrder = False

Dim VALIDRS As ADODB.Recordset
Set VALIDRS = New ADODB.Recordset

Dim SUBRS As ADODB.Recordset
Set SUBRS = New ADODB.Recordset

If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT ISNULL(SUM(NTWGT),0) AS NWGT FROM BOXREGISTER WHERE COMP='" & compPth & _
             "' AND UNIT='" & UNCD & "' AND ORDN='" & CHKORDN & "' AND RECSTAT<>'D' AND PLTNO='" & TxtPallet & "'", CN, adOpenDynamic, adLockOptimistic
             
If Not VALIDRS.EOF Then
   SQL = "SELECT ISNULL(SUM(QNTY),0) AS QNTY,ISNULL(SUM(QNTY),0) - ISNULL(SUM(DOQTY),0) - ISNULL(SUM(DISPATCHQTY),0) - ISNULL(SUM(CANCELQTY),0) AS BALQTY FROM ORDMAN WHERE COMP='" & compPth & _
         "' AND UNIT='" & UNCD & "' AND ORDN='" & CHKORDN & "' AND RECSTAT<>'D'"
         
   If SUBRS.State = 1 Then SUBRS.Close
   SUBRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
   If Not SUBRS.EOF Then
      Dim BALQTY As Double, ACCESSQTY As Double
      
      'ACCESSQTY = (FindDivAccessQty * Val(SUBRS!QNTY)) / 100
      'BALQTY = ACCESSQTY + Val(SUBRS!BALQTY)
      
      BALQTY = FindDivAccessQty + Val(SUBRS!BALQTY)
      
      If Val(VALIDRS!nwgt) > BALQTY Then
         IsAccessOrder = True
         MsgBox "Access Packing not Allowed.", vbCritical
         VALIDRS.Close
         SUBRS.Close
         Exit Function
      End If
   End If
End If

End Function


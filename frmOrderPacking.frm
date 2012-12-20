VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmOrderPacking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pallet Packing Against Order"
   ClientHeight    =   6765
   ClientLeft      =   375
   ClientTop       =   1110
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Packing"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6794.673
   ScaleMode       =   0  'User
   ScaleWidth      =   11550.68
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   4
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
         TabIndex        =   5
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
      TabIndex        =   2
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Tag             =   "0"
         Top             =   6360
         Width           =   855
      End
      Begin VB.ComboBox cmbPackingType 
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         ItemData        =   "frmOrderPacking.frx":0000
         Left            =   2040
         List            =   "frmOrderPacking.frx":0002
         TabIndex        =   0
         Tag             =   "0"
         Text            =   "Select Type of Packing"
         Top             =   480
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   285
         Left            =   9840
         TabIndex        =   12
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Format          =   53149697
         CurrentDate     =   39347
      End
      Begin WelchButton.lvButtons_H cmddelitm 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   6270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Delete Row"
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
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   5250
         Left            =   50
         TabIndex        =   13
         Top             =   960
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9260
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColor       =   -2147483628
         BackColorBkg    =   12632256
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
      Begin VB.Label LblItem 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Back Date Packing not Allowed."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   6480
         Width           =   3135
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   6195
         Width           =   1170
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Press {F2} on grid for master help"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   6240
         Width           =   3495
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
         TabIndex        =   14
         Top             =   6195
         Width           =   1050
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Packing :"
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
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Da&te :"
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
         Left            =   9240
         TabIndex        =   10
         Top             =   480
         Width           =   615
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   3
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
Attribute VB_Name = "frmOrderPacking"
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


Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       FLEX.SetFocus
       FLEX.COL = 0
       FLEX.ROW = 1
    End If
End Sub

Private Sub cmddelitm_Click()
  If FLEX.ROW > 1 Then
    FLEX.RemoveItem (FLEX.ROW)
    Call CalculateTotal
    FLEX.Refresh
    FLEX.ROW = FLEX.Rows - 1
    FLEX.COL = 0
    FLEX.SetFocus
  End If
  cmddelitm.Enabled = False
End Sub

Private Sub Flex_Click()
  cmddelitm.Enabled = True
End Sub

Private Sub FLEX_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case FLEX.COL
Case 0
    
    If KeyCode = vbKeyF2 Or (KeyCode = 13 And FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty) Then
        NEW_VISIBLE = False: M_DESC = Empty: Key = Empty
        SQL = "SELECT DISTINCT ORDN,ORDN FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND RECSTAT<>'D' AND OFLG<>'Y' AND FIN_APRV = 'O' AND (QNTY - DOQTY - DISPATCHQTY - CANCELQTY) > 0 "
              
        If FLEX.ROW > 1 Then
           If RS.State = 1 Then RS.Close
            RS.Open "SELECT DISTINCT PCOD FROM ORDMAN WHERE COMP='" & compPth & _
                    "' AND UNIT='" & UNCD & "' AND ORDN ='" & Trim(FLEX.TextMatrix(1, 0)) & "' ", CN, adOpenDynamic, adLockOptimistic
            If Not RS.EOF Then
               SQL = SQL & " AND PCOD='" & RS!PCOD & "' "
            End If
            RS.Close
        End If
        
        Dim SAMEORDN As String
        SAMEORDN = Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))
              
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = SearchList1(SQL, 0, "", "SELECT ORDER FROM LIST")
        FLEX.TextMatrix(FLEX.ROW, 1) = PALETNO
        FLEX.TextMatrix(FLEX.ROW, 2) = GenPackSlipNo(LSPKGCOD)
        If SAMEORDN <> Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
           FLEX.TextMatrix(FLEX.ROW, 3) = ""
           FLEX.TextMatrix(FLEX.ROW, 4) = ""
           FLEX.TextMatrix(FLEX.ROW, 5) = ""
        End If
        
    End If
    
Case 3
    Dim FILTER As String
    Dim FILTER2 As String
    
    If Len(Trim(FLEX.TextMatrix(FLEX.ROW, 0))) = 10 Then
       If RS.State = 1 Then RS.Close
       RS.Open "SELECT * FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
       "' AND RECSTAT<>'D' AND OFLG<>'Y' AND FIN_APRV = 'O' AND ORDN ='" & Trim(FLEX.TextMatrix(FLEX.ROW, 0)) & "' ", CN, adOpenDynamic, adLockOptimistic
       Do While Not RS.EOF
       
          If FindDivAccessQty + (RS!QNTY - RS!DOQTY - RS!DISPATCHQTY - RS!CANCELQTY) > 0 Then
             If FILTER <> Empty Then FILTER = FILTER & ","
             FILTER = FILTER & "'" & Trim(RS!TRCD & "") & "'"
          End If
          
          If FindDivAccessQty + (RS!QNTY - RS!DOQTY - RS!DISPATCHQTY - RS!CANCELQTY) > 0 Then
             If FILTER2 <> Empty Then FILTER2 = FILTER2 & ","
             FILTER2 = FILTER2 & "'" & Trim(RS!ICOD & "") & "'"
          End If
          
       RS.MoveNext
       Loop
       RS.Close
    End If
            
    If KeyCode = vbKeyF2 Or (KeyCode = 13 And FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty) Then
        NEW_VISIBLE = False: M_DESC = Empty: Key = Empty
        SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND DVCD='" & DIVCODE & "' AND ACTIVE = 'Y' AND SHCD IN (" & FILTER & ")  AND FICD IN (" & FILTER2 & ") "
        FLEX.TextMatrix(FLEX.ROW, 3) = SearchList(SQL)
        FLEX.TextMatrix(FLEX.ROW, 4) = FindFinishItem(FLEX.TextMatrix(FLEX.ROW, 3))
        LBLitem.Caption = FindFinItemName(FLEX.TextMatrix(FLEX.ROW, 4))
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
       
TXTVBDT.Value = Now
'TXTVBDT.MinDate = FSDT
Call SetLastDateForPacking
TXTVBDT.MaxDate = FEDT

Call SetPackingType

If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 1

JUMP:
End Sub

Private Sub cmbPackingType_GotFocus()
 If COUNTER > 0 Then cmbPackingType.Locked = True
 SendKeys "{HOME}+{END}"
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
    cmbPackingType.AddItem Trim(PKTYPRS!NAME)
    PKTYPRS.MoveNext
Loop
    PKTYPRS.Close
End Sub

Private Sub SetGlobal()
Dim DBCDRS As ADODB.Recordset
Set DBCDRS = New ADODB.Recordset

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME = '" & cmbPackingType.Text & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   M_DBCD = Trim(DBCDRS!CODE & "")
Else
   M_DBCD = Empty
End If
DBCDRS.Close

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

Private Function FindFinItemName(ICOD As String) As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND CODE ='" & ICOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   FindFinItemName = GRRS!NAME
Else
   FindFinItemName = Empty
End If
GRRS.Close
End Function

Private Sub SetLastDateForPacking()
Dim DTRS As ADODB.Recordset
Set DTRS = New ADODB.Recordset

If DTRS.State = 1 Then DTRS.Close
DTRS.Open "SELECT IsNull(LSTPCKDT,'" & FSDT & "') AS LSTPCKDT FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & LSPKGCOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DTRS.EOF Then
   TXTVBDT.MinDate = Format(DTRS!LSTPCKDT, "DD/MM/YYYY")
End If
DTRS.Close
End Sub

Private Function IsBoxExistInUnit(BOXNUM As String) As Boolean
IsBoxExistInUnit = False

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND VBNO='" & BOXNUM & "'", CN, adOpenDynamic, adLockOptimistic
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
    FLEX.TextMatrix(0, 6) = "Machine"
    FLEX.TextMatrix(0, 7) = "Cops"
    FLEX.TextMatrix(0, 8) = "Gross Wgt"
    FLEX.TextMatrix(0, 9) = "Tare Wgt"
    FLEX.TextMatrix(0, 10) = "Net Wgt"
    
    FLEX.ColWidth(0) = 1250
    FLEX.ColWidth(1) = 1250
    FLEX.ColWidth(2) = 1250
    FLEX.ColWidth(3) = 1300
    FLEX.ColWidth(4) = 0
    FLEX.ColWidth(5) = 1800
    FLEX.ColWidth(6) = 1800
    FLEX.ColWidth(7) = 600
    FLEX.ColWidth(8) = 1100
    FLEX.ColWidth(9) = 1100
    FLEX.ColWidth(10) = 1400
    
    FLEX.ColAlignment(8) = 1
    FLEX.ColAlignment(9) = 1
    FLEX.ColAlignment(10) = 1
    
End Sub

Private Sub FLEX_EnterCell()
  FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
  Emptycell = True
  FLEX.TextMatrix(FLEX.ROW, 4) = FindFinishItem(FLEX.TextMatrix(FLEX.ROW, 3))
  LBLitem.Caption = FindFinItemName(FLEX.TextMatrix(FLEX.ROW, 4))
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
      
      If FLEX.COL = 10 Then
        If FLEX.Rows - 1 = FLEX.ROW Then
            
            If Trim(FLEX.TextMatrix(FLEX.ROW, 0)) = Empty Then
               MsgBox "Order is must"
               FLEX.SetFocus
               FLEX.COL = 0
               Exit Sub
            End If
            
            Dim AYS
            AYS = MsgBox("Is Pallet Complete ?", vbYesNo + vbDefaultButton2)
                    
            If AYS = vbYes Then
              
              If Not IsDataOK Then
                 Exit Sub
              End If
              
              CN.BeginTrans
              'FOR PALLET NO.
              PALETNO = GenPackSlipNo(LSPKGCOD, "LPNO")
              ERROROCCUR = False
              
              If IsPalletExistInUnit(PALETNO) Then
                 MsgBox "PalletNo. " & PALETNO & " Already Exist."
                 CN.RollbackTrans
                 Exit Sub
              End If
              
              Call SavePallet(PALETNO) ' PALLET NO.
              
              If ERROROCCUR Then
                 CN.RollbackTrans
                 Exit Sub
              End If
            Else
              FLEX.LeftCol = 0
            End If
              
            If AYS = vbYes Then
               Dim ORDN As String
               ORDN = FLEX.TextMatrix(FLEX.ROW, 0)
               FLEX.Rows = 1
               FLEX.Rows = 2
               FLEX.ROW = FLEX.Rows - 1
               FLEX.TextMatrix(FLEX.ROW, 0) = ORDN
               FLEX.TextMatrix(FLEX.ROW, 2) = GenPackSlipNo(LSPKGCOD)
            Else
              FLEX.Rows = FLEX.Rows + 1
              FLEX.ROW = FLEX.Rows - 1
              FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.TextMatrix(FLEX.ROW - 1, 0)
              FLEX.TextMatrix(FLEX.ROW, 3) = FLEX.TextMatrix(FLEX.ROW - 1, 3)
              FLEX.TextMatrix(FLEX.ROW, 5) = FLEX.TextMatrix(FLEX.ROW - 1, 5)
              FLEX.TextMatrix(FLEX.ROW, 7) = FLEX.TextMatrix(FLEX.ROW - 1, 7)
            End If
            
              FLEX.TextMatrix(FLEX.ROW, 1) = PALETNO
        End If
        
        FLEX.ROW = FLEX.Rows - 1
        FLEX.COL = 3
        FLEX.SetFocus
        
      ElseIf FLEX.COL = 3 Then
       FLEX.COL = FLEX.COL + 2
      ElseIf FLEX.COL = 9 Then
          FLEX.COL = FLEX.COL + 1
          FLEX.LeftCol = FLEX.LeftCol + 1
      Else
        FLEX.COL = FLEX.COL + 1
        FLEX.LeftCol = 0
      End If
      Emptycell = True
    End If
  Exit Sub
LAST:
  MsgBox "Error In Item Detail"
  FLEX.SetFocus
  Exit Sub
End Sub

Private Sub SavePallet(PALLETNO As String)
On Error GoTo ERRDESC
Dim i As Long, BOXNO As String
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset

Call SetGlobal

Dim ITMCODE As String, ITMQTY As Double, ITMRATE As Double
Dim SRCH As Long: SRCH = 0
Dim TTL_QTY As Double
TTL_QTY = 0

For i = 1 To FLEX.Rows - 1
  If FLEX.TextMatrix(i, 1) = PALLETNO Then
    BOXNO = GenPackSlipNo(LSPKGCOD)
           
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
            "' AND NAME='" & FLEX.TextMatrix(i, 6) & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       MCCD = Trim(RS!CODE & "")
    Else
       MCCD = ""
    End If
    RS.Close
    
    If IsBoxExistInUnit(Trim(BOXNO)) Then
       MsgBox "BoxNo. " & BOXNO & " Already Exist. Either Deleted Entry or Exist in Other Packing Station. Update No. in Station", vbCritical
       ERROROCCUR = True
       Exit Sub
    End If
               
    SQL = "INSERT INTO BOXREGISTER(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,PLTNO,VBDT,CHLN,PKG_STCOD,PKGNG_COD,"
    SQL = SQL & "LOCCOD,PCOD,ISRETURNABLE,LOTNO,ICOD,GRAD,SUBGRD,MCCD,COPS,BOXWGT,COPSWGT,GRSWGT,TRWGT,"
    SQL = SQL & "NTWGT,PACKER,RMRK,RECSTAT,ORDN)VALUES('" & compPth & _
    "','" & UNCD & "','" & DIVCODE & "','" & M_DBCD & "','PPF','" & BOXNO & "','" & PALLETNO & _
    "','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & CHALLAN & _
    "','" & LSPKGCOD & "','000001','000001','000001','N','" & FLEX.TextMatrix(i, 3) & _
    "','" & FindFinishItem(FLEX.TextMatrix(i, 3)) & "','" & GetCode("GRDMST", FLEX.TextMatrix(i, 5), "GRAD", "CODE") & _
    "','0','" & MCCD & "','" & Val(FLEX.TextMatrix(i, 7)) & _
    "','0','0','" & Val(FLEX.TextMatrix(i, 8)) & "','" & Val(FLEX.TextMatrix(i, 9)) & _
    "','" & Val(FLEX.TextMatrix(i, 10)) & "','" & cUName & "','','A','" & Trim(FLEX.TextMatrix(i, 0)) & "')"
       
    CN.Execute SQL
    FLEX.TextMatrix(i, 2) = BOXNO
    
    CN.Execute "UPDATE PCKMST SET [LBNO]='" & BOXNO & "',LSTPCKDT = '" & Format(TXTVBDT, "MM/DD/YYYY") & _
           "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & LSPKGCOD & "'"
           
    If IsAccessOrder(Trim(FLEX.TextMatrix(i, 0))) Then
       ERROROCCUR = True
       Exit Sub
    End If
    
  End If
Next

TXTVBDT.MinDate = Format(TXTVBDT, "DD/MM/YYYY")

'CODE FOR GENERATE SUMARY AND UPDATE IN ORDMAN
Dim SETRS As ADODB.Recordset
Set SETRS = New ADODB.Recordset

If SETRS.State = 1 Then SETRS.Close
SETRS.Open "SELECT ORDN,ICOD,GRAD,SUM(NTWGT) AS NWGT FROM BOXREGISTER " & _
             "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
             "' AND PKG_STCOD='" & LSPKGCOD & "' AND CHLN='" & CHALLAN & _
             "' AND RECSTAT<>'D' GROUP BY ORDN,ICOD,GRAD", CN, adOpenDynamic, adLockOptimistic
Do While Not SETRS.EOF
    CN.Execute "UPDATE ORDMAN SET DOQTY = DOQTY + " & Val(SETRS!nwgt) & " WHERE COMP='" & compPth & _
               "' AND UNIT='" & UNCD & "' AND ORDN='" & SETRS!ORDN & "' AND ICOD='" & SETRS!ICOD & _
               "' AND TRCD='" & SETRS!grad & "' AND RECSTAT<>'D' AND OFLG<>'Y' AND FIN_APRV = 'O'"
SETRS.MoveNext
Loop
SETRS.Close

'==============================================================================================
    Dim MAINRS As ADODB.Recordset
    Set MAINRS = New ADODB.Recordset
    'TO FIND TTL QNTY USING GROUP OF LOTNO AND GRAD
    CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
    If MAINRS.State = 1 Then MAINRS.Close
    MAINRS.Open "SELECT LOTNO,GRAD,SUM(GRSWGT) AS GWGT,SUM(TRWGT) AS TWGT,SUM(NTWGT) AS NWGT,SUM(COPS) AS COPS FROM BOXREGISTER " & _
                "INNER JOIN MACMST ON MACMST.COMP = BOXREGISTER.COMP AND MACMST.UNIT = BOXREGISTER.UNIT " & _
                "AND MACMST.DVCD = BOXREGISTER.DVCD AND MACMST.CODE = BOXREGISTER.MCCD AND MACMST.WIPEFFECT='Y' " & _
                "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
                "' AND BOXREGISTER.DVCD='" & DIVCODE & "' AND BOXREGISTER.PKG_STCOD='" & LSPKGCOD & _
                "' AND BOXREGISTER.CHLN='" & CHALLAN & "' GROUP BY LOTNO,GRAD", CN, adOpenDynamic, adLockOptimistic
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
                     "VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & "','PPF','" & M_DBCD & _
                     "','" & LSPKGCOD & "','" & SRCH & "','" & CHALLAN & "','" & CHALLAN & _
                     "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
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
    
    CN.Execute "UPDATE PCKMST SET [LCNO]='" & CHALLAN & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND CODE='" & LSPKGCOD & "'"
    
    CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
    '==============================================================================================
    CN.Execute "UPDATE PCKMST SET [LPNO]='" & PALETNO & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                         "' AND CODE='" & LSPKGCOD & "'"
              
    MsgBox "Assigned Pallet No. " & PALETNO
    PALETNO = GenPackSlipNo(LSPKGCOD, "LPNO")
    CN.CommitTrans
       
Exit Sub
ERRDESC:
ERROROCCUR = True
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Function IsAccessOrder(CHKORDN As String) As Boolean
IsAccessOrder = False

Dim VALIDRS As ADODB.Recordset
Set VALIDRS = New ADODB.Recordset

Dim SUBRS As ADODB.Recordset
Set SUBRS = New ADODB.Recordset

'INCLUDE PRODUCTION AND DISPATCH
If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT ISNULL(SUM(NTWGT),0) AS NWGT FROM BOXREGISTER WHERE COMP='" & compPth & _
             "' AND UNIT='" & UNCD & "' AND ORDN='" & CHKORDN & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
             
If Not VALIDRS.EOF Then
            
    SQL = "SELECT ISNULL(SUM(QNTY),0) - ISNULL(SUM(CANCELQTY),0) AS BALQTY FROM ORDMAN WHERE COMP='" & compPth & _
          "' AND UNIT='" & UNCD & "' AND ORDN='" & CHKORDN & "' AND RECSTAT<>'D'"
         
   If SUBRS.State = 1 Then SUBRS.Close
   SUBRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
   If Not SUBRS.EOF Then
      Dim BALQTY As Double, ACCESSQTY As Double
      
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

For i = 1 To FLEX.Rows - 1
    If Len(Trim(FLEX.TextMatrix(i, 0))) = 10 Then
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                "' AND ORDN IN (" & FILTER & ") AND ICOD='" & Trim(FLEX.TextMatrix(i, 4)) & _
                "' AND TRCD = '" & GetCode("GRDMST", Trim(FLEX.TextMatrix(i, 5)), "GRAD", "CODE") & "'", CN, adOpenDynamic, adLockOptimistic
                
        If RS.EOF Then
           MsgBox "Order Not Match Item and Shade Details", vbCritical
           IsDataOK = False
           Exit Function
        End If
    End If
Next

'for machine master
For i = 1 To FLEX.Rows - 1
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DVCD='" & DIVCODE & _
            "' AND NAME = '" & Trim(FLEX.TextMatrix(i, 6)) & "'", CN, adOpenDynamic, adLockOptimistic
                
    If RS.EOF Then
       MsgBox "Machine Not Match with Details", vbCritical
       IsDataOK = False
       Exit Function
    End If
Next

End Function

Private Sub CalculateTotal()
NETBOXES.Text = 0
    NETCOPS.Text = 0
    NETWGT.Text = 0
    
    Dim i As Double
    i = 1
    For i = 1 To FLEX.Rows - 1
      NETCOPS.Text = Format(Val(NETCOPS.Text) + Val(FLEX.TextMatrix(i, 7)), "######")
      NETWGT.Text = Format(Val(NETWGT.Text) + Val(FLEX.TextMatrix(i, 10)), "########.00")
      FLEX.TextMatrix(FLEX.ROW, 10) = Val(FLEX.TextMatrix(FLEX.ROW, 8)) - Val(FLEX.TextMatrix(FLEX.ROW, 9))
    Next
    NETBOXES.Text = FLEX.Rows - 1
End Sub

Private Function IsPalletExistInUnit(PLTNUM As String) As Boolean
IsPalletExistInUnit = False

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND PLTNO='" & PLTNUM & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
   IsPalletExistInUnit = True
End If
CHKRS.Close
End Function



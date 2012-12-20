VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_OrderVsProduction 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   Pending Order V/s Production"
   ClientHeight    =   4335
   ClientLeft      =   1680
   ClientTop       =   1170
   ClientWidth     =   6480
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      Height          =   675
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   6330
      Begin VB.TextBox TXTPACKINGTYPE 
         Height          =   330
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   165
         Width           =   4860
      End
      Begin VB.Label Label1 
         Caption         =   "Packing &Type "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   120
      TabIndex        =   19
      Top             =   645
      Width           =   6330
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   165
         Width           =   4860
      End
      Begin VB.Label Label14 
         Caption         =   "&Division "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      Height          =   795
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   6330
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Text            =   "100"
         Top             =   300
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2640
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Preview"
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
         Image           =   "frmRPT_OrderVsProduction.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "frmRPT_OrderVsProduction.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label13 
         Caption         =   "Report &Zoom %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   6330
      Begin VB.TextBox TXTPKGSTATION 
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   4860
      End
      Begin VB.TextBox txtLTNo 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtDENIER 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   4920
      End
      Begin VB.Label Label2 
         Caption         =   "Pkg &Station"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label9 
         Caption         =   "&Lot No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label8 
         Caption         =   "Finish &Item "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.Frame framDIVISION 
      Height          =   630
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6330
      Begin VB.TextBox txtUNIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   195
         Width           =   4860
      End
      Begin VB.Label lblUnit 
         Caption         =   "&Unit Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmRPT_OrderVsProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_DBCD As String
Dim PERIOD As String
Dim M_MCNO As String
Dim PACK As String

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
Dim FLAG As Boolean
Dim PKG As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim M_DVCD As String
    CRPT.Reset
    crptConnect CRPT
    
    If Not RIGHTDATA Then Exit Sub
    
    M_MCNO = Empty
    rptsql = Empty
    
    
       M_DBCD = GetVtypCode(txtUNIT.Tag, "PPF", TXTPACKINGTYPE)
    
       rptsql = "{BOXREGISTER.COMP}='" & compPth & "' AND {BOXREGISTER.UNIT}='" & txtUNIT.Tag & _
                "' AND {BOXREGISTER.DVCD}='" & txtDVCD.Tag & "' AND {BOXREGISTER.RECSTAT}<>'D' " & _
                "AND ({BOXREGISTER.VTYP}='PPF' OR {BOXREGISTER.VTYP}='OPN')"
        
'FILTERATION CRITERIA

    If Not TXTPACKINGTYPE = Empty Then rptsql = rptsql & " AND {BOXREGISTER.DBCD}='" & TXTPACKINGTYPE.Tag & "'"
    If Not txtDENIER = Empty Then rptsql = rptsql & " AND {BOXREGISTER.ICOD}='" & txtDENIER.Tag & "'"
    If Not txtLTNo = Empty Then rptsql = rptsql & " AND {BOXREGISTER.LOTNO}= '" & txtLTNo & "'"
    If Not TXTPKGSTATION = Empty Then rptsql = rptsql & " AND {BOXREGISTER.PKG_STCOD} = '" & TXTPKGSTATION.Tag & "'"
     
    ReportName = App.PATH & "\Reports\frmrpt_PendingOrderVsDispatch.rpt"
    RPTN = "PENDING ORDER V/s PRODUCTION "
    
    If Not TXTPACKINGTYPE = Empty Then RPTN = RPTN & "  For ," & TXTPACKINGTYPE
        
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula rptsql
    
        With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        .Formulas(3) = "DIVISION='" & txtDVCD & "'"
        .Formulas(4) = "REPORTHEAD='" & RPTN & "'"
        
         RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        
        
         CRPT.WindowShowPrintBtn = True
         CRPT.WindowShowPrintSetupBtn = True
               
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
         txtUNIT.SetFocus
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    Exit Sub
errPreview:
 MsgBox ERR.Description
 End Sub


Private Sub Form_Activate()
     Call ColorComponent(Me)
    If Not txtUNIT = Empty Then Exit Sub
    
    txtUnit_KeyDown vbKeyReturn, 0
    
    If txtUNIT = Empty Then
        cmdpreview.Enabled = False
    End If
    
    If txtUNIT = Empty Then Unload Me: Exit Sub
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtDVCD" And txtDVCD = Empty Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  Me.Left = Me.Left - 900
  
  txtUNIT = UntNm
  txtUNIT.Tag = UNCD
End Sub


Private Sub txtDenier_GotFocus()
   txtDENIER.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDenier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        txtDENIER = Empty
        txtDENIER.Tag = Empty
        M_DESC = Empty
        NEW_VISIBLE = False
        txtDENIER = SearchList1("Select TOP 20 Code,Name  From FINITMMST WHERE COMP='" & compPth & _
        "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Denier !!")
        txtDENIER.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDENIER = Empty
        txtDENIER.Tag = Empty
    End If
    
End Sub

Private Sub txtDenier_LostFocus()
 txtDENIER.BackColor = vbWhite
End Sub

Private Sub txtDVCD_Change()


TXTPKGSTATION.Text = Empty
txtDENIER.Text = Empty
txtLTNo.Text = Empty



End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "'  AND CODE<>'000001' AND RECSTAT='A'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    End If
End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtltno_GotFocus()
txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   txtLTNo = SearchList1("SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
   "' AND DVCD = '" & txtDVCD.Tag & "'", 0, txtLTNo, "SELECT LOTNO FROM LIST")
   txtLTNo.Tag = Key
ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
   txtLTNo = Empty
   txtLTNo.Tag = Empty
End If
End Sub

Private Sub txtltno_LostFocus()
 txtLTNo.BackColor = vbWhite
End Sub

Private Sub TXTPACKINGTYPE_GotFocus()
  TXTPACKINGTYPE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPACKINGTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTPACKINGTYPE = Empty) Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   TXTPACKINGTYPE = SearchList1("SELECT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & _
   "' AND UNIT='" & txtUNIT.Tag & "' AND VTYP='PPF' AND FYCD='" & FYCD & "'", 0, TXTPACKINGTYPE, "SELECT PACKING TYPE FROM LIST")
   TXTPACKINGTYPE.Tag = Key
End If
End Sub

Private Sub TXTPACKINGTYPE_LostFocus()
TXTPACKINGTYPE.BackColor = vbWhite
End Sub

Private Sub TXTPKGSTATION_GotFocus()
 TXTPKGSTATION.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPKGSTATION_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
        TXTPKGSTATION = SearchList1("SELECT CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "' AND RECSTAT='A'", 0, TXTPKGSTATION, "SELECT PACKING STATION FROM LIST")
        TXTPKGSTATION.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTPKGSTATION = Empty
        TXTPKGSTATION.Tag = Empty
    End If
End Sub

Private Sub TXTPKGSTATION_LostFocus()
 TXTPKGSTATION.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
 txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    End If
    
End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub

Private Sub TXTZOOM_GotFocus()
txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTZOOM_LostFocus()
txtZoom.BackColor = vbWhite
End Sub

Private Function RIGHTDATA() As Boolean
 RIGHTDATA = True
 
 If txtUNIT = Empty Then
    RIGHTDATA = False
    MsgBox "Please Select Unit !!", vbInformation, "Unit Is Key Field Missing"
    txtUNIT.SetFocus
    Exit Function
 End If
    
 If txtDVCD = Empty Then
    RIGHTDATA = False
    MsgBox "Please Select Division !!", vbInformation, "Division Missing !!"
    txtDVCD.SetFocus
    txtDVCD_KeyDown vbKeyReturn, 0
    Exit Function
 End If
   
End Function


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form FRM_WGHTRANG 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grade Wise Weight Range"
   ClientHeight    =   7305
   ClientLeft      =   4590
   ClientTop       =   1935
   ClientWidth     =   10425
   Icon            =   "FRM_WGHTRANG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10425
   Begin VB.TextBox TXTENDWGT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5400
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox TXTSTARTWGT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox M_GRAD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Frame Frame4 
      Height          =   4815
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   9975
      Begin VB.TextBox SUBPKGNG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox SEQNO 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox M_RAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6600
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox M_GRD1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin WelchButton.lvButtons_H CMDOK 
         Height          =   375
         Left            =   8640
         TabIndex        =   9
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
         Image           =   "FRM_WGHTRANG.frx":0442
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   8640
         TabIndex        =   18
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         Image           =   "FRM_WGHTRANG.frx":07DC
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   3255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   7
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
      Begin VB.Label LBLSUB 
         Caption         =   "Sub Packaging Type"
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
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line Line3 
         X1              =   7560
         X2              =   7560
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   1680
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Seq.No."
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
         Height          =   255
         Left            =   7560
         TabIndex        =   19
         Tag             =   "S"
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   3480
         X2              =   3480
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   8520
         X2              =   8520
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         Height          =   1095
         Left            =   120
         Top             =   240
         Width           =   9735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Diff."
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
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Tag             =   "S"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Starting Weight"
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
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label LBLSHADE 
         Alignment       =   2  'Center
         Caption         =   "Sub Grade"
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
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Tag             =   "S"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   6480
         X2              =   6480
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Label Label8 
         Caption         =   "Ending Weight"
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
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   5040
         X2              =   5040
         Y1              =   240
         Y2              =   1320
      End
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   6480
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
      Image           =   "FRM_WGHTRANG.frx":0C2E
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
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
      Image           =   "FRM_WGHTRANG.frx":19B8
      cBack           =   -2147483633
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "SUB - GRADE MASTER "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   240
      Width           =   2895
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   240
      X2              =   10440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label M_DVNM 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   720
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   7095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label Label3 
      Caption         =   "DIVISION  :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label LBLGRADE 
      Caption         =   "GRADE      :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "FRM_WGHTRANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************
'Module Name : FRM_WGHTRANG
'
'Develope By :
'
'Develope Date : 29 April 2010
'
'Change Date :
'
'Change By :
'
'Remark :
'*******************************************

Option Explicit
Dim SWITCH As Boolean
Dim M_DVCD As String
Dim M_BXMC As String
Dim M_SUBPKG As String
Dim SUBPKGCODE As String
Dim SUBRS As New ADODB.Recordset
Dim ROWNO As Long
Dim INFORS As New ADODB.Recordset
Dim IsShadeReq As Boolean
Dim SAVEFLAG As Boolean

Private Sub cmdCancel_Click()

If Not SAVEFLAG Then
   MsgBox "Not Allow such operation", vbCritical
   Exit Sub
End If

Dim CURSOR As Long
Dim J As Long

For J = ROWNO To FLEX.Rows - 2
 FLEX.TextMatrix(J, 0) = FLEX.TextMatrix(J + 1, 0)
 FLEX.TextMatrix(J, 1) = FLEX.TextMatrix(J + 1, 1)
 FLEX.TextMatrix(J, 2) = FLEX.TextMatrix(J + 1, 2)
 FLEX.TextMatrix(J, 3) = FLEX.TextMatrix(J + 1, 3)
 FLEX.TextMatrix(J, 4) = FLEX.TextMatrix(J + 1, 4)
 FLEX.TextMatrix(J, 6) = FLEX.TextMatrix(J + 1, 6)
Next J

FLEX.Rows = FLEX.Rows - 1
Call CLEARDATA
SWITCH = False
M_GRD1.SetFocus
cmdOk.Caption = "&Add"
cmdCancel.Enabled = False

End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub Flex_Click()
If FLEX.Rows > 1 And FLEX.TextMatrix(FLEX.ROW, 0) <> Empty Then
    cmdOk.Caption = "Upd&ate"
    cmdCancel.Enabled = True
    ROWNO = FLEX.ROW
    M_GRD1 = Trim(FLEX.TextMatrix(ROWNO, 0))
    TXTSTARTWGT = FLEX.TextMatrix(ROWNO, 1)
    TXTENDWGT = FLEX.TextMatrix(ROWNO, 2)
    M_RAT = FLEX.TextMatrix(ROWNO, 3)
    SEQNO = FLEX.TextMatrix(ROWNO, 4)
    SUBPKGNG.Text = Trim(FLEX.TextMatrix(ROWNO, 6))
    SWITCH = True
End If
End Sub

Private Sub Form_Activate()
If DIVNAM = Empty Or DIVCOD = Empty Then
   MsgBox "SELECT DIVISION"
   Unload Me
End If

  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  LBLHEAD.BackColor = &H80&
  LBLHEAD.ForeColor = &HFFFFFF
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If UCase(ActiveControl.NAME) = "M_GRD1" And M_GRD1 = Empty Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  SAVEFLAG = True
  DIVCOD = Empty: DIVNAM = Empty: Key = Empty
  DIVNAM = SearchList1("SELECT TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE <> '000001' AND RECSTAT='A' ", 0, "", "SELECT DIVISION FOR SUBGRADE")
  M_DVNM.Caption = DIVNAM
  DIVCOD = Key:  Me.Tag = DIVCOD
  
  Call SETFLEX
  IsShadeReq = False
  If SetIsShadeReq(DIVCOD) = "Y" Then
     IsShadeReq = True
     Call SetShadeInfo
  End If
    
  
    
  Me.KeyPreview = True
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdSave_Click()

   On Error GoTo LAST
   
   If IsShadeReq Then
      Call SaveShadeInfo
      Exit Sub
   End If
   
    Dim INDEX As Long
    If Trim(M_GRAD) = Empty Then
        MsgBox "Please Enter Grade !!", vbInformation
        M_GRAD.SetFocus
        Exit Sub
    Else
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM GRDMST WHERE GRAD='" & M_GRAD & "'", CN, adOpenKeyset, adLockPessimistic
        If RS.EOF Then
            MsgBox "Grade Does Not Exist", vbCritical
            M_GRAD.SetFocus
            Exit Sub
        End If
        RS.Close
    End If
    
    
    If FLEX.Rows < 3 Then
       Exit Sub
    End If
    
   
        With FLEX
            For INDEX = 1 To FLEX.Rows - 2
                If RS.State = 1 Then RS.Close
                RS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                        "' AND  DVCD='" & DIVCOD & "' AND NAME='" & .TextMatrix(INDEX, 0) & _
                        "' AND GRAD <>'" & M_GRAD.Tag & "'", CN, adOpenKeyset, adLockPessimistic
                If Not RS.EOF Then
                    MsgBox "Sub Grade " & Trim(.TextMatrix(INDEX, 0)) & " Already Exist.", vbCritical
                    M_GRD1.SetFocus
                    Exit Sub
                End If
                RS.Close
            Next INDEX
        End With
         
        CN.Execute "DELETE FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND  DVCD='" & DIVCOD & "' AND GRAD='" & M_GRAD.Tag & "'"
                   
     
With FLEX
For INDEX = 1 To FLEX.Rows - 2
'-----------------------------------------------------------------------------------
' FINDING SUBPACKAGING CODE
  SUBPKGCODE = GetCode("SUBPKGNGMST", Trim(FLEX.TextMatrix(INDEX, 6)), "NAME", "CODE")
'-----------------------------------------------------------------------------------
CN.Execute "INSERT INTO SUBGRDMST(COMP,UNIT,DVCD,GRAD,SUBGRD,NAME,SWGT,EWGT,RDIFF,SEQNO,STATUS,RECSTAT,SUBPKGCODE) VALUES('" & compPth & _
"','" & UNCD & "','" & DIVCOD & "','" & Val(M_GRAD.Tag) & "','" & INDEX & "','" & .TextMatrix(INDEX, 0) & _
"','" & .TextMatrix(INDEX, 1) & "','" & .TextMatrix(INDEX, 2) & "','" & .TextMatrix(INDEX, 3) & _
"','" & .TextMatrix(INDEX, 4) & "','A','A','" & Trim(SUBPKGCODE) & "')"
Next INDEX
End With

Call RESETALL
Exit Sub
LAST:
  
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub

Private Sub RESETALL()
 M_GRAD = Empty
 Call CLEARDATA
 SAVEFLAG = True
 FLEX.Rows = 1
 FLEX.Rows = 2
 If M_GRD1.Enabled Then M_GRD1.SetFocus
End Sub

Private Sub M_GRAD_GotFocus()
  M_GRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
  M_GRAD.SelStart = 0
  M_GRAD.SelLength = Len(M_GRAD)
  Msg "Select Grade From Master(Press ->> F2) for Help"
End Sub

Private Sub M_GRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If Trim(M_GRAD.Text) = Empty Or KeyCode = vbKeyF2 Then
    M_GRAD.Text = SearchList1("select TOP 20 grad AS CODE,grad from grdmst", 0, M_GRAD.Text, "SELECT GRAD FROM MASTER")
  End If
End Sub

Private Sub M_GRAD_LostFocus()
  M_GRAD.BackColor = vbWhite
End Sub

Private Sub M_GRAD_Validate(Cancel As Boolean)
Dim INDEX As Double
On Error GoTo LAST
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM GRDMST WHERE GRAD='" & M_GRAD & "'", CN, adOpenKeyset, adLockPessimistic
  If RS.EOF Then
      MsgBox "Grade Does Not Exist", vbCritical
      M_GRAD.SetFocus
      Cancel = True
      Exit Sub
  Else
      M_GRAD.Tag = RS!CODE
  End If
    
  FLEX.Rows = 1
  FLEX.Rows = 2
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND  DVCD='" & DIVCOD & "' AND GRAD='" & M_GRAD.Tag & "' AND RECSTAT='A'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
     SAVEFLAG = False
  Else
     SAVEFLAG = True
  End If
  
  Do While Not RS.EOF
  
    ROWNO = FLEX.Rows - 1
    FLEX.TextMatrix(ROWNO, 0) = RS!NAME
    FLEX.TextMatrix(ROWNO, 1) = Format(STR(RS!SWGT), "000.000")
    FLEX.TextMatrix(ROWNO, 2) = Format(STR(RS!EWGT), "000.000")
    FLEX.TextMatrix(ROWNO, 3) = RS!RDIFF
    FLEX.TextMatrix(ROWNO, 4) = RS!SEQNO
    FLEX.TextMatrix(ROWNO, 5) = M_GRAD
    FLEX.TextMatrix(ROWNO, 6) = GetCode("SUBPKGNGMST", RS!SUBPKGCODE & "", "CODE", "NAME")
    FLEX.Rows = FLEX.Rows + 1
    RS.MoveNext
  Loop

  RS.Close
    
  Exit Sub
LAST:
  MsgBox ERR.Description
End Sub

Private Sub M_GRD1_GotFocus()
    M_GRD1.BackColor = RGB(BRED, BGREEN, BBLUE)
    M_GRD1.SelStart = 0
    M_GRD1.SelLength = Len(M_GRD1)
    Msg "Enter SubGrade"
End Sub

Private Sub M_GRD1_LostFocus()
M_GRD1.BackColor = vbWhite
End Sub

Private Sub M_RAT_GotFocus()
M_RAT.BackColor = RGB(BRED, BGREEN, BBLUE)
M_RAT.SelStart = 0
M_RAT.SelLength = Len(M_RAT)
Msg "Enter Rate Difference"
End Sub

Private Sub M_RAT_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub M_RAT_LostFocus()
M_RAT.BackColor = vbWhite
End Sub

Private Sub SEQNO_GotFocus()
    SEQNO.BackColor = RGB(BRED, BGREEN, BBLUE)
    SEQNO.SelStart = 0
    SEQNO.SelLength = Len(SEQNO)
    Msg "Enter Rate Difference"
End Sub

Private Sub SEQNO_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case 48 To 57, 8
            'Valid Keys
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub SETFLEX()
  FLEX.Clear
  FLEX.ColWidth(0) = 1800
  FLEX.ColWidth(1) = 1250
  FLEX.ColWidth(2) = 1250
  FLEX.ColWidth(3) = 1250
  FLEX.ColWidth(4) = 1050
  FLEX.ColWidth(5) = 0
  FLEX.ColWidth(6) = 1500
  
  FLEX.Clear
  FLEX.TextMatrix(0, 0) = "SubGrade"
  FLEX.TextMatrix(0, 1) = "Start Wgt"
  FLEX.TextMatrix(0, 2) = "End Wgt"
  FLEX.TextMatrix(0, 3) = "Rate Diff."
  FLEX.TextMatrix(0, 4) = "Seq No."
  FLEX.TextMatrix(0, 5) = "Grade"
  FLEX.TextMatrix(0, 6) = "SubPackage"
      
  FLEX.ColAlignment(0) = vbLeftJustify
  FLEX.ColAlignment(1) = vbRightJustify
  FLEX.ColAlignment(2) = vbRightJustify
  FLEX.ColAlignment(3) = vbRightJustify
  FLEX.ColAlignment(4) = vbRightJustify
 'FLEX.ColAlignment(6) = vbRightJustify
End Sub


Private Sub CMDOK_Click()
 Dim INDEX As Long
 
 If Not SWITCH Then
      ROWNO = FLEX.Rows - 1
 End If
 
 If CheckData(ROWNO) Then Exit Sub
 
    FLEX.TextMatrix(ROWNO, 0) = Trim(M_GRD1)
    FLEX.TextMatrix(ROWNO, 1) = Val(TXTSTARTWGT)
    FLEX.TextMatrix(ROWNO, 2) = Val(TXTENDWGT)
    FLEX.TextMatrix(ROWNO, 3) = Val(M_RAT)
    FLEX.TextMatrix(ROWNO, 4) = SEQNO
    FLEX.TextMatrix(ROWNO, 5) = M_GRAD.Tag
    FLEX.TextMatrix(ROWNO, 6) = Trim(SUBPKGNG.Text)
    
    If Not SWITCH Then
      FLEX.Rows = FLEX.Rows + 1
    End If
    
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
    Call CLEARDATA
     If M_GRD1.Enabled Then M_GRD1.SetFocus
     If SUBPKGNG.Enabled Then SUBPKGNG.SetFocus
     
    cmdOk.Caption = "&Add"
    SWITCH = False
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
Dim WGTRANGEREQ As Boolean
    
    WGTRANGEREQ = True
    If ShadeLabelDisplay(DIVCOD, UNCD) = "Shade" Then
       WGTRANGEREQ = False
    End If

    If Trim(M_GRD1) = Empty Then
        MsgBox "Enter SubGrade", vbInformation
        M_GRD1.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If M_RAT = Empty Then
        MsgBox "Enter Rate Difference !!", vbInformation, "Rate Difference Missing !!"
        M_RAT.SetFocus
        CheckData = True
        Exit Function
    End If
      
    If WGTRANGEREQ Then
    If SEQNO = Empty Then
        MsgBox "Enter Sequence Number !!", vbInformation, "Sequence Missing !!"
        SEQNO.SetFocus
        CheckData = True
        Exit Function
    End If
    End If
    
    If WGTRANGEREQ Then
        If Val(TXTENDWGT) < Val(TXTSTARTWGT) Or Val(TXTENDWGT) = Val(TXTSTARTWGT) Then
            MsgBox "Enter Valid Weight range!!", vbInformation, "Ending Weight must be greater than Starting Weight !!"
            TXTSTARTWGT.SetFocus
            CheckData = True
            Exit Function
        End If
    End If
    
    For INDEX = 1 To FLEX.Rows - 1
        If Trim(FLEX.TextMatrix(INDEX, 0)) = Trim(M_GRD1) And Trim(FLEX.TextMatrix(INDEX, 6)) = Trim(SUBPKGNG.Text) And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
           If IsShadeReq Then
              MsgBox "Shade Already Exist with this name", vbCritical
           Else
              MsgBox "Invalid Sub-Grade"
           End If
           M_GRD1.SetFocus
           CheckData = True
           Exit Function
        End If
        If WGTRANGEREQ Then
            If Trim(FLEX.TextMatrix(INDEX, 4)) = Trim(SEQNO) And Trim(FLEX.TextMatrix(INDEX, 6)) = Trim(SUBPKGNG.Text) And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
               MsgBox "Invalid Sequence No."
               SEQNO.SetFocus
               CheckData = True
               Exit Function
            End If
            
            If Val(TXTSTARTWGT) >= Val(FLEX.TextMatrix(INDEX, 1)) And Val(TXTENDWGT) <= Val(FLEX.TextMatrix(INDEX, 2)) And Trim(FLEX.TextMatrix(INDEX, 6)) = Trim(SUBPKGNG.Text) And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
               MsgBox "Invalid Starting - Ending Range."
               TXTSTARTWGT.SetFocus
               CheckData = True
               Exit Function
            End If
        End If
    Next INDEX
    
End Function

Private Sub CLEARDATA()
    M_GRD1.Text = Empty
    TXTSTARTWGT = Empty
    TXTENDWGT = Empty
    M_RAT = Empty
    SEQNO = Empty
    SUBPKGNG.Text = Empty
End Sub

Private Sub SEQNO_LostFocus()
SEQNO.BackColor = vbWhite
End Sub

Private Sub Text1_Change()

End Sub

Private Sub SUBPKGNG_GotFocus()
SUBPKGNG.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SUBPKGNG_KeyDown(KeyCode As Integer, Shift As Integer)
  If Trim(SUBPKGNG.Text) = Empty Or KeyCode = vbKeyF2 Then
    SUBPKGNG.Text = SearchList1("select TOP 20  CODE,NAME from SUBPKGNGMST WHERE RECSTAT <> 'D'", 0, SUBPKGNG.Text, "SELECT SUB PACKAGING TYPE FROM MASTER")
    End If
  If KeyCode = vbKeyDelete Then SUBPKGNG.Text = Empty
End Sub

Private Sub SUBPKGNG_LostFocus()
SUBPKGNG.BackColor = vbWhite
End Sub

Private Sub TXTENDWGT_GotFocus()
TXTENDWGT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTENDWGT_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTENDWGT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTENDWGT_LostFocus()
TXTENDWGT.BackColor = vbWhite
End Sub

Private Sub TXTSTARTWGT_GotFocus()
TXTSTARTWGT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSTARTWGT_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTSTARTWGT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTSTARTWGT_LostFocus()
TXTSTARTWGT.BackColor = vbWhite
End Sub

Private Sub SetShadeInfo()
  
  LBLSHADE.Caption = "Shade"
  LBLHEAD.Caption = "Shade Master"
  Me.Caption = "Shade Master With rate Difference"
  LBLGRADE.Enabled = False
  M_GRAD.Enabled = False
  LBLGRADE.Enabled = False
  M_GRAD.Enabled = False
  LBLSUB.Enabled = False
  SUBPKGNG.Enabled = False
  FLEX.TextMatrix(0, 0) = "Shade"
  
  FLEX.Rows = 1
  FLEX.Rows = 2
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT DISTINCT NAME,SWGT,EWGT,RDIFF,SEQNO FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND  DVCD='" & DIVCOD & "' AND RECSTAT='A'", CN, adOpenKeyset, adLockPessimistic
  Do While Not RS.EOF
    
    SAVEFLAG = False
    ROWNO = FLEX.Rows - 1
    FLEX.TextMatrix(ROWNO, 0) = Trim(RS!NAME & "")
    FLEX.TextMatrix(ROWNO, 1) = Format(STR(RS!SWGT), "000.000")
    FLEX.TextMatrix(ROWNO, 2) = Format(STR(RS!EWGT), "000.000")
    FLEX.TextMatrix(ROWNO, 3) = RS!RDIFF
    FLEX.TextMatrix(ROWNO, 4) = RS!SEQNO
    FLEX.Rows = FLEX.Rows + 1
    RS.MoveNext
  Loop

  RS.Close
  
     
End Sub

Private Sub SaveShadeInfo()
On Error GoTo LAST
Dim INDEX As Long

CN.BeginTrans

CN.Execute "DELETE FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND  DVCD='" & DIVCOD & "' ", INDEX
           
If INFORS.State = 1 Then INFORS.Close
INFORS.Open "SELECT * FROM GRDMST ", CN, adOpenDynamic, adLockOptimistic
Do While Not INFORS.EOF

     With FLEX
        For INDEX = 1 To FLEX.Rows - 2
        
            CN.Execute "INSERT INTO SUBGRDMST(COMP,UNIT,DVCD,GRAD,SUBGRD,NAME,SWGT,EWGT,RDIFF,SEQNO," & _
            "STATUS,RECSTAT,SUBPKGCODE) VALUES('" & compPth & "','" & UNCD & "','" & DIVCOD & _
            "','" & INFORS!CODE & "','" & INDEX & "','" & .TextMatrix(INDEX, 0) & _
            "','" & .TextMatrix(INDEX, 1) & "','" & .TextMatrix(INDEX, 2) & "','" & .TextMatrix(INDEX, 3) & _
            "','" & .TextMatrix(INDEX, 4) & "','A','A','" & Trim(SUBPKGCODE) & "')"
            
        Next INDEX
     End With
   
INFORS.MoveNext
Loop
INFORS.Close

CN.CommitTrans

MsgBox "Shade Details Saved Successfully", vbInformation

Call RESETALL
Unload Me

Exit Sub
LAST:
  
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub

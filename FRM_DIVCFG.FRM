VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FRM_DIVCFG 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Division Configuration"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6105
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   14
      Top             =   7080
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
         TabIndex        =   15
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Last Dispatch No."
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   5895
      Begin VB.TextBox DPF_EXTRA 
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   10
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_WASTAGE 
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
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_EXPORT 
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
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_MARKET 
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
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_CAPTIVE 
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_JOB 
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
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Extra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Wastage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Captive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Job"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Market"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1560
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
      Image           =   "FRM_DIVCFG.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   1560
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
      Image           =   "FRM_DIVCFG.frx":059A
      cBack           =   -2147483633
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
      TabIndex        =   17
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape BORDER1 
      BorderColor     =   &H80000002&
      Height          =   300
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5895
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
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FRM_DIVCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DIVCODE As String
Dim Ctrl

Private Sub CMDEXIT_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo CFGERR
For Each Ctrl In Me
  If TypeOf Ctrl Is TextBox Then
     If Len(Ctrl) <> 6 Then MsgBox "Enter 6 Digit Serial.": Ctrl.SetFocus: Exit Sub
  End If
Next

CN.BeginTrans
Call SetData
CN.CommitTrans
        
Unload Me

Exit Sub
CFGERR:
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub DPF_EXTRA_GotFocus()
    DPF_EXTRA.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_EXTRA_LostFocus()
    DPF_EXTRA.BackColor = vbWhite
End Sub

Private Sub Form_Activate()

  If DIVCODE = Empty Or LBLDIV.Caption = Empty Then
     Unload Me
  End If
  
  Dim PASS
  PASS = InputBox("Enter the Password", "Enter Password")
    
  If UCase(PASS) <> "9825198251" Or UCase(PASS) = Empty Then
     Unload Me
  End If
  
  On Error Resume Next
  
  Dim STFY As String, ENFY As String, SRNO As String
  STFY = Format(FSDT, "YYYY/MM/DD")      'Mid(Year(FSDT), 1, 4)
  ENFY = Format(FEDT, "YYYY/MM/DD")      'Mid(Year(FSDT), 1, 4)
  SRNO = "000000" & FYCD
  
  Dim SERIALRS As ADODB.Recordset
  Set SERIALRS = New ADODB.Recordset
  
  Dim INRS As ADODB.Recordset
  Set INRS = New ADODB.Recordset
  
  If SERIALRS.State = 1 Then SERIALRS.Close
  SERIALRS.Open "SELECT DISTINCT CODE,NAME FROM SERIALMASTER WHERE VTYP='DPF' AND CODE<>'' AND NAME<>''", CN, adOpenDynamic, adLockOptimistic
  Do While Not SERIALRS.EOF
  
     If INRS.State = 1 Then INRS.Close
     INRS.Open "SELECT * FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND CODE='" & SERIALRS!CODE & _
               "' AND NAME='" & SERIALRS!NAME & "' ", CN, adOpenDynamic, adLockOptimistic
     If INRS.EOF Then
        CN.Execute "INSERT INTO SERIALMASTER(COMP,UNIT,DVCD,VTYP,CODE,NAME,SRNO,FYCD,STFY,ENFY) " & _
             "VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
             "','DPF','" & SERIALRS!CODE & "','" & SERIALRS!NAME & "','" & SRNO & _
             "','" & FYCD & "','" & STFY & "','" & ENFY & "')"
     End If
       
  SERIALRS.MoveNext
  Loop
  SERIALRS.Close
  
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If TypeOf ActiveControl Is TextBox Then
   If Len(Trim(ActiveControl.Text)) <> 6 Then Exit Sub
End If
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
Dim CFGRS As New ADODB.Recordset
Set CFGRS = New ADODB.Recordset

Call CenterChild(frm_Main, Me)

M_DESC = Empty: Key = Empty: NEW_VISIBLE = False: DIVCODE = Empty
If DIVCODE = Empty Then
   LBLDIV.Caption = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
   DIVCODE = Key
End If

If DIVCODE <> Empty Then
   Call GetData
   
   For Each Ctrl In Me
        If TypeOf Ctrl Is Frame Then
            Ctrl.BackColor = RGB(RED, GREEN, BLUE)
            Ctrl.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        ElseIf TypeOf Ctrl Is Label Then
            Ctrl.BackColor = RGB(RED, GREEN, BLUE)
            Ctrl.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        End If
        Me.BackColor = RGB(RED, GREEN, BLUE)
    Next
End If

End Sub

Private Sub DPF_CAPTIVE_GotFocus()
  DPF_CAPTIVE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_CAPTIVE_LostFocus():  DPF_CAPTIVE.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_EXPORT_GotFocus()
  DPF_EXPORT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_EXPORT_LostFocus():  DPF_EXPORT.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_JOB_GotFocus()
  DPF_JOB.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_JOB_LostFocus():  DPF_JOB.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_MARKET_GotFocus()
  DPF_MARKET.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_MARKET_LostFocus():  DPF_MARKET.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_WASTAGE_GotFocus()
  DPF_WASTAGE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_WASTAGE_LostFocus():  DPF_WASTAGE.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub GetData()
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT CODE,SRNO AS LVNO,VTYP,ACTIVE FROM SERIALMASTER WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
If GETRS.EOF Then Exit Sub

Do While Not GETRS.EOF

    Select Case UCase(Trim(GETRS!CODE))
        Case "000001"
             DPF_MARKET = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
        Case "000002"
             DPF_EXPORT = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
        Case "000003"
             DPF_JOB = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
        Case "000004"
             DPF_CAPTIVE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
        Case "000005"
             DPF_WASTAGE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
        Case "000006"
             DPF_EXTRA = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    End Select

    GETRS.MoveNext
Loop
GETRS.Close
End Sub

Private Sub SetData()
Dim SETRS As ADODB.Recordset
Set SETRS = New ADODB.Recordset

Dim AddSQL As String
AddSQL = " WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND FYCD ='" & FYCD & "' AND VTYP='DPF' "

If SETRS.State = 1 Then SETRS.Close
SETRS.Open "SELECT CODE,SRNO,VTYP,ACTIVE FROM SERIALMASTER " & AddSQL, CN, adOpenDynamic, adLockOptimistic
If SETRS.EOF Then Exit Sub

Dim FY As String
FY = FYCD

Do While Not SETRS.EOF

    Select Case UCase(Trim(SETRS!CODE))
    Case "000001"
         CN.Execute "UPDATE SERIALMASTER SET SRNO='" & DPF_MARKET & FY & "' " & AddSQL & _
                    " AND CODE ='" & UCase(Trim(SETRS!CODE)) & "' "
    Case "000002"
         CN.Execute "UPDATE SERIALMASTER SET SRNO='" & DPF_EXPORT & FY & "' " & AddSQL & _
                    " AND CODE ='" & UCase(Trim(SETRS!CODE)) & "' "
    Case "000003"
         CN.Execute "UPDATE SERIALMASTER SET SRNO='" & DPF_JOB & FY & "' " & AddSQL & _
                    " AND CODE ='" & UCase(Trim(SETRS!CODE)) & "' "
    Case "000004"
         CN.Execute "UPDATE SERIALMASTER SET SRNO='" & DPF_CAPTIVE & FY & "' " & AddSQL & _
                    " AND CODE ='" & UCase(Trim(SETRS!CODE)) & "' "
    Case "000005"
         CN.Execute "UPDATE SERIALMASTER SET SRNO='" & DPF_WASTAGE & FY & "' " & AddSQL & _
                    " AND CODE ='" & UCase(Trim(SETRS!CODE)) & "' "
    Case "000006"
         CN.Execute "UPDATE SERIALMASTER SET SRNO='" & DPF_EXTRA & FY & "' " & AddSQL & _
                    " AND CODE ='" & UCase(Trim(SETRS!CODE)) & "' "
    End Select

SETRS.Update
SETRS.MoveNext
Loop

End Sub


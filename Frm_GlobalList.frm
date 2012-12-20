VERSION 5.00
Begin VB.Form Frm_GlobalList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List"
   ClientHeight    =   3945
   ClientLeft      =   1800
   ClientTop       =   3045
   ClientWidth     =   8085
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_GlobalList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3405
      TabIndex        =   5
      Top             =   3540
      Width           =   975
   End
   Begin VB.ListBox lstCode 
      Height          =   1035
      Left            =   1005
      TabIndex        =   4
      Top             =   1410
      Visible         =   0   'False
      Width           =   1575
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
      Height          =   300
      Left            =   1230
      TabIndex        =   3
      Top             =   3540
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   2
      Top             =   3540
      Width           =   975
   End
   Begin VB.ListBox lstName 
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   4245
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   7830
   End
   Begin VB.Frame framDetails 
      Caption         =   "Details"
      Height          =   3360
      Left            =   4455
      TabIndex        =   6
      Top             =   435
      Width           =   3495
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax Catagoery"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1800
      End
      Begin VB.Label M_RTTX 
         Caption         =   "Retail/Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2160
         TabIndex        =   22
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Label lblAdd3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   21
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblAdd2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblAdd1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   885
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label txtAcGrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1200
         TabIndex        =   18
         Top             =   1344
         Width           =   1320
      End
      Begin VB.Label txtCPCD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPCD of Party"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1290
         TabIndex        =   17
         Top             =   1667
         Width           =   1020
      End
      Begin VB.Label txtArea 
         Caption         =   "Area of the Party"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   1990
         Width           =   2145
      End
      Begin VB.Label txtBroker 
         Caption         =   "Agent of the Party"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   840
         TabIndex        =   15
         Top             =   2358
         Width           =   2115
      End
      Begin VB.Label txtCURB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1740
         TabIndex        =   14
         Top             =   2700
         Width           =   1635
      End
      Begin VB.Label lblCurb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Balance:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2700
         Width           =   3135
      End
      Begin VB.Label lblAgent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2358
         Width           =   630
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1990
         Width           =   480
      End
      Begin VB.Label lblCPCD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Cpcd:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1667
         Width           =   1095
      End
      Begin VB.Label lblAcGrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/c Group:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1344
         Width           =   945
      End
      Begin VB.Label txtAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Address of the Party"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4080
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
End
Attribute VB_Name = "Frm_GlobalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public open_flag As Boolean
Public form_Para As String
Public SQL As String

Private Sub cmdAddNew_Click()
    
    If IsImportActive = False Then
        Key = ""
        MS_KEY = ""
        M_DESC = ""
        sTxt = ""
        key_PressNew = True
        Unload Me
    Else
        If IsObject(frm_Ac) Then
            If MsgBox("Are You Sure For New A/c ", vbYesNo + vbQuestion, "Create New A/c ?") = vbYes Then
              AddNewAccount = True
             Else
              Exit Sub
            End If
        End If
    End If
    
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    
    key_PressNew = False
    M_DESC = ""
    Unload Me
    
End Sub

Private Sub cmdOk_Click()
    If lstName.ListIndex = -1 Then Exit Sub
    Me.Hide
    Key = lstCode.List(lstName.ListIndex)
    MS_KEY = lstCode.List(lstName.ListIndex)
    M_DESC = Trim(lstName.List(lstName.ListIndex))
    lstCode.Clear
    lstName.Clear
    txtName = Empty
    Unload Frm_GlobalList
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
'QRY = SQL
On Error GoTo ERRACTIVATE
Dim TEMPRS As New ADODB.Recordset
Dim ANS As String
Dim SRCHITM As String
Dim I As Long
    If Trim(sTxt) = Empty Then
        Set TEMPRS = New Recordset
        If InStr(1, QRY, "SERIALMASTER") <> 0 Then
          TEMPRS.Open QRY & " ORDER BY NAME", DEST_CN, adOpenKeyset, adLockPessimistic
        ElseIf InStr(1, QRY, "REFMST") <> 0 Then
          TEMPRS.Open QRY & " ORDER BY REFMST.NAME", DEST_CN, adOpenKeyset, adLockPessimistic
        ElseIf InStr(1, QRY, "LOCATION") <> 0 Then
          TEMPRS.Open QRY & " ORDER BY LOCNAME", DEST_CN, adOpenKeyset, adLockPessimistic
        ElseIf InStr(1, QRY, "ORDTRN") <> 0 Then
          TEMPRS.Open QRY & " WHERE VTYP='DOS'", DEST_CN, adOpenKeyset, adLockPessimistic
        ElseIf InStr(1, QRY, "ordtrn") <> 0 Then
          TEMPRS.Open QRY & " ", DEST_CN, adOpenKeyset, adLockPessimistic
        Else
          TEMPRS.Open QRY, DEST_CN, adOpenKeyset, adLockPessimistic
        End If
        TEMPRS.Requery

        'If Not temprs.EOF Then temprs.MoveLast

        If TEMPRS.EOF = True Then
            cmdCancel.Visible = True
            cmdCancel.Cancel = True
            open_flag = True
            M_DESC = "X"
            MsgBox "There are No record.", vbInformation, App.Title
        Else
            'temprs.MoveFirst
            lstCode.Clear
            lstName.Clear
            Do While Not TEMPRS.EOF
                lstCode.AddItem Trim(TEMPRS.Fields(0).Value)
                lstName.AddItem Trim(TEMPRS.Fields(1).Value)
                TEMPRS.MoveNext
            Loop
          End If

        TEMPRS.Close
      End If
        If IsImportActive Then
            txtName = Empty
            txtName.Text = sTxt
        End If

        If IsImportActive Then
            If lstName.ListIndex > 0 Then
                CMDOK.Default = True
            Else
                If cmdAddNew.Visible And txtName <> Empty Then
                    cmdAddNew.Default = True
                Else
                    CMDOK.Default = True
                End If
            End If
        Else
            CMDOK.Default = True
        End If
    
        If lstName.ListCount <> 0 Then
            lstName.ListIndex = 0
        Else
            CANCEL_VISIBLE = True
        End If
        
        If open_flag = False And Me.Visible = False Then
            Me.Show
        End If
        
        key_PressNew = False
        
        SetFormSize
        
        If IsImportActive And sTxt <> Empty Then
            Dim M_STRTMP  As String
            M_STRTMP = Mid(sTxt, 1, InStr(1, sTxt, Space(1), vbTextCompare))
            For I = 0 To lstName.ListCount - 1
                If Left(UCase(lstName.List(I)), Len(Trim(M_STRTMP))) = UCase(M_STRTMP) Then
                    lstName.ListIndex = I
                    Exit Sub
                End If
            Next
        End If
        
        SRCHITM = txtName
        If txtName <> Empty Then
            For I = 0 To lstName.ListCount - 1
                If UCase(Trim(lstName.List(I))) = UCase(Trim(SRCHITM)) Then
                    lstName.ListIndex = I
                    Exit For
                End If
            Next
        End If
        
        Exit Sub
        
ERRACTIVATE:

    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
    Unload Me
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If IsImportActive Then
        If KeyAscii = vbKeyEscape Then
            IsImportActive = False
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
On Error GoTo errLoad
    If lstCaption = Empty Then Me.Caption = "List" Else Me.Caption = lstCaption

    
    Call CenterChild(frm_Main, Me)
    
    cmdAddNew.Visible = NEW_VISIBLE
    cmdCancel.Visible = CANCEL_VISIBLE
    
    If IsImportActive Then
        Me.Caption = "Select Account Name For"
        cmdAddNew.Visible = True
        Key = Empty
        MS_KEY = Empty
    End If
    
    txtName = sTxt
  Exit Sub

errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Msg Empty
    lstCaption = Empty
    sTxt = Empty
    
End Sub

Private Sub lstName_Click()
    dispDetails lstCode.List(lstName.ListIndex)
End Sub

Private Sub lstName_DblClick()
    Call cmdOk_Click
End Sub

Private Sub lstName_GotFocus()
    Msg "Select Relevant Value From List"
End Sub

Private Sub TXTNAME_Change()
Dim TEMPRS As New ADODB.Recordset, SQL As String
SQL = QRY
   lstName.Clear
   lstCode.Clear
    
    If VBA.Strings.InStr(UCase(SQL), "WHERE") <> 0 Then
        If InStr(1, SQL, "ITMMST") <> 0 Then
            SQL = SQL & " AND ITMMST.[NAME] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "COMPMAST") <> 0 Then
            SQL = SQL & " AND COMPMAST.[COMP_NAME] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "TXULOT") <> 0 Then
            SQL = SQL & " AND TXULOT.[LTNO] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "MFGSHDMST") <> 0 Then
            SQL = SQL & " AND MFGSHDMST.[COLORNO] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "TTYP") <> 0 Then
            SQL = SQL & " AND ACCMST.[TTYP] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "JOBCARDNO") <> 0 Then
            SQL = SQL & " AND JOB_SUMMARY.[JOBCARDNO] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "MRGMST") <> 0 Then
            SQL = SQL & " AND MRGMST.[MRGN] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "TRNMAN") <> 0 Then
            SQL = SQL & " AND TRNMAN.[VTYP] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "grdmst") <> 0 Then
            SQL = SQL & " AND GRDMST.[GRAD] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "EXTRA4") <> 0 Then
            SQL = SQL & " AND EXTRA4 LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "BOXREG") <> 0 Then
            SQL = SQL & " AND BOXREG.[GRAD] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "LOCATION") <> 0 Then
            SQL = SQL & " AND LOCNAME LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "ORDTRN") <> 0 Then
            SQL = SQL & " AND DONO LIKE ('" & txtName.Text & "%') AND VTYP='DOS'"
        ElseIf InStr(1, SQL, "ordtrn") <> 0 Then
            SQL = SQL & " AND DONO LIKE ('" & txtName.Text & "%') AND VTYP='DOS'"
        ElseIf InStr(1, SQL, "PROD_CONNING") <> 0 Then
            SQL = SQL & " AND PROD_CONNING.JOBCARDNO LIKE ('" & txtName.Text & "%')"
        Else
            SQL = SQL & " AND [NAME] LIKE ('" & txtName.Text & "%')"
        End If
    Else
        If InStr(1, SQL, "ITMMST") <> 0 Then
            SQL = SQL & " WHERE  ITMMST.[NAME] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "COMPMAST") <> 0 Then
            SQL = SQL & " WHERE  COMPMAST.[COMP_NAME] LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "RATIOMST") <> 0 Then
            SQL = SQL & " WHERE HEAD LIKE '" & txtName & "%'"
        ElseIf InStr(1, SQL, "REPCNF") <> 0 Then
            SQL = SQL & " WHERE RPNM LIKE '" & txtName & "%'"
        ElseIf InStr(1, SQL, "TXULOT") <> 0 Then
            SQL = SQL & " WHERE LTNO LIKE '" & txtName & "%'"
        ElseIf InStr(1, SQL, "MRGMST") <> 0 Then
            SQL = SQL & " WHERE MRGN LIKE '" & txtName & "%'"
        ElseIf InStr(1, SQL, "grdmst") <> 0 Then
            SQL = SQL & " WHERE GRAD LIKE '" & txtName & "%'"
        ElseIf InStr(1, SQL, "EXTRA4") <> 0 Then
            SQL = SQL & " WHERE EXTRA4 LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "BOXREG") <> 0 Then
            SQL = SQL & " WHERE GRAD LIKE '" & txtName & "%'"
        ElseIf InStr(1, SQL, "LOCATION") <> 0 Then
            SQL = SQL & " WHERE LOCNAME LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "ORDTRN") <> 0 Then
            SQL = SQL & " WHERE DONO LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "PROD_CONNING") <> 0 Then
            SQL = SQL & " WHERE PROD_CONNING.JOBCARDNO LIKE ('" & txtName.Text & "%')"
        ElseIf InStr(1, SQL, "REFMST") <> 0 Then
            SQL = SQL & " WHERE REFMST.NAME LIKE ('" & txtName.Text & "%')"
        Else
            If IsImportActive And InStr(1, SQL, "ACCMST") <> 0 And Len(txtName) > 0 Then
                lstCode.Clear
                lstName.Clear
                Dim M_STRTMP  As String
                Dim M_POS As Integer
                
                M_POS = InStr(1, sTxt, Space(1), vbTextCompare) - 1
                If M_POS = 0 Then M_POS = Len(txtName)
                If M_POS > 0 Then
                    M_STRTMP = Mid(txtName, 1, M_POS)
                
'                For i = 0 To lstName.ListCount - 1
'                    If Left(UCase(lstName.list(i)), Len(Trim(M_STRTMP))) = UCase(M_STRTMP) Then
'                        lstName.ListIndex = i
'                        Exit Sub
'                    End If
'                Next
                Else
                    M_STRTMP = txtName
                End If
                    SQL = SQL & " WHERE  [NAME] LIKE ('" & M_STRTMP & "%')"
            Else
                SQL = SQL & " WHERE  [NAME] LIKE ('" & txtName.Text & "%')"
            End If
        End If
    End If
    
    If InStr(1, SQL, "ITMMST") <> 0 Then
        SQL = SQL & " ORDER BY ITMMST.NAME"
    ElseIf InStr(1, SQL, "COMPMAST") <> 0 Then
        SQL = SQL & " ORDER BY COMP_NAME"
    ElseIf InStr(1, SQL, "RATIOMST") <> 0 Then
        SQL = SQL & " ORDER BY ROCD"
    ElseIf InStr(1, SQL, "TRNMAN") <> 0 Then
    
    ElseIf InStr(1, SQL, "TTYP") <> 0 Then
    
    ElseIf InStr(1, SQL, "CHRGMST") <> 0 Then
        SQL = SQL & " Order By Code"
    ElseIf InStr(1, SQL, "REPCNF") <> 0 Then
        SQL = SQL & " ORDER BY RPNM"
    ElseIf InStr(1, SQL, "grdmst") <> 0 Then
        SQL = SQL & " ORDER BY GRAD"
    ElseIf InStr(1, SQL, "EXTRA4") <> 0 Then
        SQL = SQL & ""
    ElseIf InStr(1, SQL, "BOXREG") <> 0 Then
        SQL = SQL & " ORDER BY GRAD"
    ElseIf InStr(1, SQL, "TTYP") <> 0 Then
        SQL = SQL & " ORDER BY TTYP"
    ElseIf InStr(1, SQL, "TXULOT") <> 0 Then
        SQL = SQL & " ORDER BY LTNO"
    ElseIf InStr(1, SQL, "MRGMST") <> 0 Then
        SQL = SQL & " ORDER BY MRGN"
    ElseIf InStr(1, SQL, "GRAD") <> 0 Then
        SQL = SQL & " ORDER BY GRAD"
    ElseIf InStr(1, SQL, "LOCATION") <> 0 Then
        SQL = SQL & " ORDER BY LOCNAME"
    ElseIf InStr(1, SQL, "ORDTRN") <> 0 Then
        SQL = SQL & " ORDER BY DONO"
    ElseIf InStr(1, SQL, "JOBCARDNO") <> 0 Then
       SQL = SQL & " ORDER BY JOBCARDNO"
    ElseIf InStr(1, SQL, "ordtrn") <> 0 Then
        SQL = SQL & " ORDER BY DONO"
    ElseIf InStr(1, SQL, "PROD_CONNING") <> 0 Then
       SQL = SQL & " ORDER BY PROD_CONNING.JOBCARDNO"
    ElseIf InStr(1, SQL, "REFMST") <> 0 Then
       SQL = SQL & " ORDER BY REFMST.NAME"
    ElseIf InStr(1, SQL, "MFGSHDMST") <> 0 Then
       SQL = SQL & " ORDER BY MFGSHDMST.COLORNO"
    Else
        SQL = SQL & " ORDER BY NAME"
    End If
    
    TEMPRS.Open SQL, DEST_CN, adOpenDynamic, adLockOptimistic
    

    If Not TEMPRS.EOF Then TEMPRS.MoveLast
    
    If TEMPRS.EOF = True Then
        lstName.Clear
        lstCode.Clear
    Else
        TEMPRS.MoveFirst
        Do While Not TEMPRS.EOF
            lstCode.AddItem TEMPRS.Fields(0).Value
            lstName.AddItem TEMPRS.Fields(1).Value
            TEMPRS.MoveNext
        Loop
    End If
    
    TEMPRS.Close
    If lstName.ListCount <> 0 Then lstName.ListIndex = 0: CMDOK.Default = True Else txtAddress.Caption = "N/A"
End Sub

Private Sub txtName_GotFocus()
 txtName.BackColor = RGB(BRED, BGREEN, BBLUE)
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub TXTNAME_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstName.SetFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("[")
            KeyAscii = 0
        Case Asc("*")
            KeyAscii = 0
        Case Asc("]")
            KeyAscii = 0
        Case Asc("("), Asc(")")
            KeyAscii = 0
        Case Asc("&"), Asc("!"), Asc("@"), Asc("#"), Asc("^")
            KeyAscii = 0
        Case Asc("'")
            KeyAscii = 0
    End Select
End Sub

Private Sub dispDetails(ByVal cCode As String)
On Error GoTo errDispDetails
Dim TEMPRS As New ADODB.Recordset

    If TEMPRS.State = adStateOpen Then TEMPRS.Close
    
    If InStr(1, UCase(QRY), "WHERE CATA='Y'") <> 0 Then
        TEMPRS.Open "SELECT * FROM REFMST WHERE CATA='Y' AND NAME='" & lstName.Text & "'", DEST_CN
        If TEMPRS.EOF = False Then
            lblAdd1 = Trim(TEMPRS!adL1 & "")
            lblAdd2 = Trim(TEMPRS!ADL2 & "")
            lblAdd3 = TEMPRS!area & ""
        End If
        TEMPRS.Close
    ElseIf InStr(1, UCase(QRY), "TTYP") <> 0 Then
    
    ElseIf InStr(1, UCase(QRY), "ACCMST") <> 0 Then
        
        TEMPRS.Open "SELECT * FROM LIST where code='" & cCode & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        'If Not TEMPRS.EOF Then
        txtAcGrp.Caption = TEMPRS!ACGRP & ""
        txtAddress.Caption = TEMPRS![adro] & ""
        
        lblAdd1.Caption = Mid(txtAddress.Caption, 1, 130)
        lblAdd2.Caption = Mid(txtAddress.Caption, 31, 40)
        lblAdd3.Caption = Mid(txtAddress.Caption, 71, 40)
        lblAdd2.Visible = False
        lblAdd3.Visible = False
        
        If Not IsNull(TEMPRS![CPCD]) Then txtCPCD.Caption = TEMPRS![CPCD] Else txtCPCD.Caption = "N/A"
        If Not IsNull(TEMPRS![area]) Then txtArea.Caption = TEMPRS![area] Else txtArea.Caption = "N/A"
        If Not IsNull(TEMPRS![BROKER]) Then txtBroker.Caption = TEMPRS![BROKER] Else txtBroker.Caption = "N/A"
        'If Not IsNull(TEMPRS!BALN) Then
        '    txtCURB.Caption = Format(TEMPRS![BALN], "##########.00")
        'Else
            txtCURB.Caption = ".00"
            M_RTTX.Caption = TEMPRS!TTYP & ""
        'End If
        
        If TEMPRS.State = 1 Then TEMPRS.Close
        TEMPRS.Open "SELECT BALN FROM ACCBALN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & cCode & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        If Not TEMPRS.EOF Then
          txtCURB.Visible = True
          txtCURB.Caption = "Current Balance : " + Format(TEMPRS!BALN, "##############.00")
         Else
          txtCURB.Caption = "Current Blalance: " + "0.00"
        End If
        TEMPRS.Close
        If TEMPRS.State = 1 Then TEMPRS.Close
        'TEMPRS.Open "SELECT ISNULL(ISNULL(SUM(DAMT),0)-ISNULL(SUM(CAMT),0),0) AS CURB FROM TRNMAN WHERE COMP='" & compPth & "' AND ACOD='" & cCode & "' AND RECSTAT<>'D'", DEST_CN, adOpenDynamic, adLockOptimistic
        'txtCURB.Caption = TEMPRS!curb
        
        'TEMPRS.Close
    ElseIf InStr(1, UCase(QRY), "ITMMST") <> 0 Then
        TEMPRS.Open "SELECT * FROM ITEMLIST WHERE CODE='" & cCode & "' AND COMP='" & compPth & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        Label2.Visible = False
        M_RTTX.Visible = False
        lblAdd2.Visible = False
        lblAdd3.Visible = False
        Dim hlp_blq As Long
        Dim hlp_pqty As Long
        Dim hlp_sqty As Long
        hlp_pqty = 0
        hlp_sqty = 0
        Dim hlp_ignm As String
        
        'If TEMPRS.State = 1 Then TEMPRS.Close
        'TEMPRS.Open "select isnull(sum(qnty),0) as pur_qty from purtran where oper='+' and recstat<>'D' and icod='" & cCode & "' and comp='" & compPth & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        'If Not TEMPRS.EOF Then
        '  hlp_pqty = TEMPRS!pur_qty
        ' Else
        '  hlp_pqty = 0
        'End If
        'If TEMPRS.State = 1 Then TEMPRS.Close
        'TEMPRS.Open "select isnull(sum(qnty),0) as pur_qty from storetran where oper='+' and recstat<>'D' and icod='" & cCode & "' and comp='" & compPth & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        'If Not TEMPRS.EOF Then
        '  hlp_pqty = hlp_pqty + TEMPRS!pur_qty
        ' Else
        '  hlp_pqty = hlp_pqty + 0
        'End If
        'If TEMPRS.State = 1 Then TEMPRS.Close
        'TEMPRS.Open "select isnull(sum(qnty),0) as pur_qty from sptran where oper='+' and recstat<>'D' and icod='" & cCode & "' and comp='" & compPth & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        'If Not TEMPRS.EOF Then
        '  hlp_pqty = hlp_pqty + TEMPRS!pur_qty
        ' Else
        '  hlp_pqty = hlp_pqty + 0
        'End If
       '
        'If TEMPRS.State = 1 Then TEMPRS.Close
        'TEMPRS.Open "select isnull(sum(qnty),0) as sal_qty from purtran where oper='-' and recstat<>'D' and icod='" & cCode & "' and comp='" & compPth & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        'If Not TEMPRS.EOF Then
        '  hlp_sqty = TEMPRS!sal_qty
        ' Else
        '  hlp_sqty = 0
        'End If
        'If TEMPRS.State = 1 Then TEMPRS.Close
        'TEMPRS.Open "select isnull(sum(qnty),0) as sal_qty from storetran where oper='-' and recstat<>'D' and icod='" & cCode & "' and comp='" & compPth & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        'If Not TEMPRS.EOF Then
         ' hlp_sqty = hlp_sqty + TEMPRS!sal_qty
         'Else
         ' hlp_sqty = hlp_sqty + 0
        'End If
        'If TEMPRS.State = 1 Then TEMPRS.Close
        'TEMPRS.Open "select isnull(sum(qnty),0) as sal_qty from sptran where oper='-' and recstat<>'D' and icod='" & cCode & "' and comp='" & compPth & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        'If Not TEMPRS.EOF Then
        '  hlp_sqty = hlp_sqty + TEMPRS!sal_qty
        ' Else
        '  hlp_sqty = hlp_sqty + 0
        'End If
        'hlp_blq = hlp_pqty - hlpsqty
        'blAddress.Caption = "BALANCE:"
        'lblAddress.AutoSize = True
        'lblAdd1.Caption = Space(5) & Format(hlp_blq, "###########.000")
        lblAdd1.Caption = Empty
        lblAcGrp.Visible = False
        If TEMPRS.State = 1 Then TEMPRS.Close
        Dim IGMCOD As String
        TEMPRS.Open "SELECT * FROM ITMMST WHERE CODE='" & cCode & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        If Not TEMPRS.EOF Then
          IGMCOD = TEMPRS!igcd
         Else
          IGMCOD = ""
        End If
        If TEMPRS.State = 1 Then TEMPRS.Close
        TEMPRS.Open "SELECT * FROM IGMMST WHERE CODE='" & IGMCOD & "'", DEST_CN, adOpenDynamic, adLockOptimistic
        If Not TEMPRS.EOF Then
          lblAcGrp.Visible = True
          lblAcGrp.Caption = "Item Group : " + TEMPRS!NAME & ""
        End If
        'lblAcGrp.Caption = "RATE : " & Format(TEMPRS!Rate, "#.00")
        TEMPRS.Close
    End If
    
    Exit Sub
    
errDispDetails:
    
    
    'Resume
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub SetFormSize()
   
    If InStr(1, UCase(QRY), "ACCMST") <> 0 Then
        lblAcGrp.Visible = True
        txtAcGrp.Visible = True
        lblCPCD.Visible = True
        txtCPCD.Visible = True
        lblArea.Visible = True
        txtArea.Visible = True
        lblAgent.Visible = True
        txtBroker.Visible = True
        
        lblCurb.Visible = True
        txtCURB.Visible = True
        lblCurb.Visible = False
    ElseIf InStr(1, UCase(QRY), "ITMMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False
        
        lblAddress.Caption = "BALANCE :"
        txtAddress.Caption = "  0.00"
        
    ElseIf InStr(1, UCase(QRY), "IGMMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False
        framDetails.Visible = False
        Me.WIDTH = lstName.WIDTH + 200
    ElseIf InStr(1, QRY, "WHERE CATA='Y'", vbTextCompare) > 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False
        lblAdd1.Visible = True
        lblAdd2.Visible = True
        lblAdd3.Visible = True
        framDetails.Visible = True
        Me.WIDTH = 8205

    ElseIf InStr(1, UCase(QRY), "REFMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        txtName.WIDTH = lstName.WIDTH
        Me.WIDTH = lstName.WIDTH + 350

    ElseIf InStr(1, UCase(QRY), "GRPMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        txtName.WIDTH = lstName.WIDTH
        Me.WIDTH = lstName.WIDTH + 350


    ElseIf InStr(1, UCase(QRY), "HEDMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        txtName.WIDTH = lstName.WIDTH
        Me.WIDTH = lstName.WIDTH + 350
    
    Else
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        txtName.WIDTH = lstName.WIDTH
        Me.WIDTH = lstName.WIDTH + 350
    End If
    If InStr(1, UCase(QRY), "TTYP") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        txtArea.Visible = False
        lblAgent.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        txtName.WIDTH = lstName.WIDTH
        Me.WIDTH = lstName.WIDTH + 350
    End If
End Sub

Private Function getRate(sItemCode As String, sCompName As String) As Double
Dim rsTemp As Recordset

    Set rsTemp = New Recordset
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "Select AVG(RATE) AS AVGRATE FROM SPTRAN WHERE VTYP='PUR' AND ICOD='" & sItemCode & "' AND LEFT(VBNO,1)<>'*' AND COMP='" & compPth & "'", DEST_CN
    
    If rsTemp.EOF = False And IsNull(rsTemp!AVGRATE) = False Then getRate = rsTemp!AVGRATE
    
    rsTemp.Close
    
End Function

Private Sub TXTNAME_LostFocus()
txtName.BackColor = vbWhite
End Sub

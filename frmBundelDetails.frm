VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHB~1.OCX"
Begin VB.Form frmBundelDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOX ENTRY"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6015
   Begin VB.Frame ITMFRM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   60
      TabIndex        =   1
      Top             =   465
      Width           =   5925
      Begin WelchButton.lvButtons_H CMDREMOVE 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmBundelDetails.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Close"
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
         Image           =   "frmBundelDetails.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtMts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox txtPcs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   5040
         Width           =   915
      End
      Begin MSFlexGridLib.MSFlexGrid mfgBndlDet 
         Height          =   4290
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   7567
         _Version        =   393216
         Cols            =   6
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPcs 
         AutoSize        =   -1  'True
         Caption         =   "Total Box"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   4800
         Width           =   825
      End
      Begin VB.Label lblMtrs 
         AutoSize        =   -1  'True
         Caption         =   "Total Net Wt."
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
         Left            =   4080
         TabIndex        =   5
         Top             =   4785
         Width           =   1170
      End
   End
   Begin VB.Label lblItemName 
      AutoSize        =   -1  'True
      Caption         =   "Item Name"
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
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   5310
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBundelDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Emptycell As Boolean
Public BndlType As Integer '0-NOT APPLICABLE, 1-DESIGN NO * ROLLS * SUIT, 2-PCS * METER ,3-PCS
Dim DRS As New ADODB.Recordset

Private Sub SetFlexHead()
    mfgBndlDet.Clear
    mfgBndlDet.Rows = 100
    mfgBndlDet.Cols = 5
    
    mfgBndlDet.ColWidth(0) = 400
    mfgBndlDet.TextMatrix(0, 0) = "Sr."
    
    
        mfgBndlDet.ColWidth(1) = 1700
        mfgBndlDet.TextMatrix(0, 0) = "Box No."
    
        mfgBndlDet.ColWidth(2) = 1100
        mfgBndlDet.TextMatrix(0, 0) = "Gross Wt."
    
        mfgBndlDet.ColWidth(3) = 1100
        mfgBndlDet.TextMatrix(0, 0) = "Tare Wt."
        
        mfgBndlDet.ColWidth(4) = 1100
        mfgBndlDet.TextMatrix(0, 0) = "Net Wt."
        
        
        lblPcs.Caption = "Total"
        lblMtrs.Visible = True
        txtMts.Visible = True
        
    
    For I = 1 To mfgBndlDet.Rows - 1
        mfgBndlDet.TextMatrix(I, 0) = I
    Next
End Sub

Private Sub cmdClose_Click()
Dim I As Long
Dim J As Long
    mfgBndlDet.ColWidth((FRMBOXGRN.Flex.ROW * 4) - 3) = 0
    mfgBndlDet.ColWidth((FRMBOXGRN.Flex.ROW * 4) - 2) = 0
    mfgBndlDet.ColWidth((FRMBOXGRN.Flex.ROW * 4) - 1) = 0
    mfgBndlDet.ColWidth(FRMBOXGRN.Flex.ROW * 4) = 0
    
    Call mfgBndlDet_LostFocus
    frmBundelDetails.Hide
    
    If FRMBOXGRN.Flex.Enabled = True Then FRMBOXGRN.Flex.COL = 9: FRMBOXGRN.Flex.SetFocus
    If Val(txtPcs.Text) = 0 Then
      ' FRMBOXGRN.Flex.TextMatrix(FRMBOXGRN.Flex.ROW, 8) = 0
        FRMBOXGRN.Flex.TextMatrix(FRMBOXGRN.Flex.ROW, 7) = Val(txtPcs.Text)
        
    End If
    
    If (Val(txtMts.Text) <> 0) Then
    FRMBOXGRN.Flex.TextMatrix(FRMBOXGRN.Flex.ROW, 8) = Val(txtMts.Text)
    Else
    FRMBOXGRN.Flex.TextMatrix(FRMBOXGRN.Flex.ROW, 8) = 0
    End If
    FRMBOXGRN.Flex.COL = 9
   ' Call FRMBOXGRN.Flex_LeaveCell
End Sub

Private Sub cmdRemove_Click()
If mfgBndlDet.ROW > 0 Then
mfgBndlDet.RemoveItem (mfgBndlDet.ROW)
For I = 1 To mfgBndlDet.Rows - 1
mfgBndlDet.TextMatrix(I, 0) = I
Next
End If
Call mfgBndlDet_LeaveCell
End Sub

Private Sub Form_Activate()
    Me.KeyPreview = False
    Call mfgBndlDet_GotFocus
End Sub

Private Sub Form_Deactivate()
'    Call cmdClose_Click
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    Call CenterChild(frm_Main, Me)
    Call SetFlexHead
    Me.KeyPreview = False
End Sub

Private Sub mfgBndlDet_EnterCell()
    mfgBndlDet.CellBackColor = RGB(BRED, BGREEN, BBLUE)
    Emptycell = True
End Sub

Private Sub mfgBndlDet_GotFocus()
    Dim NewColNo As Integer
    
    If FRMBOXGRN.Flex.TextMatrix(FRMBOXGRN.Flex.ROW, 3) = "" Then
        frmBundelDetails.Hide
        FRMBOXGRN.Flex.COL = 1
        Exit Sub
    End If
    
    For I = 1 To mfgBndlDet.Cols - 1
        mfgBndlDet.ColWidth(I) = 0
    Next
    
    Me.KeyPreview = False
    NewColNo = (FRMBOXGRN.Flex.ROW * 4) - 3
    
    If mfgBndlDet.Cols < (FRMBOXGRN.Flex.ROW * 4) + 1 Then
        mfgBndlDet.Cols = mfgBndlDet.Cols + 4
    End If
    
    mfgBndlDet.ColWidth(0) = 400
    mfgBndlDet.TextMatrix(0, 0) = "Sr."
    
    
        mfgBndlDet.ColWidth(NewColNo) = 1700
        mfgBndlDet.TextMatrix(0, NewColNo) = "Box No."
    
        mfgBndlDet.ColWidth(NewColNo + 1) = 1100
        mfgBndlDet.TextMatrix(0, NewColNo + 1) = "Gross.Wt."
    
        mfgBndlDet.ColWidth(NewColNo + 2) = 1100
        mfgBndlDet.TextMatrix(0, NewColNo + 2) = "Tare Wt."
        
        mfgBndlDet.ColWidth(NewColNo + 3) = 1100
        mfgBndlDet.TextMatrix(0, NewColNo + 3) = "Net Wt."
        
        lblPcs.Caption = "Total Box"
        lblMtrs.Visible = True
        txtMts.Visible = True
        txtPcs.Text = ""
        txtMts.Text = ""
        
    Call mfgBndlDet_LeaveCell
    mfgBndlDet.COL = NewColNo
    mfgBndlDet.ROW = 1
    If BndlType <> 1 Then
        mfgBndlDet.COL = NewColNo + 1
    End If
    If mfgBndlDet.Visible = True Then mfgBndlDet.SetFocus
End Sub

Private Sub mfgBndlDet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        mfgBndlDet.ROW = mfgBndlDet.ROW - 1
    ElseIf KeyCode = vbKeyDown Then
        mfgBndlDet.ROW = mfgBndlDet.ROW + 1
    
    End If
End Sub

Private Sub mfgBndlDet_KeyPress(KeyAscii As Integer)
    Dim ALLOW_KEY As Boolean
    Dim FWD_COL As Boolean
    Dim ENTER_PRESS As Boolean

    Dim MSTDAT As New ADODB.Recordset
    Set MSTDAT = New ADODB.Recordset

    FWD_COL = False
    ALLOW_KEY = False

    If KeyAscii = vbKeyEscape Then
        Call mfgBndlDet_LostFocus
        Exit Sub
    End If
    
 If mfgBndlDet.COL = 2 Or mfgBndlDet.COL = 3 Or mfgBndlDet.COL = 4 Then
    If InStr(1, mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL), ".") > 0 And KeyAscii = 46 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
    

    If mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) Then
        If InStr(1, mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL), ".") > 0 And KeyAscii = 46 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    Select Case mfgBndlDet.COL
        Case (FRMBOXGRN.Flex.ROW * 4) - 3
            If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
                    ALLOW_KEY = True
            ElseIf KeyAscii = 46 Then                              '.
                    ALLOW_KEY = True
             ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then         ' A-Z
                    ALLOW_KEY = True
             ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then         'a-z
                    ALLOW_KEY = True
             ElseIf KeyAscii = 45 Then
                    ALLOW_KEY = True
             Else
                    ALLOW_KEY = False
             End If
             If mfgBndlDet.TextMatrix(mfgBndlDet.ROW, 1) = Empty And mfgBndlDet.TextMatrix(mfgBndlDet.ROW, 2) <> Empty And mfgBndlDet.TextMatrix(mfgBndlDet.ROW, 3) <> Empty And mfgBndlDet.TextMatrix(mfgBndlDet.ROW, 4) <> Empty Then
                MsgBox "Box No. Can Not be Empty", vbOKOnly
                Exit Sub
             End If
            
        Case (FRMBOXGRN.Flex.ROW * 4) - 2
            If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
                ALLOW_KEY = True
            ElseIf KeyAscii = 46 Then
                ALLOW_KEY = True
            Else
                ALLOW_KEY = False
            End If
            
        Case (FRMBOXGRN.Flex.ROW * 4) - 1
            If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
                ALLOW_KEY = True
            ElseIf KeyAscii = 46 Then                              '.
                ALLOW_KEY = True
            Else
                ALLOW_KEY = False
            End If
        Case (FRMBOXGRN.Flex.ROW * 4)
           If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
           '     ALLOW_KEY = True
            ElseIf KeyAscii = 46 Then                              '.
           '     ALLOW_KEY = True
            Else
                ALLOW_KEY = False
            End If
            ALLOW_KEY = False
    End Select
    If KeyAscii = vbKeyReturn Then
        ENTER_PRESS = True
    Else
        ENTER_PRESS = False
    End If
    If KeyAscii = 8 Then
        Dim lnth As Double
        lnth = Len(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL))
        If lnth > 0 Then
            If mfgBndlDet.COL <> 4 Then
            mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL) = Mid(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL), 1, lnth - 1)
            End If
            If mfgBndlDet.COL = 1 Then mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL) = ""
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
        mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL) = Trim(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL)) + Chr(KeyAscii)
        End If
    End If
    FWD_COL = False

    If ENTER_PRESS = True Then
        Select Case mfgBndlDet.COL
            Case (FRMBOXGRN.Flex.ROW * 4) - 2
                If MSTDAT.State = 1 Then MSTDAT.Close
                If IsNumeric(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL)) Then
                    FWD_COL = True
                Else
                    FWD_COL = False
                End If
    Case (FRMBOXGRN.Flex.ROW * 4) - 1
        Dim RSITM As ADODB.Recordset: Set RSITM = New ADODB.Recordset
        Dim FICD As String
        Dim DENIWT As Double
        
       '  If RSITM.State = 1 Then RSITM.Close
       '         RSITM.Open "SELECT *  FROM itmmst WHERE COMP='" & compPth & "' AND code='" & FRMBOXGRN.FLEX.TextMatrix(FRMBOXGRN.FLEX.ROW, 11) & "'", CN, adOpenDynamic, adLockOptimistic
       '  If Not RSITM.EOF Then
       '         DENIWT = Val(RSITM!DENI & "")
       '  End If
       
        mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL + 1) = mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL - 1) - Val(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL))
        mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL + 1) = nstr(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL + 1), 10, 3)
        If IsNumeric(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL)) Then
                    FWD_COL = True
                Else
                    FWD_COL = False
                End If
                
            Case (FRMBOXGRN.Flex.ROW * 4)
                If IsNumeric(mfgBndlDet.TextMatrix(mfgBndlDet.ROW, mfgBndlDet.COL)) Then
                    FWD_COL = True
                Else
                    FWD_COL = False
                End If
                
            Case (FRMBOXGRN.Flex.ROW * 4) - 3
                
               FWD_COL = True
            
        End Select

        If FWD_COL = True Then
            If BndlType = 1 Then
                If mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) - 3 Then
                    mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) - 2
                ElseIf mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) - 2 Then
                    mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) - 1
                ElseIf mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) - 1 Then
                    mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4)
                ElseIf mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) Then
                    If mfgBndlDet.ROW <> mfgBndlDet.Rows - 1 Then
                        mfgBndlDet.ROW = mfgBndlDet.ROW + 1
                        mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 4) - 3
                    End If
                End If
            ElseIf BndlType = 2 Then
                If mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 3) - 1 Then
                    mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 3)
                ElseIf mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 3) Then
                    If mfgBndlDet.ROW <> mfgBndlDet.Rows - 1 Then
                        mfgBndlDet.ROW = mfgBndlDet.ROW + 1
                        mfgBndlDet.COL = (FRMBOXGRN.Flex.ROW * 3) - 1
                    End If
                End If
            Else
                If mfgBndlDet.ROW <> mfgBndlDet.Rows - 1 Then
                    mfgBndlDet.ROW = mfgBndlDet.ROW + 1
                    mfgBndlDet.COL = 2
                End If
            End If
            If (mfgBndlDet.RowIsVisible(mfgBndlDet.ROW) = False) And (mfgBndlDet.ROW >= 14) Then mfgBndlDet.TopRow = mfgBndlDet.ROW - 14
        End If
    End If
       
End Sub


Private Sub mfgBndlDet_LeaveCell()
    Dim FLEXROW As Integer
    Dim FLEXCOL As Integer
    Dim I As Long
    Dim J As Long
    I = 0
    J = 0
    
 
 If mfgBndlDet.TextMatrix(mfgBndlDet.ROW, 1) <> Empty Then
    For I = 1 To mfgBndlDet.Rows - 1
    For J = I + 1 To mfgBndlDet.Rows - 1
         If mfgBndlDet.TextMatrix(I, Val(FRMBOXGRN.Flex.ROW * 4) - 3) = mfgBndlDet.TextMatrix(J, Val(FRMBOXGRN.Flex.ROW * 4) - 3) And mfgBndlDet.TextMatrix(J, Val(FRMBOXGRN.Flex.ROW * 4) - 3) <> Empty Then
            MsgBox "Box No. Can Not Similar for This Party", vbOKOnly
            
            Exit Sub
         End If
      Next J
     Next I
  End If
  
  
  
  
    FRMBOXGRN.Flex.TextMatrix(FRMBOXGRN.Flex.ROW, 7) = txtPcs
    If mfgBndlDet.Cols - 1 < (FRMBOXGRN.Flex.ROW * 4) Then
        frmBundelDetails.Hide
        If FRMBOXGRN.Flex.Enabled = True Then FRMBOXGRN.Flex.COL = 9: FRMBOXGRN.Flex.SetFocus
        Exit Sub
    End If
    mfgBndlDet.CellBackColor = vbWhite
    FLEXROW = mfgBndlDet.ROW
    FLEXCOL = mfgBndlDet.COL
    txtPcs.Text = 0
    txtMts.Text = 0
    
    
    For I = 1 To mfgBndlDet.Rows - 1
         If Val(mfgBndlDet.TextMatrix(I, (FRMBOXGRN.Flex.ROW * 4))) <> 0 Then
            txtPcs.Text = Format(Val(txtPcs.Text) + 1, "########")
        End If
        If Val(mfgBndlDet.TextMatrix(I, Val(FRMBOXGRN.Flex.ROW * 4))) <> 0 Then
            txtMts.Text = Format(Val(txtMts.Text) + Val(mfgBndlDet.TextMatrix(I, (FRMBOXGRN.Flex.ROW * 4))), "########.000")
        End If
Next
    mfgBndlDet.ROW = FLEXROW
    mfgBndlDet.COL = FLEXCOL
End Sub

Private Sub mfgBndlDet_LostFocus()
    Call mfgBndlDet_LeaveCell
End Sub


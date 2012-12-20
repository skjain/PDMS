VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUSERRIGHTS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Rights"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   1110
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7231.581
   ScaleMode       =   0  'User
   ScaleWidth      =   20515.38
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkAll 
      Alignment       =   1  'Right Justify
      Caption         =   "All Rights Assigned "
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CheckBox ChkCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "View/Print All"
      Height          =   255
      Index           =   4
      Left            =   8640
      TabIndex        =   7
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CheckBox ChkCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "View All"
      Height          =   255
      Index           =   3
      Left            =   7320
      TabIndex        =   3
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CheckBox ChkCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Delete All"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CheckBox ChkCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Modify All"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CheckBox ChkCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Add All"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin VB.ComboBox CMBGRP 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "frmUSERRIGHTS.frx":0000
      Left            =   2640
      List            =   "frmUSERRIGHTS.frx":0002
      TabIndex        =   6
      Tag             =   "0"
      Text            =   "Transaction Group"
      Top             =   240
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid RPTFLEX 
      Bindings        =   "frmUSERRIGHTS.frx":0004
      Height          =   5295
      Left            =   7080
      TabIndex        =   5
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   128
      FixedCols       =   0
      ForeColorFixed  =   128
      GridLines       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Bindings        =   "frmUSERRIGHTS.frx":0019
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   128
      FixedCols       =   0
      ForeColorFixed  =   128
      ForeColorSel    =   -2147483633
      GridLines       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc ADOTRAN 
      Height          =   330
      Left            =   1200
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc ADORPT 
      Height          =   330
      Left            =   7440
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   $"frmUSERRIGHTS.frx":002F
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   11895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Height          =   495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Select Transaction Group"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Advance Rights Module"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   8760
      TabIndex        =   9
      Top             =   405
      Width           =   2730
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   8760
      Top             =   360
      Width           =   2775
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   5565
      Left            =   120
      Top             =   720
      Width           =   11415
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   15035.79
      X2              =   15035.79
      Y1              =   6387.896
      Y2              =   6870.002
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   12494.53
      X2              =   12494.53
      Y1              =   6387.896
      Y2              =   6870.002
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   9953.269
      X2              =   9953.269
      Y1              =   6387.896
      Y2              =   6870.002
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   7412.009
      X2              =   7412.009
      Y1              =   6387.896
      Y2              =   6870.002
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   5082.521
      X2              =   5082.521
      Y1              =   6387.896
      Y2              =   6870.002
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   120
      Top             =   6360
      Width           =   11415
   End
End
Attribute VB_Name = "frmUSERRIGHTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const strChecked = "þ"
Const strUnChecked = "q"
Dim TRANTYP As String
Dim INIT_FLAG As Boolean

Private Sub ChkAll_Click()
Dim K As Long, i As Long, J As Long
Dim abc
If ChkAll.Value = 1 Then
   abc = strChecked
Else
   abc = strUnChecked
End If

If ChkAll.Value = 1 Then
 For K = 0 To ChkCtrl.COUNT - 1
    ChkCtrl(K).Value = 0
 Next K
End If
'----------------------------------------ALL----------------------------------
  'TRANSACTION
  For i = 2 To 4
    For J = 1 To FLEX.Rows - 1
      FLEX.TextMatrix(J, i) = abc
    Next J
  Next i
  'REPORTS
  For i = 2 To 3
    For J = 1 To RPTFLEX.Rows - 1
      RPTFLEX.TextMatrix(J, i) = abc
    Next J
  Next i
'-----------------------------------------------------------------------------
End Sub

Private Sub ChkCtrl_Click(INDEX As Integer)
Dim i As Long, J As Long
Dim abc
If ChkCtrl(INDEX).Value = 1 Then abc = strChecked: ChkAll.Value = 0 Else abc = strUnChecked

Select Case INDEX
'---------------------------------------ADD ALL---------------------------
Case 0
    For J = 1 To FLEX.Rows - 1    'TRANSACTION
      FLEX.TextMatrix(J, 2) = abc
      FLEX.ROW = J
    Next J
'-------------------------------------MODIFY ALL---------------------------
Case 1
    For J = 1 To FLEX.Rows - 1    'TRANSACTION
      FLEX.ROW = J
      FLEX.TextMatrix(J, 3) = abc
    Next J
'-------------------------------------DELETE ALL---------------------------
Case 2
    For J = 1 To FLEX.Rows - 1    'TRANSACTION
      FLEX.TextMatrix(J, 4) = abc
    Next J
'-------------------------------------VIEW ALL---------------------------
Case 3
    For J = 1 To RPTFLEX.Rows - 1    'REPORTS
      RPTFLEX.TextMatrix(J, 2) = abc
    Next J
'-------------------------------------PRINT ALL---------------------------
Case 4
    For J = 1 To RPTFLEX.Rows - 1    'REPORTS
      RPTFLEX.TextMatrix(J, 2) = abc
      RPTFLEX.TextMatrix(J, 3) = abc
    Next J
End Select

End Sub

Private Sub CMBGRP_Click()
 
 Call SetInfo
   
 Call SetInitial
 ADOTRAN.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
 ADOTRAN.CommandType = adCmdText
 ADOTRAN.RecordSource = "SELECT CODE,NAME,USERRIGHTS.ADDNEW,USERRIGHTS.CHANGE,USERRIGHTS.REMOVE FROM TRANMST " & _
 " LEFT JOIN USERRIGHTS ON TRANMST.CODE=USERRIGHTS.MODULE AND USERCODE='" & Me.Tag & _
 "' AND USERRIGHTS.PDMSNO=3 WHERE TRANMST.GRP='" & CMBGRP.Text & "' AND TYP='M' AND TRANMST.PDMSNO=3 "
 
 ADOTRAN.Refresh
 
 ADORPT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
 ADORPT.CommandType = adCmdText
 ADORPT.RecordSource = "SELECT CODE,NAME,USERRIGHTS.VIEWING,USERRIGHTS.PRINTING FROM TRANMST " & _
 " LEFT JOIN USERRIGHTS ON TRANMST.CODE=USERRIGHTS.MODULE AND USERCODE='" & Me.Tag & _
 "' AND USERRIGHTS.PDMSNO=3 WHERE TRANMST.GRP='" & CMBGRP.Text & "' AND TYP='R' AND TRANMST.PDMSNO=3 "
 ADORPT.Refresh

 Call SETFLEX
 Call FillList
 Call RPTFillList
 TRANTYP = CMBGRP.Text
 
End Sub

Private Sub CMBGRP_KeyPress(KeyAscii As Integer): KeyAscii = 0: End Sub

Private Sub Flex_Click()
With FLEX
    If .COL = 2 Or .COL = 3 Or .COL = 4 Then Call TriggerCheckbox(.ROW, .COL)
End With
End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
        With FLEX
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
                '.CellBackColor = RGB(BRED, BGREEN, BBLUE)
            Else
                 .TextMatrix(iRow, iCol) = strUnChecked
                 '.CellBackColor = vbWhite
            End If
        
        If (.ROW > 13) Then .TopRow = .TopRow + 1
        End With
End Sub

Private Sub SETFLEX()
    With FLEX
        .Cols = 5
        .FixedRows = 1
                
        .TextMatrix(0, 1) = "Transaction"
        .TextMatrix(0, 2) = "Add"
        .TextMatrix(0, 3) = "Modify"
        .TextMatrix(0, 4) = "Delete"
               
        .ColWidth(0) = 0
        .ColWidth(1) = 4200
        .ColWidth(2) = 600
        .ColWidth(3) = 800
        .ColWidth(4) = 800
                
        FLEX.ColAlignment(1) = 0
        FLEX.ColAlignment(2) = flexAlignCenterCenter
        FLEX.ColAlignment(3) = flexAlignCenterCenter
        FLEX.ColAlignment(4) = flexAlignCenterCenter
          
    End With
    
    With RPTFLEX
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        
        .TextMatrix(0, 1) = "Report Format"
        .TextMatrix(0, 2) = "View"
        .TextMatrix(0, 3) = "Print"
               
        .ColWidth(0) = 0
        .ColWidth(1) = 2800
        .ColWidth(2) = 600
        .ColWidth(3) = 600
        
        .ColAlignment(1) = 0
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
          
    End With
End Sub

Private Sub Form_Activate()
    TRANTYP = "A.MASTER"
    Call SetTranGrp
    INIT_FLAG = True
End Sub

Private Sub Form_Load()
  INIT_FLAG = False
  Call ColorComponent(Me)
  With frm_UserCreation
   Me.Caption = "Modifying User : " & Trim(.txtFName) & " " & Trim(.txtMName) & " " & Trim(.txtLName)
  End With
  FLEX.BackColor = RGB(255, 255, 236)
  RPTFLEX.BackColor = RGB(255, 255, 236)
  'CMBGRP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Public Sub FillList()
Screen.MousePointer = vbHourglass
Dim i As Long, J As Long
     
     With FLEX
           ''''Code For option Button
            For J = 2 To 4
                For i = 1 To .Rows - 1
                    .ROW = i
                    .COL = J
                    .CellFontName = "Wingdings"
                    .CellFontSize = 14
                    .CellForeColor = RGB(128, 0, 0)
                    .CellAlignment = flexAlignCenterCenter
                    '.Text = strUnChecked
                Next i
           Next J
           '''''End of Option Button
  End With
  ''''''''''''''''''''''''''''''''''''''''''''
  If FLEX.Rows > 1 Then FLEX.ROW = 1
  FLEX.COL = 1
Screen.MousePointer = vbNormal
End Sub

Public Sub RPTFillList()
Screen.MousePointer = vbHourglass
Dim i As Long, J As Long
     
     With RPTFLEX
           ''''Code For option Button
            For J = 2 To 3
            For i = 1 To .Rows - 1
                    .ROW = i
                    .COL = J
                    .CellFontName = "Wingdings"
                    .CellFontSize = 14
                    .CellForeColor = RGB(128, 0, 0)
                    .CellAlignment = flexAlignCenterCenter
                    '.Text = strUnChecked
            Next i
            Next J
           '''''End of Option Button
  End With
  ''''''''''''''''''''''''''''''''''''''''''''
  
  If RPTFLEX.Rows > 1 Then RPTFLEX.ROW = 1
  RPTFLEX.COL = 1
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call SetInfo
End Sub

Private Sub RPTFLEX_Click()
With RPTFLEX
    If .COL = 2 Or .COL = 3 Then Call TriggerRPTCheckbox(.ROW, .COL)
End With
End Sub

Private Sub TriggerRPTCheckbox(iRow As Integer, iCol As Integer)
        With RPTFLEX
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
                '.CellBackColor = RGB(214, 218, 254)
            Else
                 .TextMatrix(iRow, iCol) = strUnChecked
                 '.CellBackColor = vbWhite
            End If
        
        If (.ROW > 13) Then .TopRow = .TopRow + 1
        End With
End Sub

Private Sub SetInitial()
  FLEX.FixedCols = 0
  FLEX.FixedRows = 1
  FLEX.Cols = 2
  FLEX.Rows = 2
  
  RPTFLEX.FixedCols = 0
  RPTFLEX.FixedRows = 1
  RPTFLEX.Cols = 2
  RPTFLEX.Rows = 2
End Sub

Private Sub SetTranGrp()
Dim GRPRS As ADODB.Recordset
Set GRPRS = New ADODB.Recordset
If GRPRS.State = 1 Then GRPRS.Close

GRPRS.Open "SELECT DISTINCT GRP FROM TRANMST WHERE PDMSNO=3 ", CN, adOpenDynamic, adLockOptimistic
Do While Not GRPRS.EOF
 CMBGRP.AddItem Trim(GRPRS!Grp)
GRPRS.MoveNext
Loop
If CMBGRP.ListCount > 1 Then CMBGRP.ListIndex = 0
End Sub

Private Sub SetInfo()
'FORM NOT LOADED

If Not INIT_FLAG Then Exit Sub
'---------------
Dim i As Long

  CN.Execute "DELETE FROM USERRIGHTS WHERE COMP='" & compPth & "' AND USERCODE='" & Me.Tag & _
             "' AND GRP='" & TRANTYP & "' AND PDMSNO=3 ", i
  
  With FLEX
  For i = 1 To FLEX.Rows - 1
  
  CN.Execute "INSERT INTO USERRIGHTS (COMP,MODULE,MODNAME,USERCODE,ADDNEW,CHANGE,REMOVE,VIEWING," & _
  "PRINTING,CATA,GRP,PDMSNO) VALUES('" & compPth & "','" & .TextMatrix(i, 0) & "','" & Trim(.TextMatrix(i, 1)) & "','" & Me.Tag & _
  "','" & IIf((.TextMatrix(i, 3) = strChecked) Or (.TextMatrix(i, 4) = strChecked), strChecked, .TextMatrix(i, 2)) & "','" & .TextMatrix(i, 3) & _
  "','" & .TextMatrix(i, 4) & "','" & strUnChecked & "','" & strUnChecked & "','M','" & TRANTYP & "',3)"
    
  Next i
  End With
  
  With RPTFLEX
  For i = 1 To RPTFLEX.Rows - 1
  
  CN.Execute "INSERT INTO USERRIGHTS (COMP,MODULE,MODNAME,USERCODE,ADDNEW,CHANGE,REMOVE,VIEWING,PRINTING," & _
  "CATA,GRP,PDMSNO) VALUES('" & compPth & "','" & .TextMatrix(i, 0) & "','" & Trim(.TextMatrix(i, 1)) & "','" & Me.Tag & _
  "','" & strUnChecked & "','" & strUnChecked & "'" & _
  ",'" & strUnChecked & "','" & IIf(.TextMatrix(i, 3) = strChecked, strChecked, .TextMatrix(i, 2)) & "','" & .TextMatrix(i, 3) & "','R','" & TRANTYP & "',3)"
    
  Next i
  End With

'clear check box : all
Dim K As Long
For K = 0 To ChkCtrl.COUNT - 1
 ChkCtrl(K).Value = 0
Next K
ChkAll.Value = 0
'------------------------
End Sub

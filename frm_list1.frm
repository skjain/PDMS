VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frm_list1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Search"
   ClientHeight    =   3855
   ClientLeft      =   495
   ClientTop       =   1845
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8010
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   5400
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
   Begin VB.Frame framDetails 
      Caption         =   "Details"
      Height          =   3360
      Left            =   4455
      TabIndex        =   5
      Top             =   420
      Width           =   3495
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   765
      End
      Begin VB.Label txtAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Address of the Party"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.Label lblAcGrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/c Group:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1344
         Width           =   945
      End
      Begin VB.Label lblCPCD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Cpcd:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1667
         Width           =   1095
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1990
         Width           =   480
      End
      Begin VB.Label lblAgent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2358
         Width           =   630
      End
      Begin VB.Label lblCurb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Balance:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2700
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label txtCURB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1740
         TabIndex        =   15
         Top             =   2700
         Width           =   1635
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
         TabIndex        =   14
         Top             =   2358
         Width           =   2115
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
         TabIndex        =   13
         Top             =   1990
         Width           =   2145
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
         TabIndex        =   12
         Top             =   1667
         Width           =   1020
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
         TabIndex        =   11
         Top             =   1344
         Width           =   1320
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
         TabIndex        =   10
         Top             =   240
         Width           =   2415
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
         TabIndex        =   9
         Top             =   720
         Width           =   3255
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
         TabIndex        =   8
         Top             =   1080
         Width           =   3255
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
         TabIndex        =   7
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax Catagoery"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   1800
      End
   End
   Begin VB.TextBox txtName 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7830
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
      Top             =   3525
      Width           =   975
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
      Top             =   3525
      Visible         =   0   'False
      Width           =   975
   End
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
      TabIndex        =   4
      Top             =   3525
      Width           =   975
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frm_list1.frx":0000
      DataSource      =   "Adodc1"
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5265
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      BoundColumn     =   ""
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
End
Attribute VB_Name = "frm_list1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public open_flag As Boolean
Public form_Para As String
Public SQL As String

Private Sub DataList1_DblClick()
    Call CMDOK_Click
End Sub

Private Sub Form_Load()
  On Error GoTo errLoad
    CANCEL_CLICK = False
    Call CenterChild(frm_Main, Me)

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
    
    TXTNAME = sTxt
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

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
    CANCEL_CLICK = True
End Sub

Private Sub CMDOK_Click()
    If Trim(DataList1.Text) = Empty Then
      Exit Sub
      If TXTNAME.Enabled Then TXTNAME.SetFocus
    End If
    
    M_DESC = Trim(DataList1.Text)
    
    Call SetCode
    TXTNAME = Empty
    Unload frm_list1
End Sub

Private Sub Form_Activate()

Call ColorComponent(Me)
On Error GoTo ERRACTIVATE

    If Trim(sTxt) = Empty Then
       Dim QUERY As String
       QUERY = QRY & AddSQL(QRY)
       Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
       Adodc1.CommandType = adCmdText
       Adodc1.RecordSource = QUERY
       Adodc1.Refresh
       Call SetDataListField
    End If
    
    If IsImportActive Then
       TXTNAME = Empty
       TXTNAME.Text = sTxt
    End If

    If IsImportActive Then
       If lstName.ListIndex > 0 Then
          cmdOk.Default = True
       ElseIf cmdAddNew.Visible And TXTNAME <> Empty Then
          cmdAddNew.Default = True
       Else
          cmdOk.Default = True
       End If
     End If
         
     CANCEL_VISIBLE = True
       
        
     If open_flag = False And Me.Visible = False Then
        Me.Show
     End If
        
     key_PressNew = False
        
     SetFormSize
        
     If TXTNAME.Enabled Then TXTNAME.SetFocus
        
     Exit Sub
        
ERRACTIVATE:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Msg Empty
    lstCaption = Empty
    sTxt = Empty
End Sub

Private Sub DataList1_Click()
    dispDetails (DataList1.Text)
End Sub

Private Sub DATALIST1_GotFocus()
    DataList1.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Select Relevant Value From List"
End Sub

Private Sub TXTNAME_Change()
Dim TEMPRS As New ADODB.Recordset, SQL As String
SQL = QRY
       
If VBA.Strings.InStr(UCase(SQL), "WHERE") = 0 Then
   SQL = SQL & " WHERE "
Else
   SQL = SQL & " AND "
End If
      
   If DataList1.Text = "" Then
      SetDataListField
   End If
   
   SQL = SQL & " " & UCase(DataList1.ListField) & " LIKE ('" & TXTNAME.Text & "%') "
   SQL = SQL & " ORDER BY " & UCase(DataList1.ListField) & " "

   Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
   Adodc1.CommandType = adCmdText
   Adodc1.RecordSource = SQL
   Adodc1.Refresh
   
End Sub


Private Sub txtName_GotFocus()
    TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    TXTNAME.SelStart = 0
    TXTNAME.SelLength = Len(TXTNAME.Text)
End Sub

Private Sub TXTNAME_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then DataList1.SetFocus
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
        TEMPRS.Open "SELECT * FROM REFMST WHERE CATA='Y' AND NAME='" & DataList1.Text & "'", CN
        If TEMPRS.EOF = False Then
            lblAdd1 = Trim(TEMPRS!ADL1 & "")
            lblAdd2 = Trim(TEMPRS!ADL2 & "")
            lblAdd3 = TEMPRS!Area & ""
        End If
        TEMPRS.Close
    ElseIf InStr(1, UCase(QRY), "ACCMST") <> 0 Then
        
        TEMPRS.Open "SELECT * FROM LIST where NAME='" & DataList1.Text & "'", CN, adOpenDynamic, adLockOptimistic
        If Not TEMPRS.EOF Then
        
            txtAcGrp.Caption = TEMPRS!ACGRP & ""
            TXTADDRESS.Caption = TEMPRS![adro] & ""
            If Not IsNull(TEMPRS![CPCD]) Then txtCPCD.Caption = TEMPRS![CPCD] Else txtCPCD.Caption = "N/A"
            If Not IsNull(TEMPRS![Area]) Then TXTAREA.Caption = TEMPRS![Area] Else TXTAREA.Caption = "N/A"
            If Not IsNull(TEMPRS![BROKER]) Then txtBroker.Caption = TEMPRS![BROKER] Else txtBroker.Caption = "N/A"
            txtCURB.Caption = ".00"
            M_RTTX.Caption = TEMPRS!TTYP & ""
        
        End If
        TEMPRS.Close
        
        lblAdd1.Caption = Mid(TXTADDRESS.Caption, 1, 130)
        lblAdd2.Caption = Mid(TXTADDRESS.Caption, 31, 40)
        lblAdd3.Caption = Mid(TXTADDRESS.Caption, 71, 40)
        lblAdd2.Visible = False
        lblAdd3.Visible = False
    End If
    
    Exit Sub
    
errDispDetails:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub SetFormSize()
    If InStr(1, UCase(QRY), "ACCMST") <> 0 Then
        lblAcGrp.Visible = True
        txtAcGrp.Visible = True
        lblCPCD.Visible = True
        txtCPCD.Visible = True
        lblArea.Visible = True
        TXTAREA.Visible = True
        LBLAGENT.Visible = True
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
        TXTAREA.Visible = False
        LBLAGENT.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False
        
        lblAddress.Caption = "BALANCE :"
        TXTADDRESS.Caption = "  0.00"
        
    ElseIf InStr(1, UCase(QRY), "IGMMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        TXTAREA.Visible = False
        LBLAGENT.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False
        framDetails.Visible = False
        Me.WIDTH = DataList1.WIDTH + 200
    ElseIf InStr(1, QRY, "WHERE CATA='Y'", vbTextCompare) > 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        TXTAREA.Visible = False
        LBLAGENT.Visible = False
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
        TXTAREA.Visible = False
        LBLAGENT.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        TXTNAME.WIDTH = DataList1.WIDTH
        Me.WIDTH = DataList1.WIDTH + 350

    ElseIf InStr(1, UCase(QRY), "GRPMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        TXTAREA.Visible = False
        LBLAGENT.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        TXTNAME.WIDTH = DataList1.WIDTH
        Me.WIDTH = DataList1.WIDTH + 350


    ElseIf InStr(1, UCase(QRY), "HEDMST") <> 0 Then
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        TXTAREA.Visible = False
        LBLAGENT.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        TXTNAME.WIDTH = DataList1.WIDTH
        Me.WIDTH = DataList1.WIDTH + 350
    
    Else
        lblAcGrp.Visible = False
        txtAcGrp.Visible = False
        lblCPCD.Visible = False
        txtCPCD.Visible = False
        lblArea.Visible = False
        TXTAREA.Visible = False
        LBLAGENT.Visible = False
        txtBroker.Visible = False
        
        lblCurb.Visible = False
        txtCURB.Visible = False

        framDetails.Visible = False
        TXTNAME.WIDTH = DataList1.WIDTH
        Me.WIDTH = DataList1.WIDTH + 350
    End If
End Sub

Private Function getRate(sItemCode As String, sCompName As String) As Double
Dim rsTemp As Recordset

    Set rsTemp = New Recordset
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "Select AVG(RATE) AS AVGRATE FROM SPTRAN WHERE VTYP='PUR' AND ICOD='" & sItemCode & "' AND LEFT(VBNO,1)<>'*' AND COMP='" & compPth & "'", CN
    
    If rsTemp.EOF = False And IsNull(rsTemp!AVGRATE) = False Then getRate = rsTemp!AVGRATE
    
    rsTemp.Close
    
End Function

Private Sub txtName_LostFocus()
  TXTNAME.BackColor = vbWhite
End Sub

Private Sub SetDataListField()
Dim SQL As String

SQL = QRY

If InStr(1, SQL, "TXULOT") <> 0 Then
   DataList1.ListField = "LTNO"
ElseIf InStr(1, SQL, "ORDMAN") <> 0 Then
   DataList1.ListField = "ORDN"
ElseIf InStr(1, SQL, "GRNTRAN") <> 0 Then
   DataList1.ListField = "VBNO"
ElseIf InStr(1, SQL, "EGPMAN") <> 0 Then
   DataList1.ListField = "CHLN"
ElseIf InStr(1, SQL, "TTYP") <> 0 Then
   DataList1.ListField = "NAME"
ElseIf InStr(1, SQL, "MRGMST") <> 0 Then
   DataList1.ListField = "MRGN"
ElseIf InStr(1, SQL, "TRNMAN") <> 0 Then
   DataList1.ListField = "VTYP"
ElseIf InStr(1, SQL, "SUBGRDMST") <> 0 Then
   DataList1.ListField = "NAME"
ElseIf InStr(1, UCase(SQL), "GRDMST") <> 0 Then
   DataList1.ListField = "GRAD"
ElseIf InStr(1, SQL, "EXTRA4") <> 0 Then
   DataList1.ListField = "EXTRA4"
ElseIf InStr(1, SQL, "LOCATION") <> 0 Then
   DataList1.ListField = "LOCNAME"
ElseIf InStr(1, SQL, "ORDTRN") <> 0 Or InStr(1, SQL, "ordtrn") <> 0 Then
   DataList1.ListField = "DONO" 'AND VTYP='DOS'"
ElseIf InStr(1, SQL, "SERIALMASTER") <> 0 Then
   DataList1.ListField = "NAME"
ElseIf InStr(1, SQL, "PROD_CONNING") <> 0 Then
   DataList1.ListField = "JOBCARDNO"
ElseIf InStr(1, SQL, "REPCNF") <> 0 Then
   DataList1.ListField = "RPNM"
ElseIf InStr(1, SQL, "CHRGMST") <> 0 Then
   DataList1.ListField = "NAME"
Else
   DataList1.ListField = "NAME"
End If
End Sub

Private Sub SetCode()

Dim SQL As String
TXTNAME = DataList1.Text
SQL = QRY

If VBA.Strings.InStr(UCase(SQL), "WHERE") = 0 Then  'NOT WHERE EXIST
   SQL = SQL & " WHERE "
Else
   SQL = SQL & " AND "
End If

If InStr(1, SQL, "TXULOT") <> 0 Then
   SQL = SQL & "LTNO = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "GRNTRAN") <> 0 Then
   SQL = SQL & "VBNO = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "ORDMAN") <> 0 Then
   SQL = SQL & "ORDN = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "EGPMAN") <> 0 Then
   SQL = SQL & "CHLN = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "TTYP") <> 0 Then
   SQL = SQL & "TTYP = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "MRGMST") <> 0 Then
   SQL = SQL & "MRGN = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "TRNMAN") <> 0 Then
   SQL = SQL & "VTYP = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "SUBGRDMST") <> 0 Then
   SQL = SQL & "NAME = '" & TXTNAME & "' "
ElseIf InStr(1, UCase(SQL), "GRDMST") <> 0 Then
   SQL = SQL & "GRAD = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "EXTRA4") <> 0 Then
   SQL = SQL & "EXTRA4 = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "LOCATION") <> 0 Then
   SQL = SQL & "LOCNAME = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "ORDTRN") <> 0 Or InStr(1, SQL, "ordtrn") <> 0 Then
   SQL = SQL & "DONO = '" & TXTNAME & "' AND VTYP='DOS'"
ElseIf InStr(1, SQL, "SERIALMASTER") <> 0 Then
   SQL = SQL & "NAME = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "PROD_CONNING") <> 0 Then
   SQL = SQL & "JOBCARDNO = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "REPCNF") <> 0 Then
   SQL = SQL & "RPNM = '" & TXTNAME & "' "
ElseIf InStr(1, SQL, "CHRGMST") <> 0 Then
   SQL = SQL & "NAME = '" & TXTNAME & "' "
Else
   SQL = SQL & "NAME = '" & TXTNAME & "' "
End If

Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not TEMPRS.EOF Then
  Key = Trim(TEMPRS.Fields(0).Value)
End If

End Sub


Private Function AddSQL(SQL As String) As String
AddSQL = ""

If InStr(1, SQL, "TXULOT") <> 0 Then
   AddSQL = " ORDER BY LTNO"
ElseIf InStr(1, SQL, "ORDMAN") <> 0 Then
   AddSQL = " ORDER BY ORDN"
ElseIf InStr(1, SQL, "GRNTRAN") <> 0 Then
   AddSQL = " ORDER BY VBNO"
ElseIf InStr(1, SQL, "EGPMAN") <> 0 Then
   AddSQL = " ORDER BY CHLN"
ElseIf InStr(1, SQL, "TTYP") <> 0 Then
   AddSQL = " ORDER BY TTYP"
ElseIf InStr(1, SQL, "MRGMST") <> 0 Then
   AddSQL = " ORDER BY MRGN"
ElseIf InStr(1, SQL, "TRNMAN") <> 0 Then
   AddSQL = " ORDER BY VTYP"
ElseIf InStr(1, SQL, "SUBGRDMST") <> 0 Then
   AddSQL = " ORDER BY NAME"
ElseIf InStr(1, UCase(SQL), "GRDMST") <> 0 Then
   AddSQL = " ORDER BY GRAD"
ElseIf InStr(1, SQL, "EXTRA4") <> 0 Then
   AddSQL = " ORDER BY EXTRA4"
ElseIf InStr(1, SQL, "LOCATION") <> 0 Then
   AddSQL = " ORDER BY LOCNAME"
ElseIf InStr(1, SQL, "ORDTRN") <> 0 Or InStr(1, SQL, "ordtrn") <> 0 Then
   AddSQL = " ORDER BY DONO" 'AND VTYP='DOS'"
ElseIf InStr(1, SQL, "SERIALMASTER") <> 0 Then
   AddSQL = " ORDER BY NAME"
ElseIf InStr(1, SQL, "PROD_CONNING") <> 0 Then
   AddSQL = " ORDER BY JOBCARDNO"
ElseIf InStr(1, SQL, "REPCNF") <> 0 Then
   AddSQL = " ORDER BY RPNM"
ElseIf InStr(1, SQL, "CHRGMST") <> 0 Then
   AddSQL = " ORDER BY NAME"
   
Else
   AddSQL = " ORDER BY NAME"
End If

End Function

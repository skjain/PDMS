VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHB~1.OCX"
Begin VB.Form frmItemOpnMerge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge Wise Item Opening"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7575
      Begin VB.TextBox TXTMRGN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   295
         Left            =   3840
         TabIndex        =   7
         Top             =   1560
         Width           =   3450
      End
      Begin VB.TextBox TXTNAME 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   5850
      End
      Begin VB.TextBox TXTPCS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   295
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox TXTQNTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   295
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1200
         Width           =   1290
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   295
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1200
         Width           =   1410
      End
      Begin VB.TextBox TXTCOPS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   295
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1560
         Width           =   1050
      End
      Begin VB.TextBox TXTGDN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   5835
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merge No.:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Item  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Box/Pcs. :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   18
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Opening Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cops   :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Godown :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   7575
      Begin WelchButton.lvButtons_H CMDDEL 
         Height          =   495
         Left            =   4200
         TabIndex        =   9
         Top             =   240
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
         Image           =   "frmItemOpnMerge.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDSAVE 
         Height          =   495
         Left            =   1080
         TabIndex        =   8
         Top             =   240
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
         Image           =   "frmItemOpnMerge.frx":059A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5760
         TabIndex        =   10
         Top             =   240
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
         Image           =   "frmItemOpnMerge.frx":0B34
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         Image           =   "frmItemOpnMerge.frx":10CE
         cBack           =   -2147483633
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7011
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Bindings        =   "frmItemOpnMerge.frx":1520
         Height          =   3615
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSAdodcLib.Adodc ADOOPN 
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "DIVISION WISE ITEM OPENING RECORD SET"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmItemOpnMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_DVCD As String, M_VBNO As String, STR As String
Dim spara As String, SQL As String, icode As String, GDNCOD As String
Dim OPNDATE As Date
Dim i As Integer

Private Sub cmdCancel_Click()
    TXTNAME.Text = Empty
    TXTQNTY.Text = Empty
    TXTQNTY.Tag = Empty
    txtValue.Text = Empty
    TXTPCS.Text = Empty
    txtCops.Text = Empty
    TXTMRGN = Empty
    TXTNAME.SetFocus
End Sub

Private Sub FLEX_DblClick()
If FLEX.Rows <= 1 Or FLEX.ROW < 1 Then Exit Sub
If Trim(FLEX.TextMatrix(1, 0)) = Empty Then Exit Sub

    TXTNAME.Text = FLEX.TextMatrix(FLEX.ROW, 0)
    icode = FLEX.TextMatrix(FLEX.ROW, 4)
    TXTGDN.Text = FLEX.TextMatrix(FLEX.ROW, 6)
    GDNCOD = Trim(FLEX.TextMatrix(FLEX.ROW, 7))
    TXTMRGN = LTrim(RTrim(FLEX.TextMatrix(FLEX.ROW, 8)))
    Call txtName_LostFocus
    Call EditHelp
End Sub

Private Sub FLEX_KeyDown(KeyCode As Integer, Shift As Integer)
If FLEX.Rows <= 1 Or Trim(FLEX.TextMatrix(1, 0)) = Empty Or FLEX.ROW < 1 Then Exit Sub
    TXTNAME.Text = FLEX.TextMatrix(FLEX.ROW, 0)
    icode = FLEX.TextMatrix(FLEX.ROW, 4)
    Call txtName_LostFocus
    TXTNAME.SetFocus
End Sub

Private Sub Form_Activate()
   Call ColorComponent(Me)
   Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case UCase(ActiveControl.NAME)
Case "TXTPCS"
    If TXTPCS = Empty Then Exit Sub
Case "TXTQNTY"
    If TXTQNTY = Empty Then Exit Sub
Case "TXTVALUE"
    If txtValue = Empty Then Exit Sub
Case "TXTMRGN"
    If TXTMRGN = Empty Then Exit Sub
End Select
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
   Call CenterChild(frm_Main, Me)
   OPNDATE = FSDT - 1
   
   ADOOPN.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
   ADOOPN.CommandType = adCmdText
   ADOOPN.RecordSource = "SELECT ITMMST.NAME AS ITEM,PCES,QNTY,AMNT,STORETRAN.ICOD,COPS,LOCMST.NAME AS GODOWN,STORETRAN.GDNCOD,STORETRAN.LTNO FROM STORETRAN " & _
   " INNER JOIN LOCMST ON STORETRAN.GDNCOD=LOCMST.CODE " & _
   " INNER JOIN ITMMST ON STORETRAN.ICOD=ITMMST.CODE " & _
   " WHERE STORETRAN.COMP='" & compPth & "' AND STORETRAN.UNIT='" & UNCD & _
   "' AND STORETRAN.DVCD='000001' AND STORETRAN.VTYP='IVR' AND STORETRAN.RECSTAT<>'D' AND " & _
   "STORETRAN.DATE <= '" & Format(OPNDATE, "MM/DD/YYYY") & "' ORDER BY STORETRAN.VBNO DESC"
      ADOOPN.Refresh
   
   Call SETFLEX
End Sub

Private Sub FLEX_EnterCell()
   FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub Flex_LeaveCell()
   FLEX.CellBackColor = vbWhite
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("0012", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    If Val(txtValue) <= 0 Then
       MsgBox "Value Must be Greater then 0"
       txtValue.Enabled = True
       txtValue.SetFocus
       Exit Sub
    End If
       
    'CHECK DATA
    If TXTNAME = Empty Or TXTGDN = Empty Or (Val(TXTPCS) = 0 And Val(TXTQNTY) = 0) Then
       If TXTNAME = Empty Then TXTNAME.SetFocus
       If TXTGDN = Empty Then TXTGDN.SetFocus
       If TXTMRGN = Empty Then TXTMRGN.SetFocus
       If (Val(TXTPCS) = 0 And Val(TXTQNTY) = 0) Then TXTQNTY.SetFocus
       Exit Sub
    End If
    If TXTMRGN = Empty Then
    TXTMRGN.SetFocus
    Exit Sub
    End If
    '-------------------
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM ITMMST WHERE NAME ='" & TXTNAME & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       icode = Trim(RS!CODE & "")
    Else
       MsgBox "CHECK ITEM"
       Exit Sub
    End If
    
    '--------------------
    'SET OPENING DATE
     OPNDATE = FSDT - 1
    '--------------------
       
    'FIND RATE USING VALUE AND QNTY
       Dim RATE As Double
       RATE = 0
    
       If Val(txtValue) > 0 And Val(TXTQNTY) > 0 Then
          RATE = Val(txtValue) / Val(TXTQNTY)
       End If
    '-------------------
    
    If spara = "NEW" Then
       M_VBNO = GENOPNCODE
    End If
    
    If spara = "EDIT" Then
      If STOPEDIT Then Exit Sub
    End If
    
    CN.BeginTrans
    
        If spara = "NEW" Then
                        
            SQL = "INSERT INTO STORETRAN(COMP,UNIT,VTYP,DBCD,VBNO,SRCH,DVCD,DATE,PCOD,ICOD,PCES,QNTY,AMNT,QORP," & _
                  "OPER,SYSR,[USER],RATE,GWGT,RECSTAT,COPS,GDNCOD,LTNO,MRGN) VALUES ('" & compPth & "','" & UNCD & "','IVR','XXXXXX','" & M_VBNO & _
                  "',1,'000001','" & Format(OPNDATE, "mm/dd/yyyy") & "','XXXXXX','" & icode & _
                  "'," & Val(TXTPCS) & "," & Val(TXTQNTY) & "," & Val(txtValue) & ",'Q','+','T','" & cUName & _
                  "'," & RATE & "," & Val(TXTQNTY) & ",'A','" & Val(txtCops) & "','" & GetCode("LOCMST", TXTGDN, "NAME", "CODE") & "','" & Trim(TXTMRGN) & "','" & Trim(TXTMRGN) & "')"
            CN.Execute SQL
            
            
            SQL = "INSERT INTO GRNTRAN([COMP],[UNIT],[VTYP],[VBNO],[DBCD],[SRCH],DATE,[ICOD],[RATE],[GRN_QNTY],[NETRATE],[BAL_QNTY],[MRGN])"
            SQL = SQL & " VALUES('" & compPth & "','" & UNCD & "','IVR','" & M_VBNO & "','XXXXXX','1', " & _
            "'" & Format(OPNDATE, "yyyy-MM-dd hh:mm:ss") & "','" & icode & "','" & RATE & "','" & Val(TXTQNTY) & _
            "','" & Val(RATE) & "','" & Val(TXTQNTY) & "','" & Trim(TXTMRGN) & "')"
            
            CN.Execute SQL
            
            Call DAILYSTATUS("OPN", icode, "XXXXXX", Val(TXTQNTY), M_VBNO, RATE, cUName, "I", Now, Now)
            
        Else
            
            SQL = "UPDATE STORETRAN SET RATE=" & RATE & ",GWGT=" & Val(TXTQNTY) & ", QNTY=" & Val(TXTQNTY) & _
            ",PCES=" & Val(TXTPCS) & ",AMNT='" & Val(txtValue) & "',COPS = '" & Val(txtCops) & "',LTNO = '" & Trim(TXTMRGN) & "',MRGN = '" & Trim(TXTMRGN) & "' WHERE COMP='" & compPth & "' AND " & _
            "UNIT='" & UNCD & "' AND DVCD='000001' AND VTYP='IVR' AND VBNO='" & M_VBNO & "' AND ICOD='" & icode & "' AND GDNCOD='" & GDNCOD & "' AND RECSTAT='A'"
            
            CN.Execute SQL
            
            SQL = "UPDATE GRNTRAN SET RATE='" & RATE & "',GRN_QNTY='" & Val(TXTQNTY) & "',MRGN = '" & Trim(TXTMRGN) & "',NETRATE='" & RATE & _
            "',BAL_QNTY='" & Val(TXTQNTY) & "' WHERE COMP='" & compPth & _
            "' AND UNIT = '" & UNCD & "' AND VTYP='IVR' AND VBNO = '" & M_VBNO & "' AND ICOD='" & icode & "'"
            
            CN.Execute SQL
            
            Call DAILYSTATUS("OPN", icode, "XXXXXX", Val(TXTQNTY), M_VBNO, RATE, cUName, "U", Now, Now)
    End If
    
    If TXTMRGN <> Empty Then
    Dim MSTDAT As New ADODB.Recordset
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND MRGN='" & Trim(TXTMRGN) & "'", CN, adOpenDynamic, adLockOptimistic
    If MSTDAT.EOF Then
      MSTDAT.AddNew
      MSTDAT!COMP = compPth
      MSTDAT!unit = UNCD
      MSTDAT!MRGN = Trim(TXTMRGN)
      MSTDAT!PCOD = ""
      MSTDAT!ICOD = icode
      MSTDAT.Update
    End If
    End If
        
    CN.CommitTrans
    ADOOPN.Refresh
        
    Call cmdCancel_Click
    Exit Sub

LAST:
    MsgBox ERR.Description
    Resume
    CN.RollbackTrans
End Sub

Private Sub cmddel_Click()
On Error GoTo LAST
    
    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("000035", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    If spara <> "EDIT" Then Exit Sub
           
    'CHECK DATA
    If TXTNAME = Empty Then
       If TXTNAME.Enabled Then TXTNAME.SetFocus
       Exit Sub
    End If
    '-------------------
    
     If STOPEDIT Then Exit Sub
    
    STR = MsgBox("ARE YOU SURE YOU WANT TO DELETE THIS ITEM OPENING DETAIL ?", vbYesNo + vbQuestion, "Remove This Opening Detail ?")
    
    CN.BeginTrans
    
    If STR = vbYes Then
       CN.Execute "DELETE STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                  "' AND DVCD='000001' AND VTYP='IVR' AND VBNO = '" & M_VBNO & "' AND RECSTAT<>'D'"
                  
       SQL = "DELETE GRNTRAN WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
             "' AND VTYP='IVR' AND VBNO = '" & M_VBNO & "' "
            
       CN.Execute SQL
                  
    End If

       'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','ITM','XXXXXXXXXXXXX','" & txtName & "',NULL,'" & ICODE & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
       CN.CommitTrans
    

    Call cmdCancel_Click
    
    ADOOPN.Refresh
    Exit Sub

LAST:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Public Function SETFLEX()
    With FLEX
                
        .ColWidth(0) = 2800
        .TextMatrix(0, 0) = "Item"
        .ColAlignment(0) = 1
    
        .ColWidth(1) = 800
        .ColAlignment(1) = 0
        .TextMatrix(0, 1) = "Pieces"
  
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = "Quantity"
        .ColAlignment(2) = 0
  
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "Amount"
        .ColAlignment(3) = 0
                        
        .ColWidth(4) = 0
        .TextMatrix(0, 4) = "ICOD"
        
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = "Cops"
        
        .ColWidth(6) = 2800
        .TextMatrix(0, 6) = "Godown"
        .ColAlignment(6) = 1
        
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = "Godown Code"
        .ColAlignment(7) = 1
        
        .ColWidth(8) = 900
        .TextMatrix(0, 8) = "Merge No."
        .ColAlignment(8) = 1

        
    End With
End Function

Private Sub Label7_Click()

End Sub

Private Sub txtCops_GotFocus()
txtCops.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtcops_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, txtCops, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCops_LostFocus()
txtCops.BackColor = vbWhite
End Sub

Private Sub TXTGDN_GotFocus()
TXTGDN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTGDN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTGDN = Empty) Then
        M_DESC = Empty
        Key = Empty
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        TXTGDN.Text = SearchList1("select TOP 20 code,name from LOCMST ", 0, "", "List Of GODOWN")
        TXTGDN.Tag = Key
        GDNCOD = Key
    End If
    If KeyCode = vbKeyDelete Then TXTGDN = Empty
End Sub

Private Sub txtgdn_LostFocus()
TXTGDN.BackColor = vbWhite
End Sub

Private Sub TXTMRGN_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTMRGN = Empty
    ElseIf KeyCode = vbKeyF2 Then
        M_DESC = Empty
        Key = Empty
        NEW_VISIBLE = False
        TXTMRGN = SearchList1("Select DISTINCT MRGN,MRGN  From MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD = '" & GetCode("ITMMST", TXTNAME, "NAME", "CODE") & "'", 0, Empty, "Select MERGE FROM MASTER")
        TXTMRGN.Tag = Key
    End If
  
  If KeyCode = vbKeyDelete Then
     TXTMRGN = Empty
  End If
  
  If KeyCode = vbKeyReturn Or KeyCode = 20 Then
    cmdSave.SetFocus
  End If
  
  Me.KeyPreview = True
End Sub

Private Sub txtName_GotFocus()
    TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Press <F2> To Get List Of Item"
End Sub

Private Sub TXTNAME_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        M_DESC = Empty
        Key = Empty
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        TXTNAME.Text = SearchITEMLIST("SELECT TOP 20 CODE,NAME FROM VWITEM ", 0, "", "List Of Items")
        icode = Key
    End If
    
    '----------------------------------------------------------------------
    'SPECIFICATION ACCORDING ITEM GROUP
    Dim RS As New ADODB.Recordset
    Dim SPECI As String
    Dim MRGN As String
    Dim igcd As String
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT *  FROM ITMMST WHERE CODE = '" & icode & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
    igcd = RS!igcd
    End If
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM IGMMST WHERE CODE = '" & igcd & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
    SPECI = RS!SPECIFICATION
    MRGN = RS!MERGE
    End If
    
    
    If MRGN = "Y" Then
       TXTMRGN.Enabled = True
    Else
       TXTMRGN.Enabled = False
    End If
       
    If Val(SPECI) = 0 Then
       TXTPCS.Enabled = True
       TXTQNTY.Enabled = True
       txtCops.Enabled = False
    ElseIf Val(SPECI) = 1 Then
       TXTQNTY.Enabled = True
       TXTPCS.Enabled = False
       txtCops.Enabled = False
    ElseIf Val(SPECI) = 2 Then
       TXTPCS.Enabled = True
       txtCops.Enabled = True
       TXTQNTY.Enabled = True
    ElseIf Val(SPECI) = 3 Then
       txtCops.Enabled = True
       TXTQNTY.Enabled = True
       TXTPCS.Enabled = False
    End If
       
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TXTNAME.Text = "" Then
        M_DESC = Empty
        Key = Empty
        TXTNAME.Text = SearchITEMLIST("select TOP 20 code,name from VWITEM", 0, "", "List Of Items")
        icode = Key
    End If
End Sub

Private Sub txtName_LostFocus()
    TXTNAME.BackColor = vbWhite
    If TXTNAME.Text = "" Then Exit Sub
    SQL = "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND STORETRAN.UNIT='" & UNCD & _
          "' AND STORETRAN.DVCD='000001' AND VTYP='IVR' AND ICOD='" & icode & _
          "' AND RECSTAT<>'D' AND STORETRAN.DATE <= '" & Format(OPNDATE, "MM/DD/YYYY") & "' AND GDNCOD='" & GDNCOD & _
          "' AND LTNO = '" & Trim(TXTMRGN) & "'"
          
    With RS
        If .State = adStateOpen Then .Close
        .Open SQL, CN, adOpenDynamic, adLockOptimistic
        If .EOF = True Then
            spara = "NEW"
            Exit Sub
        End If
    
        TXTPCS = Trim(nstr(!PCES, 9, 0))
        TXTQNTY = !QNTY
        txtCops = !COPS & ""
        TXTQNTY.Tag = !QNTY
        txtValue = !AMNT
        M_VBNO = !VBNO
        TXTMRGN = Trim(!ltno)
        spara = "EDIT"
    End With
End Sub

Private Sub TXTPCS_GotFocus()
    TXTPCS.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Opening Pcs."
End Sub

Private Sub TXTPCS_KeyPress(KeyAscii As Integer)
     If CheckNumericKey(KeyAscii, TXTPCS, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTPCS_LostFocus()
   TXTPCS.BackColor = vbWhite
End Sub

Private Sub TXTQNTY_GotFocus()
    TXTQNTY.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Opening Quantity"
End Sub

Private Sub TXTQNTY_KeyPress(KeyAscii As Integer)
    If CheckNumericKey(KeyAscii, TXTQNTY, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTQNTY_LostFocus()
TXTQNTY.BackColor = vbWhite
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
   If CheckNumericKey(KeyAscii, txtValue, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtValue_GotFocus()
    txtValue.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Opening Value"
End Sub

Private Sub txtValue_LostFocus()
   txtValue.BackColor = vbWhite
End Sub

Private Function GENOPNCODE() As String
  
  Dim GENRS As New ADODB.Recordset
  Set GENRS = New ADODB.Recordset
           
  If GENRS.State = 1 Then GENRS.Close
  
  GENRS.Open "SELECT ISNULL(Max(VBNO),0) AS VBNO FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
           "' AND VTYP='IVR' ", CN, adOpenDynamic, adLockOptimistic
           
  If GENRS.EOF Then
    GENOPNCODE = "0000000001"
  ElseIf Trim(GENRS!VBNO) = "0" Then  'C1
   GENOPNCODE = "0000000001"
  Else
  
   GENOPNCODE = Val(GENRS!VBNO) + 1
   GENRS.Close
   
   If GENOPNCODE < 10 Then
      GENOPNCODE = "000000000" & GENOPNCODE
   ElseIf GENOPNCODE < 100 Then
      GENOPNCODE = "00000000" & GENOPNCODE
   ElseIf GENOPNCODE < 1000 Then
      GENOPNCODE = "0000000" & GENOPNCODE
   ElseIf GENOPNCODE < 10000 Then
      GENOPNCODE = "000000" & GENOPNCODE
   ElseIf GENOPNCODE < 100000 Then
      GENOPNCODE = "00000" & GENOPNCODE
   ElseIf GENOPNCODE < 1000000 Then
      GENOPNCODE = "0000" & GENOPNCODE
   ElseIf GENOPNCODE < 10000000 Then
      GENOPNCODE = "000" & GENOPNCODE
   ElseIf GENOPNCODE < 100000000 Then
      GENOPNCODE = "00" & GENOPNCODE
   ElseIf GENOPNCODE < 1000000000 Then
     GENOPNCODE = "0" & GENOPNCODE
   Else
      GENOPNCODE = GENOPNCODE
   End If
 End If    'C1
End Function

Private Function STOPEDIT() As Boolean
STOPEDIT = False

Dim STOPRS As ADODB.Recordset
Set STOPRS = New ADODB.Recordset
If STOPRS.State = 1 Then STOPRS.Close
STOPRS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND VTYP='IVR' AND UNIT='" & UNCD & _
            "' AND VBNO='" & M_VBNO & "' AND ISS_QNTY=0 AND RET_PTY_QNTY=0 AND RET_DPT_QNTY=0 ", CN, adOpenDynamic, adLockOptimistic
If STOPRS.EOF Then
    STOPEDIT = True
    MsgBox "FURTHER ENTRY EXIST", vbCritical
End If
End Function

Private Sub EditHelp()

Dim RS As New ADODB.Recordset
    Dim SPECI As String
    Dim MRGN As String
    Dim igcd As String
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT *  FROM ITMMST WHERE CODE = '" & icode & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
    igcd = RS!igcd
    End If
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM IGMMST WHERE CODE = '" & igcd & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
    SPECI = RS!SPECIFICATION
    MRGN = RS!MERGE
    End If
    
    
    If MRGN = "Y" Then
       TXTMRGN.Enabled = True
    Else
       TXTMRGN.Enabled = False
    End If
       
    If Val(SPECI) = 0 Then
       TXTPCS.Enabled = True
       TXTQNTY.Enabled = True
       txtCops.Enabled = False
    ElseIf Val(SPECI) = 1 Then
       TXTQNTY.Enabled = True
       TXTPCS.Enabled = False
       txtCops.Enabled = False
    ElseIf Val(SPECI) = 2 Then
       TXTPCS.Enabled = True
       txtCops.Enabled = True
       TXTQNTY.Enabled = True
    ElseIf Val(SPECI) = 3 Then
       txtCops.Enabled = True
       TXTQNTY.Enabled = True
       TXTPCS.Enabled = False
    End If


End Sub

VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "fraplus1.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_DivMacItmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Opening : Division + Machine Wise WIP (Work In Process)"
   ClientHeight    =   6705
   ClientLeft      =   1755
   ClientTop       =   2595
   ClientWidth     =   10365
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10365
   Begin MSAdodcLib.Adodc ADOOPN 
      Height          =   330
      Left            =   1560
      Top             =   6840
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   10095
      Begin WelchButton.lvButtons_H CMDDEL 
         Height          =   495
         Left            =   4200
         TabIndex        =   15
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
         Image           =   "frm_DivMacItmOpen.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDSAVE 
         Height          =   495
         Left            =   1080
         TabIndex        =   13
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
         Image           =   "frm_DivMacItmOpen.frx":059A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5760
         TabIndex        =   16
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
         Image           =   "frm_DivMacItmOpen.frx":0B34
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2640
         TabIndex        =   19
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
         Image           =   "frm_DivMacItmOpen.frx":10CE
         cBack           =   -2147483633
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5953
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
         Bindings        =   "frm_DivMacItmOpen.frx":1520
         Height          =   3015
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5318
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
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
      Height          =   2355
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   10095
      Begin VB.TextBox txtMachine 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   5850
      End
      Begin VB.TextBox TXTDVCD 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   5850
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   11
         Top             =   1920
         Width           =   1395
      End
      Begin VB.TextBox TXTQNTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   1920
         Width           =   1155
      End
      Begin VB.TextBox TXTPCS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1920
         Width           =   1275
      End
      Begin VB.TextBox TXTNAME 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1560
         Width           =   5850
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Machine :"
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
         TabIndex        =   2
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Division + Machine Wise WIP (Work In Process) Opening Stock"
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
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   8535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Division :"
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
         TabIndex        =   0
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         TabIndex        =   10
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   2880
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pieces :"
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
         TabIndex        =   6
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         TabIndex        =   4
         Top             =   1560
         Width           =   660
      End
   End
End
Attribute VB_Name = "frm_DivMacItmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_DVCD As String, M_VBNO As String, STR As String
Dim spara As String, SQL As String, ICODE As String
Dim OPNDATE As Date
Dim I As Integer

Private Sub cmdCancel_Click()
    TXTNAME.Text = Empty
    txtMachine.Text = Empty
    TXTQNTY.Text = Empty
    TXTQNTY.Tag = Empty
    txtValue.Text = Empty
    TXTPCS.Text = Empty
    txtMachine.SetFocus
End Sub

Private Sub FLEX_DblClick()
If Flex.Rows <= 1 Or Trim(Flex.TextMatrix(1, 0)) = Empty Or Flex.ROW < 1 Then Exit Sub
    TXTDVCD.Text = Flex.TextMatrix(Flex.ROW, 0)
    TXTDVCD.Tag = Flex.TextMatrix(Flex.ROW, 6)
    
    txtMachine.Text = Flex.TextMatrix(Flex.ROW, 1)
    txtMachine.Tag = Flex.TextMatrix(Flex.ROW, 7)
    
    TXTNAME.Text = Flex.TextMatrix(Flex.ROW, 2)
    ICODE = Flex.TextMatrix(Flex.ROW, 8)
    
    Call TXTNAME_LostFocus
End Sub

Private Sub FLEX_KeyDown(KeyCode As Integer, Shift As Integer)
If Flex.Rows <= 1 Or Trim(Flex.TextMatrix(1, 0)) = Empty Or Flex.ROW < 1 Then Exit Sub
    TXTDVCD.Text = Flex.TextMatrix(Flex.ROW, 0)
    TXTDVCD.Tag = Flex.TextMatrix(Flex.ROW, 6)
    
    txtMachine.Text = Flex.TextMatrix(Flex.ROW, 1)
    txtMachine.Tag = Flex.TextMatrix(Flex.ROW, 7)
    
    TXTNAME.Text = Flex.TextMatrix(Flex.ROW, 2)
    ICODE = Flex.TextMatrix(Flex.ROW, 8)
    
    Call TXTNAME_LostFocus
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
End Select
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
 Call CenterChild(frm_Main, Me)
 Dim QUERY As String
   
 QUERY = "SELECT DIVMST.NAME AS DIVISION,MACMST.NAME AS MACHINE,ITMMST.NAME AS ITEM,PCES,QNTY,AMNT,STORETRAN.DVCD," & _
         "STORETRAN.PCOD,STORETRAN.ICOD FROM STORETRAN " & _
         "INNER JOIN DIVMST ON STORETRAN.COMP=DIVMST.COMP AND STORETRAN.UNIT=DIVMST.UNIT AND STORETRAN.DVCD=DIVMST.CODE " & _
         "INNER JOIN MACMST ON STORETRAN.COMP=MACMST.COMP AND STORETRAN.UNIT=MACMST.UNIT AND STORETRAN.DVCD=MACMST.DVCD " & _
         "AND STORETRAN.PCOD=MACMST.CODE " & _
         "INNER JOIN ITMMST ON STORETRAN.ICOD=ITMMST.CODE " & _
         " WHERE STORETRAN.COMP='" & compPth & "' AND STORETRAN.UNIT='" & UNCD & _
         "' AND STORETRAN.VTYP='IVR' AND STORETRAN.RECSTAT<>'D' ORDER BY STORETRAN.VBNO DESC"
   
   ADOOPN.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & M_DBNM & ";Data Source=" & ServerName & ""
   ADOOPN.CommandType = adCmdText
   ADOOPN.RecordSource = QUERY
   ADOOPN.Refresh
   
   Call SetFlex
End Sub

Private Sub Flex_EnterCell()
   Flex.CellBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub Flex_LeaveCell()
   Flex.CellBackColor = vbWhite
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("0012", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
       
    'CHECK DATA
    If txtMachine = Empty Or TXTDVCD = Empty Or TXTNAME = Empty Or (Val(TXTPCS) = 0 And Val(TXTQNTY) = 0) Then
       If TXTDVCD = Empty Then TXTDVCD.SetFocus
       If txtMachine = Empty Then txtMachine.Enabled = True: txtMachine.SetFocus
       If TXTNAME = Empty Then TXTNAME.SetFocus
       If (Val(TXTPCS) = 0 And Val(TXTQNTY) = 0) Then TXTQNTY.SetFocus
       Exit Sub
    End If
    '-------------------
    
    'SET OPENING DATE
      OPNDATE = FSDT - 1
    '-------------------
       
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
    
    Call SetItemInfo(RATE)
     

        If spara = "NEW" Then
                        
          SQL = "INSERT INTO STORETRAN(COMP,UNIT,VTYP,DBCD,VBNO,SRCH,DVCD,DATE,PCOD,ICOD,PCES,QNTY,AMNT,QORP," & _
                "OPER,SYSR,[USER],RATE,GWGT,RECSTAT) VALUES ('" & compPth & "','" & UNCD & "','IVR','XXXXXX','" & M_VBNO & _
                "',1,'" & TXTDVCD.Tag & "','" & Format(OPNDATE, "mm/dd/yyyy") & "','" & txtMachine.Tag & "','" & ICODE & _
                "'," & Val(TXTPCS) & "," & Val(TXTQNTY) & "," & Val(txtValue) & ",'Q','+','T','" & cUName & _
                "'," & RATE & "," & Val(TXTQNTY) & ",'A')"
                  
            CN.Execute SQL
            '----------------
            'DAILYSTAT
             Call DAILYSTATUS("IVR", ICODE, "", Val(TXTQNTY), M_VBNO, 0, cUName, "N", Now, Now)
             '---------------
        Else
            
            SQL = "UPDATE STORETRAN SET PCOD='" & txtMachine.Tag & "',DVCD='" & TXTDVCD.Tag & _
            "',RATE='" & RATE & "',GWGT='" & Val(TXTQNTY) & "', QNTY='" & Val(TXTQNTY) & _
            "',PCES='" & Val(TXTPCS) & "',AMNT='" & Val(txtValue) & "' WHERE COMP='" & compPth & "' AND " & _
            "UNIT='" & UNCD & "' AND VTYP='IVR' AND VBNO='" & M_VBNO & "' AND ICOD='" & ICODE & "' AND RECSTAT='A'"
            
            CN.Execute SQL
            '------------------
            'DAILYSTAT
            Call DAILYSTATUS("IVR", ICODE, "", Val(TXTQNTY), M_VBNO, 0, cUName, "M", Now, Now)
            '------------------
        End If
        
    CN.CommitTrans
    ADOOPN.Refresh
       
    Call cmdCancel_Click
    Exit Sub

LAST:
    MsgBox Err.Description
    CN.RollbackTrans
End Sub

Private Sub CMDDEL_Click()
On Error GoTo LAST
    
    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("000011", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    If spara <> "EDIT" Then Exit Sub
           
    'CHECK DATA
    If TXTDVCD = Empty Or TXTNAME = Empty Then
       If TXTDVCD = Empty Then TXTDVCD.SetFocus
       If txtMachine = Empty Then txtMachine.SetFocus
       If TXTNAME = Empty Then TXTNAME.SetFocus
       Exit Sub
    End If
    '-------------------
    
    STR = MsgBox("ARE YOU SURE YOU WANT TO DELETE THIS ITEM OPENING DETAIL ?", vbYesNo + vbQuestion, "Remove This Opening Detail ?")
    
    CN.BeginTrans
    
    If STR = vbYes Then
       CN.Execute "DELETE STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                  "' AND VTYP='IVR' AND VBNO = '" & M_VBNO & "' AND RECSTAT<>'D'"
                  
       CN.Execute "DELETE FROM GRNTRAN WHERE COMP='" & compPth & "' AND VTYP='IVR' AND UNIT='" & UNCD & _
                  "' AND VBNO='" & M_VBNO & "'"
    End If

       
       CN.CommitTrans
    

    Call cmdCancel_Click
    
    ADOOPN.Refresh
    Exit Sub

LAST:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Function SetFlex()
    With Flex
    
        .ColWidth(0) = 2000
        .ColAlignment(0) = 1
        .TextMatrix(0, 0) = "Division"
        
        .ColWidth(1) = 2000
        .ColAlignment(1) = 1
        .TextMatrix(0, 1) = "Machine"
        
        .ColWidth(2) = 2800
        .TextMatrix(0, 2) = "Item"
        .ColAlignment(2) = 1
    
        .ColWidth(3) = 800
        .ColAlignment(3) = 0
        .TextMatrix(0, 3) = "Pieces"
  
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "Quantity"
        .ColAlignment(4) = 0
  
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = "Amount"
        .ColAlignment(5) = 0
        
        .ColWidth(6) = 0
        .TextMatrix(0, 6) = "DVCD"
        
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = "MCCD"
        
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = "ICOD"
        
    End With
End Function

Private Sub txtDVCD_GotFocus()
  TXTDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
        
    If KeyCode = vbKeyF2 Or (Trim(TXTDVCD) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTDVCD.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & _
        "' AND UNIT='" & UNCD & "' and RECSTAT='A' AND CODE<>'000001'", 0, TXTDVCD.Text, "SELECT DIVISION FROM LIST")
        TXTDVCD.Tag = Key
        DIVCOD = Key
    End If
        
    Me.KeyPreview = True
End Sub

Private Sub txtDVCD_LostFocus()
  TXTDVCD.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
  txtMachine.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
        
    If TXTDVCD = Empty Or TXTDVCD.Tag = Empty Then
       TXTDVCD.Enabled = True
       TXTDVCD.SetFocus
       Exit Sub
    End If
        
    If KeyCode = vbKeyF2 Or (Trim(txtMachine) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtMachine.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM MACMST WHERE COMP='" & compPth & _
        "' AND UNIT='" & UNCD & "' AND DVCD='" & TXTDVCD.Tag & "'", 0, txtMachine.Text, "SELECT MACHINE FROM LIST")
        
        txtMachine.Tag = Key
    End If
    Me.KeyPreview = True
End Sub

Private Sub txtMACHINE_LostFocus()
  txtMachine.BackColor = vbWhite
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
        TXTNAME.Text = SearchITEMLIST("select TOP 20 code,name from ITMMST", 0, "", "List Of Items")
        ICODE = Key
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TXTNAME.Text = "" Then
        M_DESC = Empty
        Key = Empty
        TXTNAME.Text = SearchITEMLIST("select TOP 20 code,name from ITMMST", 0, "", "List Of Items")
        ICODE = Key
    End If
End Sub

Private Sub TXTNAME_LostFocus()
    TXTNAME.BackColor = vbWhite
    If TXTNAME.Text = "" Then Exit Sub
    SQL = "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND STORETRAN.UNIT='" & UNCD & _
          "' AND VTYP='IVR' AND ICOD='" & ICODE & "' AND DVCD='" & TXTDVCD.Tag & _
          "'  AND PCOD='" & txtMachine.Tag & "' AND RECSTAT<>'D'"
          
    With RS
        If .State = adStateOpen Then .Close
        .Open SQL, CN, adOpenDynamic, adLockOptimistic
        If .EOF = True Then
            spara = "NEW"
            TXTPCS = Empty
            TXTQNTY = Empty
            txtValue = Empty
            Exit Sub
        End If
    
        TXTPCS = !PCES
        TXTQNTY = !QNTY
        TXTQNTY.Tag = !QNTY
        txtValue = !AMNT
        M_VBNO = !VBNO
    
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

Private Sub SetItemInfo(RATE As Double)
On Error GoTo LAST

If spara = "NEW" Then

    SQL = "INSERT INTO GRNTRAN([COMP],[UNIT],[VTYP],[VBNO],[DBCD],[SRCH],DATE,[ICOD],[RATE]," & _
    "[GRN_QNTY],[NETRATE],[ISS_QNTY],[BAL_QNTY]) VALUES('" & compPth & "','" & UNCD & "','IVR','" & M_VBNO & _
    "','XXXXXX','1','" & Format(OPNDATE, "yyyy-MM-dd hh:mm:ss") & "','" & ICODE & "','" & RATE & _
    "','" & Val(TXTQNTY) & "','" & Val(RATE) & "','" & Val(TXTQNTY) & "','0')"
    
    Call SetItemBalQty("BALQ", ICODE, Val(TXTQNTY), "+")
Else
    SQL = "UPDATE GRNTRAN SET ICOD='" & ICODE & "',RATE='" & RATE & "',NETRATE='" & RATE & _
    "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & "' AND VTYP='IVR' AND VBNO = '" & M_VBNO & "'"
    
    Call SetItemBalQty("BALQ", ICODE, Val(TXTQNTY.Tag), "-")
    Call SetItemBalQty("BALQ", ICODE, Val(TXTQNTY), "+")
    
End If
  
CN.Execute SQL

Exit Sub
LAST:
CN.RollbackTrans
MsgBox Err.Description
End Sub

Private Function STOPEDIT() As Boolean
STOPEDIT = False
'Dim STOPRS As ADODB.Recordset
'Set STOPRS = New ADODB.Recordset
'If STOPRS.State = 1 Then STOPRS.Close
'STOPRS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND VTYP='IVR' AND UNIT='" & UNCD & _
'            "' AND VBNO='" & M_VBNO & "' AND ISS_QNTY=0 AND RET_PTY_QNTY=0 AND RET_DPT_QNTY=0 ", CN, adOpenDynamic, adLockOptimistic
'If STOPRS.EOF Then
'    STOPEDIT = True
'    MsgBox "FURTHER ENTRY EXIST", vbCritical
'End If
End Function


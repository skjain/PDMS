VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmGRNWOAcEffect 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRN Entry Without A/c Effect"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11370
   Begin VB.TextBox TXTRMRK 
      Height          =   285
      Left            =   1200
      MaxLength       =   250
      TabIndex        =   19
      Top             =   5760
      Width           =   10095
   End
   Begin VB.Frame frm_head 
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   11175
      Begin VB.TextBox TXTITOT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   540
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   4545
         Width           =   2505
      End
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmddelitm 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Remove Item"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox TXTTPCS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox TXTTQTY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4680
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
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
         Format          =   53608449
         CurrentDate     =   39347
      End
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBLNET 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7080
         TabIndex        =   18
         Top             =   4680
         Width           =   1200
      End
      Begin VB.Label LBLBILLNO 
         BackStyle       =   0  'Transparent
         Caption         =   "&GRN No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LBLBILLDATE 
         BackStyle       =   0  'Transparent
         Caption         =   "G&RN Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Total Pcs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Total Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   4800
         Width           =   1455
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
      Image           =   "frmGRNWOAcEffect.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "E&dit"
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
      Image           =   "frmGRNWOAcEffect.frx":039A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   6120
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
      Image           =   "frmGRNWOAcEffect.frx":0734
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   6120
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
      Image           =   "frmGRNWOAcEffect.frx":0ACE
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   6120
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
      Image           =   "frmGRNWOAcEffect.frx":1068
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   6120
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
      Image           =   "frmGRNWOAcEffect.frx":14BA
      cBack           =   -2147483633
   End
   Begin VB.Label LBLHEAD 
      Caption         =   "GRN of Raw Material Without A/c Effect "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   11175
   End
   Begin VB.Label LBLRMRK 
      Caption         =   "Remarks : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmGRNWOAcEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Emptycell As Boolean
Dim SAVEFLAG As Boolean
Public M_DBCD  As String
Public M_ISSUE As String
Dim M_CCCD As String

Private Sub cmdAdd_Click()
     
  SAVEFLAG = True
  btn_sts (False)
  cmddelitm.Enabled = False
   
  M_DBCD = "000005"
  TXTVBNO = GenVNO("IVR", M_DBCD)
  
  FLEX.SetFocus
  FLEX.COL = 1
  FLEX.ROW = 1
  
End Sub

Private Sub cmdCancel_Click()
  ClsData (Me)
  FLEX.Rows = 1
  FLEX.Rows = 2
  btn_sts (True)
  cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000055", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  SAVEFLAG = False
  btn_sts (False)
  FrmGRNWOAEList.Show 1
      
  'FIFO
  If Trim(M_ISSUE) = "Y" Then
     MsgBox "You Can Not Delete this GRN !! Issue Entry Exist!!", vbExclamation, "Access Denied"
     Call cmdCancel_Click
     Exit Sub
  End If
  '------------
     
     Dim AYS
     AYS = MsgBox("Are you sure to delete the GRN ", vbYesNo)
     If AYS = vbYes Then
        CN.BeginTrans

        CN.Execute "UPDATE GRN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' and VTYP = 'IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
        
        CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' and VTYP = 'IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
        
        CN.Execute "DELETE FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and VTYP = 'IVR' " & _
                   "AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'"
  
        Call DAILYSTATUS("IVR", "", M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "D", Now, TXTVBDT)
       
        CN.CommitTrans
     End If
  Call cmdCancel_Click
End Sub

Private Sub cmddelitm_Click()
  If FLEX.ROW > 1 Then
    FLEX.RemoveItem (FLEX.ROW)
    TXTTPCS.Text = 0
    TXTTQTY.Text = 0
    TXTITOT.Text = 0
    Dim i As Double
    i = 1
    For i = 1 To FLEX.Rows - 1
      FLEX.TextMatrix(i, 0) = i
      TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 2)), "######")
      TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 3)), "########.000")
      TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 5)), "########.00")
    Next
    FLEX.Refresh
    FLEX.ROW = FLEX.Rows - 1
    FLEX.COL = 5
    FLEX.SetFocus
  End If
  cmddelitm.Enabled = False
End Sub

Private Sub cmdEdit_Click()
  SAVEFLAG = False
  M_ISSUE = "N"
  
  FrmGRNWOAEList.Show 1
  
  If M_DBCD <> Empty Then
     FLEX.SetFocus
     FLEX.ROW = 1
     FLEX.COL = 1
     btn_sts (False)
  Else
    Call ClsData(Me)
    btn_sts (True)
    cmdAdd.SetFocus
    Exit Sub
  End If
  
  'FIFO
  If Trim(M_ISSUE) = "Y" Then
     MsgBox "You Can Not Edit GRN Qnty !! Issue Entry Exist!!", vbExclamation, "Access Denied"
     Exit Sub
  End If
  '------------
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
  On Error GoTo LAST
  
  If CHKSAVEDATA = False Then
    Exit Sub
  End If
       
  If SAVEFLAG = True Then
    TXTVBNO = GenVNO("IVR", M_DBCD)
    If RS.State Then RS.Close
    RS.Open "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' and VTYP = 'IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       MsgBox "GRN No. " & TXTVBNO & " Already Exist. Check Last No. In Unit Configuration", vbCritical
       Exit Sub
    End If
    RS.Close
  End If
       
  Call SAVERECIVR
  
  If SAVEFLAG = True Then
    MsgBox "Your GRN No. is " + TXTVBNO.Text
  Else
    MsgBox "GRN No. " + TXTVBNO.Text + " Successfully Edited."
  End If
  
  Call cmdCancel_Click
  
  Exit Sub
LAST:
  MsgBox ERR.Description
  CN.RollbackTrans
End Sub

Private Sub Flex_Click()
  cmddelitm.Enabled = True
End Sub

Private Sub FLEX_EnterCell()
  FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub FLEX_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case FLEX.COL
  Case 1
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    If KeyCode = vbKeyF2 Then
        FLEX.TextMatrix(FLEX.ROW, 1) = SearchList1("Select Code,Name From ITMMST", 0, Empty, "Select Item")
        FLEX.TextMatrix(FLEX.ROW, 6) = Key
        If key_PressNew = True Then
           M_DESC = ""
           Key = ""
           FLEX.TextMatrix(FLEX.ROW, 1) = ""
           frm_Item.Show
        End If
    End If
End Select
End Sub

Private Sub Flex_LeaveCell()
  Dim i As Double
  FLEX.CellBackColor = vbWhite
  
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Val(Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 2)), "######"))
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 3)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 5)), "########.00")
  Next
  
End Sub

Private Sub FLEX_LostFocus()
  Dim i As Double
    
  FLEX.CellBackColor = vbWhite
  
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Val(Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 2)), "######"))
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 3)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 5)), "########.00")
  Next
  
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 Me.BackColor = RGB(RED, GREEN, BLUE)
 TXTVBNO.FontSize = 10: TXTVBNO.FontBold = True
 TXTTPCS.FontSize = 12: TXTTPCS.FontBold = True
 TXTTQTY.FontSize = 12: TXTTQTY.FontBold = True
 TXTITOT.FontSize = 18: TXTITOT.FontBold = True
 TXTITOT.ForeColor = vbRed
 LBLHEAD.FontSize = 14
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  Me.Left = 130
  Emptycell = True
    
  TXTVBDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
   
  Call setflexhead
  Call btn_sts(True)
  
  Me.KeyPreview = True
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If UCase(ActiveControl.NAME) = "FLEX" Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub setflexhead()

    FLEX.TextMatrix(0, 0) = "Sr."
    FLEX.TextMatrix(0, 1) = "Item Name"
    FLEX.TextMatrix(0, 2) = "Pcs"
    FLEX.TextMatrix(0, 3) = "Qnty"
    FLEX.TextMatrix(0, 4) = "Rate"
    FLEX.TextMatrix(0, 5) = "Amt.(For Valuation)"
    FLEX.TextMatrix(0, 6) = "ICOD"
    
    FLEX.ColWidth(0) = 350
    FLEX.ColWidth(1) = 4100
    FLEX.ColWidth(2) = 600
    FLEX.ColWidth(3) = 1300
    FLEX.ColWidth(4) = 1300
    FLEX.ColWidth(5) = 2100
    FLEX.ColWidth(6) = 0
        
    FLEX.ColAlignment(1) = 0
    FLEX.ColAlignment(2) = 1
    FLEX.ColAlignment(3) = 1
    FLEX.ColAlignment(4) = 1
    FLEX.ColAlignment(5) = 1
    
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
  On Error GoTo LAST
    
  If FLEX.COL = 5 And KeyAscii <> 13 Then
     Exit Sub
  End If
    
  FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.ROW
  
  'CONTROL VARIABLE-----------------------------------------------------
  Dim ALLOW_KEY As Boolean, FWD_COL As Boolean, ENTER_PRESS As Boolean
  'DEFAULT VALUE
  FWD_COL = False: ALLOW_KEY = False
  '---------------------------------------------------------------------
    
  'LOCAL RECORD SET
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  '--------------------------------
    
  'USER DOESN'T ENTERED MORE THAN ONE DECIMAL:4-PCS
  If FLEX.COL = 2 Or FLEX.COL = 3 Or FLEX.COL = 4 Then
    If InStr(1, FLEX.TextMatrix(FLEX.ROW, FLEX.COL), ".") > 0 And KeyAscii = 46 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
  '--------------------------------------------
  
  'NO IDEA
  If Emptycell = True And (Not KeyAscii = 13) Then
     FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty
     If FLEX.COL = 3 Or FLEX.COL = 4 Then
        FLEX.TextMatrix(FLEX.ROW, 5) = 0
     End If
     Emptycell = False
  End If
  
  '----------------------------------------------
  'COLUMN WISE ENTERED TEXT CHECKING : ALLOW KEY
  '-----------------------------------------------
  Select Case FLEX.COL
  Case 1
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    If ((KeyAscii = vbKeyF2) Or (KeyAscii = 13 And Trim(FLEX.TextMatrix(FLEX.ROW, 1)) = Empty)) Then
        FLEX.TextMatrix(FLEX.ROW, 1) = SearchList1("Select Code,Name From ITMMST", 0, Empty, "Select Item")
        FLEX.TextMatrix(FLEX.ROW, 6) = Key
        If key_PressNew = True Then
           M_DESC = ""
           Key = ""
           FLEX.TextMatrix(FLEX.ROW, 1) = ""
           frm_Item.Show
        End If
    End If
    ALLOW_KEY = True
    
   Case 2 'SIMPLE NUMBER WITHOUT DECIMAL
    If (KeyAscii >= 48 And KeyAscii <= 57) Then             ' 0- 9
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   
   Case 3, 4 'DECIMAL NUMBER
    
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '. & -
      ALLOW_KEY = True
    ElseIf Len(Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))) = 0 And KeyAscii = 45 Then
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
    
   Case 5      'AMOUNT
    ALLOW_KEY = False
   
  End Select
  
  '-------------------------------------------------------------------------------------------------
  'ENTER PRESS
  '-------------------------------------------------------------------------------------------------
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
  Else
    ENTER_PRESS = False
  End If
  
  '-------------------------------------------------------------------------------------------------
  'BACK SPACE : COMES FIRST THEN KEYASCII = 0
  '-------------------------------------------------------------------------------------------------
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))
    If lnth > 0 Then
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Mid(FLEX.TextMatrix(FLEX.ROW, FLEX.COL), 1, lnth - 1)
      If FLEX.COL = 3 Or FLEX.COL = 4 Then
          FLEX.TextMatrix(FLEX.ROW, 5) = Val(FLEX.TextMatrix(FLEX.ROW, 3)) * Val(FLEX.TextMatrix(FLEX.ROW, 4))
          FLEX.TextMatrix(FLEX.ROW, 5) = Trim(nstr(FLEX.TextMatrix(FLEX.ROW, 5), 12, 2))
      End If
      Exit Sub
    End If
  End If
  
  '-------------------------------------------------------------------------------------------------
  'RESULT OF ALLOW KEY AND ENTER PRESS
  '-------------------------------------------------------------------------------------------------
  If ENTER_PRESS = False Then
     If ALLOW_KEY = True Then
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) + Chr(KeyAscii)
        If FLEX.COL = 3 Or FLEX.COL = 4 Then
          FLEX.TextMatrix(FLEX.ROW, 5) = Val(FLEX.TextMatrix(FLEX.ROW, 3)) * Val(FLEX.TextMatrix(FLEX.ROW, 4))
          FLEX.TextMatrix(FLEX.ROW, 5) = Trim(nstr(FLEX.TextMatrix(FLEX.ROW, 5), 12, 2))
        End If
     Else
        KeyAscii = 0
        Exit Sub
     End If
  End If
    
  '=================================================================================================
  'FORWARD MOVE FROM ONE COLUMN TO ANOTHER : IS TRUE OR FALSE ?? ON BASIS OF ENTER PRESS
  '=================================================================================================
   
  FWD_COL = False
  
  If ENTER_PRESS = True Then '-------------------------MAIN
    Select Case FLEX.COL
    Case 2, 3, 4, 5
         If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
            FWD_COL = True
         End If
    Case 1
            FWD_COL = True
    End Select
  
  '-----------------------------------------------SUB
  ' RESULT OF FORWARD COLUMN
  '-----------------------------------------------SUB
  
  If FWD_COL = True Then
     If FLEX.COL = 5 Then
        FLEX.TextMatrix(FLEX.ROW, 4) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 5)) / Val(FLEX.TextMatrix(FLEX.ROW, 3)), "#########.00")
                
        Dim AYS
        AYS = MsgBox("Want to Add More Item ", vbYesNo)
        If AYS = vbYes Then
          FLEX.Rows = FLEX.Rows + 1
          FLEX.ROW = FLEX.Rows - 1
          FLEX.COL = 1
          FLEX.SetFocus
         Else
          TXTRMRK.SetFocus
         End If
         Exit Sub
     Else
          FLEX.COL = FLEX.COL + 1
     End If
  End If
  '-------------------------------------------------------SUB
  
  Emptycell = True
  End If
  
  '-------------------------------------------------------MAIN
  Exit Sub
LAST:
  MsgBox "Error In Item Detail"
  FLEX.SetFocus
  Exit Sub
End Sub

Private Function CHKSAVEDATA() As Boolean
  Dim CHKRS As New ADODB.Recordset
  Set CHKRS = New ADODB.Recordset
  Dim i As Long
  
  For i = 1 To FLEX.Rows - 1
    If CHKRS.State = 1 Then CHKRS.Close
    CHKRS.Open "SELECT * from ITMMST WHERE CODE='" & FLEX.TextMatrix(i, 6) & "'", CN, adOpenKeyset, adLockPessimistic
    If CHKRS.EOF Then
        MsgBox "Item Not Define ", vbCritical
        CHKSAVEDATA = False
        Exit Function
    End If
    
    If Not IsNumeric(FLEX.TextMatrix(i, 3)) Then
       MsgBox "Quantity Not Define ", vbCritical
       CHKSAVEDATA = False
       FLEX.ROW = i
       FLEX.COL = 3
       Exit For
    End If
    
    If Not IsNumeric(FLEX.TextMatrix(i, 4)) Then
       MsgBox "Rate Not Define ", vbCritical
       CHKSAVEDATA = False
       FLEX.ROW = i
       FLEX.COL = 4
       Exit For
    End If
    
  Next
      
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT VBNO FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' and VTYP = 'IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & _
  "' ", CN, adOpenDynamic, adLockOptimistic
  
  If Not CHKRS.EOF Then
    If CHKRS!VBNO <> TXTVBNO Then
      MsgBox "Duplicate GRN No. !!!! ", vbCritical
      CHKSAVEDATA = False
      Exit Function
    End If
  End If
  
  CHKSAVEDATA = True
End Function

Private Sub SAVERECIVR()
  
  On Error GoTo LAST
  Dim SQL As String
  
  Dim SAVDAT As New ADODB.Recordset 'USE
  Dim MSTDAT As New ADODB.Recordset
        
  Dim i As Double
  Dim J As Double
  Set SAVDAT = New ADODB.Recordset 'USE
  Set MSTDAT = New ADODB.Recordset
  
  CN.BeginTrans
  
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and VTYP = 'IVR' " & _
  "AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'"
      
  CN.Execute "DELETE FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and VTYP = 'IVR' " & _
  "AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'"
      
  SQL = Empty
    
      
 '===================================================================================================
 'STORETRAN (ADD)
 '===================================================================================================
 Dim RATE As Double
 
 i = 1
 For i = 1 To FLEX.Rows - 1
 
 RATE = Val(FLEX.TextMatrix(i, 4))
  
 CN.Execute "INSERT INTO STORETRAN(COMP,UNIT,VTYP,SRCH,VBNO,CHLN,CHDT,DATE,DBCD,CRAC,DPTC,CSHD,CSCD,DRAC,PCOD," & _
            "ICOD,PCES,QNTY,RATE,AMNT,QORP,[USER],SYSR,OPER,DVCD,RECSTAT) VALUES('" & compPth & _
            "','" & UNCD & "','IVR'," & i & ",'" & TXTVBNO.Text & "','" & TXTVBNO.Text & _
            "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
            "','" & M_DBCD & "','','','','','','','" & FLEX.TextMatrix(i, 6) & "','" & Val(FLEX.TextMatrix(i, 2)) & _
            "','" & Val(FLEX.TextMatrix(i, 3)) & "','" & RATE & "','" & Val(FLEX.TextMatrix(i, 5)) & _
            "','Q','" & cUName & "','" & IIf(SAVEFLAG = True, "N", "U") & _
            "','+','000001','A')"
            ',DTTM --- >> ','" & Format(TXTVBDT, "YYYY/MM/DD HH:MM:SS AMPM") & "'
 Next i
     
  '------------FIFO----------------------
   Call SetItemInfo
  
  '===================================================================================================
  'AUTOMATION ENTRY FOR SERVICE TAX
  '===================================================================================================
  
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND VTYP='IVR' AND DBCD ='" & M_DBCD & "' AND VBNO = '" & Trim(TXTVBNO) & _
              "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  
  SAVDAT!COMP = compPth
  SAVDAT!unit = UNCD
  SAVDAT!DVCD = "000001"
  SAVDAT!VTYP = "IVR"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!dbcd = M_DBCD
  SAVDAT!Date = Format(TXTVBDT.Value, "YYYY/MM/DD")
  SAVDAT!VBNO = Trim(TXTVBNO)
  SAVDAT!DRAC = ""
  SAVDAT!PCOD = ""
  SAVDAT!TPCS = Val(TXTTPCS)
  SAVDAT!TQTY = Val(TXTTQTY)
  SAVDAT!ITOT = Val(TXTITOT)
  SAVDAT!BADJ = 0
  SAVDAT!BNET = Val(TXTITOT)
  SAVDAT!ACEFFECT = "N"
  SAVDAT!DPTC = ""
  SAVDAT!CHEAD = ""
  SAVDAT!CSCD = ""
  SAVDAT!BRMK = Trim(TXTRMRK)
  SAVDAT!BSTS = "P"
  SAVDAT!CVBN = ""
  
  If SAVEFLAG = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  SAVDAT![User] = cUName & ""
  SAVDAT!RECSTAT = "A"
  SAVDAT.Update

  'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
    Call SetSRNO(TXTVBNO, "IVR", M_DBCD)
  End If
 '-------------------------
 'DAILYSTATUS ENTRY
  If SAVEFLAG = True Then
   Call DAILYSTATUS("IVR", "", M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "N", Now, TXTVBDT)
  Else
   Call DAILYSTATUS("IVR", "", M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "M", Now, TXTVBDT)
  End If
 '-------------------------
  CN.CommitTrans
  Exit Sub
  
LAST:
 MsgBox ERR.Description
 If SAVDAT.State = 1 Then
   SAVDAT.CancelUpdate
   SAVDAT.Close
 End If
 CN.RollbackTrans
End Sub

Private Sub SetItemInfo()
On Error GoTo LAST

Dim INDEX As Long
Dim SQL As String
Dim RATE As Double

With FLEX
 
For INDEX = 1 To .Rows - 1
    
    SQL = "INSERT INTO GRNTRAN([COMP],[UNIT],[VTYP],[VBNO],[DBCD],[SRCH],DATE,[ICOD],[RATE],[GRN_QNTY],[NETRATE],[BAL_QNTY])"
    SQL = SQL & " VALUES('" & compPth & "','" & UNCD & "','IVR','" & TXTVBNO & _
    "','" & M_DBCD & "','" & INDEX & "','" & Format(TXTVBDT, "yyyy-MM-dd hh:mm:ss") & _
    "','" & Trim(.TextMatrix(INDEX, 6)) & "','" & Val(.TextMatrix(INDEX, 4)) & "','" & Val(.TextMatrix(INDEX, 3)) & _
    "','" & Val(.TextMatrix(INDEX, 4)) & "','" & Val(.TextMatrix(INDEX, 3)) & "')"
  
CN.Execute SQL
  
Next INDEX
 
End With
Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
End If
End Sub

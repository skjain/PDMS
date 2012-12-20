VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form FrmDeliveryAddress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consignee Address Master ( May Have Muliple Address For a Party )"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11145
   Begin VB.TextBox txtname 
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
      Left            =   3120
      MaxLength       =   50
      TabIndex        =   6
      Top             =   840
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   360
      TabIndex        =   18
      Top             =   1440
      Width           =   10455
      Begin VB.TextBox TXTAREA 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox TXTPHONE 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TXTECC 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TXTLST 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox TXTCST 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox ADD1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   7
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox ADD3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox ADD2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   8
         Top             =   840
         Width           =   4095
      End
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
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
      Begin WelchButton.lvButtons_H cmdOK 
         Height          =   495
         Left            =   9240
         TabIndex        =   15
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "A&dd"
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
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdRemove 
         Height          =   495
         Left            =   9240
         TabIndex        =   16
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "&Remove"
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
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         Caption         =   "Area "
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
         Left            =   720
         TabIndex        =   28
         Tag             =   "S"
         Top             =   1635
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Phone No."
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
         Left            =   5520
         TabIndex        =   27
         Tag             =   "S"
         Top             =   555
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Ecc No."
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
         Left            =   5520
         TabIndex        =   26
         Tag             =   "S"
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "LST No."
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
         Left            =   5520
         TabIndex        =   25
         Tag             =   "S"
         Top             =   1275
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "CST No."
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
         Left            =   5520
         TabIndex        =   24
         Tag             =   "S"
         Top             =   1635
         Width           =   855
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   5400
         X2              =   5400
         Y1              =   360
         Y2              =   2040
      End
      Begin VB.Label Label2 
         Caption         =   "Address 1."
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
         Left            =   240
         TabIndex        =   21
         Tag             =   "S"
         Top             =   480
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         Height          =   1695
         Left            =   120
         Top             =   360
         Width           =   10215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   9000
         X2              =   9000
         Y1              =   360
         Y2              =   2040
      End
      Begin VB.Label Label3 
         Caption         =   "2."
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
         Left            =   960
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "3."
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
         Left            =   960
         TabIndex        =   19
         Top             =   1200
         Width           =   255
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   5880
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
      Image           =   "FrmDeliveryAddress.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   5880
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
      Image           =   "FrmDeliveryAddress.frx":039A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   5880
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
      Image           =   "FrmDeliveryAddress.frx":0734
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   5880
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
      Image           =   "FrmDeliveryAddress.frx":0ACE
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   5880
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
      Image           =   "FrmDeliveryAddress.frx":1858
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   5880
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
      Image           =   "FrmDeliveryAddress.frx":1CAA
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "CONSIGNEE NAME    :"
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
      Left            =   600
      TabIndex        =   23
      Top             =   840
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   6375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   11040
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   11040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   240
      X2              =   10920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PARTY - CONSIGNEE ADDRESS MASTER"
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
      TabIndex        =   22
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "FrmDeliveryAddress"
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
Dim ROWNO As Long
Dim SAVEFLAG As Boolean
Dim M_CODE As String

Private Sub ADD1_GotFocus()
  ADD1.BackColor = RGB(BRED, BGREEN, BBLUE)
  ADD1.SelStart = 0
  ADD1.SelLength = Len(ADD1)
  ADD1.ToolTipText = "Enter Party Address1"
End Sub

Private Sub ADD1_LostFocus()
  ADD1.BackColor = vbWhite
End Sub

Private Sub ADD2_GotFocus()
  ADD2.BackColor = RGB(BRED, BGREEN, BBLUE)
  ADD2.SelStart = 0
  ADD2.SelLength = Len(ADD2)
  ADD2.ToolTipText = "Enter Party Address2"
End Sub

Private Sub ADD2_LostFocus()
  ADD2.BackColor = vbWhite
End Sub

Private Sub ADD3_GotFocus()
  ADD3.BackColor = RGB(BRED, BGREEN, BBLUE)
  ADD3.SelStart = 0
  ADD3.SelLength = Len(ADD3)
  ADD3.ToolTipText = "Enter Party Address3"
End Sub

Private Sub ADD3_LostFocus()
  ADD3.BackColor = vbWhite
End Sub





Private Sub cmdAdd_Click()
If M_USRSECLEVL = "1" Then
   If ReadConfigMaster("0008", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
Else
   'Call chk_ldt
End If
    
Call ClsData(Me)
Call btn_sts(False)
Me.Frame1.Enabled = True
TXTNAME.SetFocus
SAVEFLAG = False
cmdCancel.Cancel = True
cmdDelete.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
    FLEX.Rows = 1
    FLEX.Rows = 2
    CMDOK.Caption = "&Add"
    CMDREMOVE.Enabled = False
    SWITCH = False
End Sub

Private Sub cmddelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000005", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  On Error GoTo ErrorHandler
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM SPTRAN WHERE DCOD='" & M_CODE & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then MsgBox "Further Entry Exist": Exit Sub
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM BILLMAIN WHERE DCOD='" & M_CODE & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then MsgBox "Further Entry Exist": Exit Sub
  
  If (MsgBox("Are You sure you want to delete :" & TXTNAME, vbYesNo) = vbYes) Then
     CN.BeginTrans
     CN.Execute "UPDATE PADDMST SET STATUS = 'D',RECSTAT='D' WHERE CODE = '" & Trim(M_CODE) & "'"
     CN.CommitTrans
     MsgBox "Successfully Deleted"
  End If
  cmdCancel_Click
  Exit Sub
ErrorHandler:
    If (ERR.Number = -2147168237) Then
       CN.RollbackTrans
    Else
       MsgBox ERR.Description
    End If
End Sub

Private Sub cmdEdit_Click()
If M_USRSECLEVL = 1 Then
    If ReadConfigMaster("000005", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
End If
SAVEFLAG = True
M_DESC = Empty
Key = Empty
Frm_List.cmdAddNew.Visible = False
M_CODE = SearchList("Select DISTINCT(CODE) ,[NAME] from PADDMST where RECSTAT='A'")
If Trim(M_CODE) = Empty Then MsgBox "No Record Found ": cmdCancel_Click: Exit Sub
Call entry(M_CODE)
Call btn_sts(False)
Me.Frame1.Enabled = True
TXTNAME.SetFocus
End Sub

Private Sub cmdRemove_Click()
Dim CURSOR As Long
Dim J As Long

For J = ROWNO To FLEX.Rows - 2
 FLEX.TextMatrix(J, 0) = FLEX.TextMatrix(J + 1, 0)
 FLEX.TextMatrix(J, 1) = FLEX.TextMatrix(J + 1, 1)
 FLEX.TextMatrix(J, 2) = FLEX.TextMatrix(J + 1, 2)
 FLEX.TextMatrix(J, 3) = FLEX.TextMatrix(J + 1, 3)
 FLEX.TextMatrix(J, 4) = FLEX.TextMatrix(J + 1, 4)
 FLEX.TextMatrix(J, 5) = FLEX.TextMatrix(J + 1, 5)
 FLEX.TextMatrix(J, 6) = FLEX.TextMatrix(J + 1, 6)
 FLEX.TextMatrix(J, 7) = FLEX.TextMatrix(J + 1, 7)
 FLEX.TextMatrix(J, 8) = FLEX.TextMatrix(J + 1, 8)
Next J

FLEX.Rows = FLEX.Rows - 1
Call CLEARDATA
SWITCH = False
ADD1.SetFocus
CMDOK.Caption = "&Add"
CMDREMOVE.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Flex_Click()
If FLEX.Rows > 1 And FLEX.TextMatrix(FLEX.ROW, 0) <> Empty Then
    CMDOK.Caption = "Upd&ate"
    CMDREMOVE.Enabled = True
    ROWNO = FLEX.ROW
      
    txtPhone = FLEX.TextMatrix(ROWNO, 1)
    txtecc = FLEX.TextMatrix(ROWNO, 2)
    txtcst = FLEX.TextMatrix(ROWNO, 3)
    txtlst = FLEX.TextMatrix(ROWNO, 4)
    ADD1 = FLEX.TextMatrix(ROWNO, 5)
    ADD2 = FLEX.TextMatrix(ROWNO, 6)
    ADD3 = FLEX.TextMatrix(ROWNO, 7)
    TXTAREA = FLEX.TextMatrix(ROWNO, 8)
    
    SWITCH = True
End If
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  LBLHEAD.BackColor = &H80&
  LBLHEAD.ForeColor = &HFFFFFF
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  Call btn_sts(True)
  Call SETFLEX
  cmdExit.Cancel = True
  Me.Frame1.Enabled = False
  Me.KeyPreview = True
  M_CODE = Empty
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdSave_Click()
Dim INDEX As Long
   On Error GoTo LAST
             
    If Trim(TXTNAME) = Empty Then
        MsgBox "Please Enter Party !!", vbInformation
        TXTNAME.SetFocus
        Exit Sub
    End If
            
    If Trim(FLEX.TextMatrix(1, 0)) = Empty Then
        MsgBox "Please Enter Detail then Save !!", vbInformation
        ADD1.SetFocus
        Exit Sub
    End If
    
    
   If SAVEFLAG = False Then
    Dim CHKRS As ADODB.Recordset
    Set CHKRS = New ADODB.Recordset
    
    If CHKRS.State = 1 Then CHKRS.Close
    CHKRS.Open "Select * from PADDMST where Upper([NAME])='" & UCase(TXTNAME.Text) & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
    If CHKRS.EOF = False Then MsgBox "This Consinee Name Is Already In Use !!!", vbInformation, App.Title: CHKRS.Close: TXTNAME.SetFocus: Exit Sub
    M_CODE = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From PADDMST WHERE CODE LIKE '0%'")
   End If
         
CN.BeginTrans
CN.Execute "DELETE FROM PADDMST WHERE CODE='" & M_CODE & "'"

Dim ARCD As String
With FLEX
For INDEX = 1 To FLEX.Rows - 2

If RS.State = 1 Then RS.Close
RS.Open "SELECT CODE FROM REFMST WHERE NAME='" & Trim(.TextMatrix(INDEX, 8)) & "' AND CATA='A'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   ARCD = Trim(RS!CODE & "")
Else
   ARCD = Empty
End If

CN.Execute "INSERT INTO PADDMST(CODE,NAME,SRNO,ARCD,ADD1,ADD2,ADD3,ADDR,PHONE,ECCNO,LSTNO,CSTNO,STATUS,RECSTAT) VALUES('" & M_CODE & _
"','" & Trim(TXTNAME) & "','" & INDEX & "','" & ARCD & "','" & .TextMatrix(INDEX, 5) & "','" & .TextMatrix(INDEX, 6) & "','" & .TextMatrix(INDEX, 7) & "','" & Trim(.TextMatrix(INDEX, 5) & .TextMatrix(INDEX, 6) & .TextMatrix(INDEX, 7)) & "','" & .TextMatrix(INDEX, 1) & _
"','" & .TextMatrix(INDEX, 2) & "','" & .TextMatrix(INDEX, 4) & "','" & .TextMatrix(INDEX, 3) & "','A','A')"
Next INDEX
CN.CommitTrans
End With
      
If SAVEFLAG = False Then
   MsgBox "Data Successfully Saved"
   Call DAILYSTATUS("DAD", M_CODE, "", 0, "", 0, cUName, "N", Now, Now)
Else
   MsgBox "Data Successfully Edited"
   Call DAILYSTATUS("DAD", M_CODE, "", 0, "", 0, cUName, "M", Now, Now)
End If

Call RESETALL
Call cmdCancel_Click

Exit Sub
LAST:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub

Private Sub RESETALL()
 TXTNAME = Empty
 Call CLEARDATA
 FLEX.Rows = 1
 FLEX.Rows = 2
 TXTNAME.SetFocus
End Sub

Private Sub SETFLEX()
  FLEX.ColWidth(0) = 4600
  FLEX.ColWidth(1) = 1400
  FLEX.ColWidth(2) = 1400
  FLEX.ColWidth(3) = 1400
  FLEX.ColWidth(4) = 1400
  FLEX.ColWidth(5) = 0
  FLEX.ColWidth(6) = 0
  FLEX.ColWidth(7) = 0
  FLEX.ColWidth(8) = 1400
  
  FLEX.Clear
  FLEX.TextMatrix(0, 0) = "Address"
  FLEX.TextMatrix(0, 1) = "Phone Number"
  FLEX.TextMatrix(0, 2) = "    Ecc No."
  FLEX.TextMatrix(0, 3) = "    CST No."
  FLEX.TextMatrix(0, 4) = "    LST No."
  FLEX.TextMatrix(0, 5) = "Add1"
  FLEX.TextMatrix(0, 6) = "Add2"
  FLEX.TextMatrix(0, 7) = "Add3"
  FLEX.TextMatrix(0, 8) = "Area"
  
  FLEX.ColAlignment(0) = vbLeftJustify
End Sub


Private Sub cmdOk_Click()
 Dim INDEX As Long
 
 If Not SWITCH Then
      ROWNO = FLEX.Rows - 1
 End If
 
 If CheckData(ROWNO) Then Exit Sub
 
    FLEX.TextMatrix(ROWNO, 0) = Trim(ADD1) + " " + Trim(ADD2) + " " + Trim(ADD3)
    FLEX.TextMatrix(ROWNO, 1) = txtPhone
    FLEX.TextMatrix(ROWNO, 2) = txtecc
    FLEX.TextMatrix(ROWNO, 3) = txtcst
    FLEX.TextMatrix(ROWNO, 4) = txtlst
    FLEX.TextMatrix(ROWNO, 5) = Trim(ADD1)
    FLEX.TextMatrix(ROWNO, 6) = Trim(ADD2)
    FLEX.TextMatrix(ROWNO, 7) = Trim(ADD3)
    FLEX.TextMatrix(ROWNO, 8) = Trim(TXTAREA)
    
    If Not SWITCH Then
      FLEX.Rows = FLEX.Rows + 1
    End If
    
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
    Call CLEARDATA
    ADD1.SetFocus
    CMDOK.Caption = "&Add"
    CMDREMOVE.Enabled = False
    SWITCH = False
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long

    If Trim(ADD1) = Empty And Trim(ADD2) = Empty And Trim(ADD3) = Empty And Trim(txtPhone) = Empty And Trim(txtecc) = Empty And Trim(txtlst) = Empty And Trim(txtcst) = Empty And Trim(TXTAREA) = Empty Then
        MsgBox "Enter Data First then ADD", vbInformation
        ADD1.SetFocus
        CheckData = True
        Exit Function
    End If
End Function

Private Sub CLEARDATA()
    ADD1.Text = Empty
    ADD2.Text = Empty
    ADD3.Text = Empty
    txtPhone.Text = Empty
    txtecc.Text = Empty
    txtlst.Text = Empty
    txtcst.Text = Empty
    TXTAREA.Text = Empty
End Sub

Private Sub TXTAREA_GotFocus()
    TXTAREA.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTAREA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or TXTAREA = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTAREA.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM REFMST WHERE CATA='A'", 0, "", "List Of Area")
        If key_PressNew = True Then
            M_DESC = ""
            Ref_Cat = "A"
            LOAD Frm_Ref_FAS
            Frm_Ref_FAS.Tag = Ref_Cat
            Frm_Ref_FAS.Show
        End If
    End If
End Sub

Private Sub TXTAREA_LostFocus()
   TXTAREA.BackColor = vbWhite
End Sub
Private Sub TXTCST_GotFocus()
  txtcst.BackColor = RGB(BRED, BGREEN, BBLUE)
  txtcst.SelStart = 0
  txtcst.SelLength = Len(txtcst)
  txtcst.ToolTipText = "Enter CST No."
End Sub

Private Sub TXTCST_LostFocus()
  txtcst.BackColor = vbWhite
End Sub

Private Sub txtecc_GotFocus()
  txtecc.BackColor = RGB(BRED, BGREEN, BBLUE)
  txtecc.SelStart = 0
  txtecc.SelLength = Len(txtecc)
  txtecc.ToolTipText = "Enter ECC No."
End Sub

Private Sub txtecc_LostFocus()
 txtecc.BackColor = vbWhite
End Sub

Private Sub txtlst_GotFocus()
  txtlst.BackColor = RGB(BRED, BGREEN, BBLUE)
  txtlst.SelStart = 0
  txtlst.SelLength = Len(txtlst)
  txtlst.ToolTipText = "Enter LST No."
End Sub

Private Sub txtlst_LostFocus()
txtlst.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
  TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTNAME.SelStart = 0
  TXTNAME.SelLength = Len(TXTNAME)
  TXTNAME.ToolTipText = "Enter Party Name"
End Sub


Private Sub txtName_LostFocus()
   TXTNAME.BackColor = vbWhite
End Sub

Private Sub txtPhone_GotFocus()
  txtPhone.BackColor = RGB(BRED, BGREEN, BBLUE)
  txtPhone.SelStart = 0
  txtPhone.SelLength = Len(txtPhone)
  txtPhone.ToolTipText = "Enter Party Address3"
End Sub

Private Sub txtPhone_LostFocus()
  txtPhone.BackColor = vbWhite
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    
    TXTNAME.Enabled = Not bool
    ADD1.Enabled = Not bool
    ADD2.Enabled = Not bool
    ADD3.Enabled = Not bool
    txtecc.Enabled = Not bool
    txtlst.Enabled = Not bool
    txtcst.Enabled = Not bool
End Sub

Private Function entry(CODE As String)
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
Dim ROW As Long

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "Select * from PADDMST where CODE ='" & CODE & "' AND STATUS='A' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If TEMPRS.EOF = False Then
 CMDOK.Caption = "Upd&ate"
 SWITCH = True
 TXTNAME = Trim(TEMPRS!NAME & "")
 ADD1 = Trim(TEMPRS!ADD1 & "")
 ADD2 = Trim(TEMPRS!ADD2 & "")
 ADD3 = Trim(TEMPRS!ADD3 & "")
 txtPhone = Trim(TEMPRS!phone & "")
 txtecc = Trim(TEMPRS!ECCNO & "")
 txtlst = Trim(TEMPRS!LSTNO & "")
 txtcst = Trim(TEMPRS!CSTNO & "")
 
If RS.State = 1 Then RS.Close
RS.Open "SELECT NAME FROM REFMST WHERE CODE='" & Trim(TEMPRS!ARCD & "") & "' AND CATA='A'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   TXTAREA = Trim(RS!NAME & "")
Else
   TXTAREA = Empty
End If
 
 ROWNO = FLEX.Rows - 1
End If

Do While Not TEMPRS.EOF
    ROW = FLEX.Rows - 1
    FLEX.TextMatrix(ROW, 0) = Trim(TEMPRS!ADD1 & "") + " " + Trim(TEMPRS!ADD2 & "") + " " + Trim(TEMPRS!ADD3 & "")
    FLEX.TextMatrix(ROW, 1) = Trim(TEMPRS!phone & "")
    FLEX.TextMatrix(ROW, 2) = Trim(TEMPRS!ECCNO & "")
    FLEX.TextMatrix(ROW, 3) = Trim(TEMPRS!CSTNO & "")
    FLEX.TextMatrix(ROW, 4) = Trim(TEMPRS!LSTNO & "")
    FLEX.TextMatrix(ROW, 5) = Trim(TEMPRS!ADD1 & "")
    FLEX.TextMatrix(ROW, 6) = Trim(TEMPRS!ADD2 & "")
    FLEX.TextMatrix(ROW, 7) = Trim(TEMPRS!ADD3 & "")
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM REFMST WHERE CODE='" & Trim(TEMPRS!ARCD & "") & "' AND CATA='A'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       FLEX.TextMatrix(ROW, 8) = Trim(RS!NAME & "")
    Else
       FLEX.TextMatrix(ROW, 8) = Empty
    End If
    
    FLEX.TextMatrix(ROW, 8) = GetCode("REFMST", Trim(TEMPRS!ARCD & ""), "CODE", "NAME")
    
    FLEX.Rows = FLEX.Rows + 1
    TEMPRS.MoveNext
  Loop
  TEMPRS.Close
  
End Function


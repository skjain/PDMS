VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmOpeningCFormRecievable 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening For Pending C-Form Receivable"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8040
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   31
      Top             =   5400
      Width           =   7575
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   240
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
         Image           =   "frmOpeningCFormRecievable.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   240
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
         Image           =   "frmOpeningCFormRecievable.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   3840
         TabIndex        =   3
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
         Image           =   "frmOpeningCFormRecievable.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1440
         TabIndex        =   1
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
         Image           =   "frmOpeningCFormRecievable.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2640
         TabIndex        =   2
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
         Image           =   "frmOpeningCFormRecievable.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6240
         TabIndex        =   5
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
         Image           =   "frmOpeningCFormRecievable.frx":1CAA
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   240
      TabIndex        =   30
      Top             =   1200
      Width           =   7575
      Begin VB.TextBox TXTQNTY 
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   27
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox TXTCHLN 
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox BRMK 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   29
         Top             =   3480
         Width           =   5895
      End
      Begin VB.ComboBox cmbSaleType 
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         ItemData        =   "frmOpeningCFormRecievable.frx":20FC
         Left            =   1440
         List            =   "frmOpeningCFormRecievable.frx":20FE
         TabIndex        =   7
         Tag             =   "0"
         Text            =   "Select Sale type"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox TXTBNET 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   25
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox txtDCOD 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox TXTBILLNO 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox M_PNAM 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox M_BRNM 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox M_TXNM 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2040
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   1440
         TabIndex        =   19
         Top             =   2400
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   54132737
         CurrentDate     =   39339
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000A&
         Caption         =   "Bill Qnty"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000A&
         Caption         =   "Challan No."
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
         Left            =   3480
         TabIndex        =   22
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks :"
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
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   3480
         Width           =   870
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of  Sale "
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000A&
         Caption         =   "Bill Amount"
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
         TabIndex        =   24
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         TabIndex        =   12
         Tag             =   "S"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee "
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
         TabIndex        =   10
         Tag             =   "S"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "Bill No."
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
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000A&
         Caption         =   "Bill Date"
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
         TabIndex        =   18
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "A/c Party"
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
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "Agent Name"
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
         TabIndex        =   14
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000A&
         Caption         =   "Tax Category"
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
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
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
      Left            =   1440
      TabIndex        =   34
      Top             =   840
      Width           =   6255
   End
   Begin VB.Shape BORDER1 
      BorderColor     =   &H80000002&
      Height          =   300
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   7575
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
      Left            =   360
      TabIndex        =   33
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Opening For Pending C-Form Receivable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Tag             =   "1343"
      Top             =   240
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   6255
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmOpeningCFormRecievable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DIVCODE As String
Dim M_PCOD As String
Dim CPCD As String, ARCD As String, TXGRPCD As String, TTYP As String
Dim M_DBCD As String
Dim M_BRCD As String
Dim M_TXCD As String
Dim M_DCOD As String
Dim M_ADDRESS As String
Dim SAVEFLAG As Boolean

Private Sub BRMK_GotFocus()
  BRMK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub BRMK_LostFocus()
BRMK.BackColor = vbWhite
End Sub

Private Sub cmbSaleType_Click()
  Call SetDBCD
End Sub

Private Sub cmbSaleType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
     KeyCode = 0
  End If
End Sub

Private Sub cmbSaleType_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
   btn_sts (False)
   M_PNAM.SetFocus
   SAVEFLAG = True
End Sub

Private Sub cmdCancel_Click()
    Call RESETALL
    Call btn_sts(True)
    TXTBILLNO.Enabled = True
    TXTBILLNO = Empty
    If cmdAdd.Enabled Then cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Error GoTo LAST
    
    
  SAVEFLAG = False
  TXTBILLNO = Empty
    
  frmEditOpnCform.VTCD = M_DBCD
  frmEditOpnCform.DIVCODE = DIVCODE
  frmEditOpnCform.DIVNAME = LBLDIV
  frmEditOpnCform.Show 1
     
  If TXTBILLNO = Empty Then
     btn_sts (True)
     Call RESETALL
     Call cmdCancel_Click
     cmdAdd.SetFocus
     Exit Sub
  End If
            
    Dim AYS
    
    AYS = MsgBox("Are You Sure To Delete the Data ", vbQuestion + vbYesNo, "Remove This ?")
    
    If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "DELETE FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTBILLNO & "' AND VTYP='OPC'"
                   
        CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) " & _
                   "VALUES('" & compPth & "','OPC','XXXXXXXXXXXXX','" & M_PCOD & "',NULL,'" & TXTBILLNO & _
                   "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
        CN.CommitTrans
    End If
    
    MsgBox "Bill Deleted Successfully"
    
    btn_sts (True)
    Call RESETALL
    Call cmdCancel_Click
    cmdAdd.SetFocus
    Exit Sub
  
  Exit Sub
LAST:
  MsgBox ERR.Description
  CN.RollbackTrans
End Sub

Private Sub cmdEdit_Click()
  
  SAVEFLAG = False
  TXTBILLNO = Empty
    
  frmEditOpnCform.VTCD = M_DBCD
  frmEditOpnCform.DIVCODE = DIVCODE
  frmEditOpnCform.DIVNAME = LBLDIV
  frmEditOpnCform.Show 1
     
  If TXTBILLNO <> Empty Then
     btn_sts (False)
     TXTBILLNO.Enabled = False
     If M_PNAM.Enabled Then M_PNAM.SetFocus
  Else
     Call RESETALL
     btn_sts (True)
     Call cmdCancel_Click
     cmdAdd.SetFocus
  End If

End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

 If CHKSAVEDATA = False Then
    Exit Sub
 End If
  
 Call SAVESAL
  
 If SAVEFLAG = True Then
   MsgBox "Bill Save Successfully"
 Else
   MsgBox "Bill Successfully Edited."
 End If
  
 Call cmdCancel_Click
 Exit Sub
    
LAST:
    MsgBox ERR.Description
    Resume
    If RS.State = 1 Then
        RS.CancelUpdate
    End If
    CN.RollbackTrans
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call CenterChild(frm_Main, Me)
    Call ColorComponent(Me)
    
    M_DESC = Empty: Key = Empty: NEW_VISIBLE = False: DIVCODE = Empty
    LBLDIV.Caption = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & _
                                 "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
        
    Call btn_sts(True)
    SAVEFLAG = True
    Call SetSaleType
    TXTVBDT.Value = FSDT - 1
    TXTVBDT.MaxDate = FSDT - 1
End Sub

Private Sub M_BRNM_GotFocus()
M_BRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_BRNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (Trim(M_BRNM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_BRNM.Text = SearchList1("SELECT CODE, NAME FROM REFMST WHERE CATA='B'", 0, M_BRNM.Text, "SELECT AGENT FROM LIST")
        If key_PressNew = True Then
           M_DESC = ""
           Key = ""
           Ref_Cat = "B"
           M_BRNM.Text = ""
           Frm_Ref_FAS.Show
        Else
           M_BRCD = Key
        End If
    End If
End Sub

Private Sub M_BRNM_LostFocus()
 M_BRNM.BackColor = vbWhite
End Sub

Private Sub M_PNAM_GotFocus()
 M_PNAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_PNAM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(M_PNAM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_PNAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST WHERE DRCR='D'", 0, M_PNAM.Text, "SELECT A/C PARTY")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            M_PNAM.Text = ""
            frm_Acc.Show
        Else
            M_PNAM.Tag = Key
        End If
    End If
    
    Me.KeyPreview = True
End Sub

Private Sub M_PNAM_Change()
   'Call SetPartyHelp
End Sub

Private Sub M_PNAM_LostFocus()
 M_PNAM.BackColor = vbWhite
 
    If SAVEFLAG Then
     Dim GETRS As ADODB.Recordset
     Set GETRS = New ADODB.Recordset
  
     If GETRS.State = 1 Then GETRS.Close
     GETRS.Open "SELECT BRCD,RCOD,TXCD,TTYP FROM ACCMST WHERE NAME='" & M_PNAM & "' ", CN, adOpenDynamic, adLockOptimistic
     If Not GETRS.EOF Then
        M_BRNM = GetCode("REFMST", GETRS!BRCD & "", "CODE", "NAME")
        M_BRCD = Trim(GETRS!BRCD & "")
        M_TXNM = GetCode("TAXMST", GETRS!TXCD & "", "CODE", "NAME")
        M_TXCD = Trim(GETRS!TXCD & "")
        M_DCOD = GetCode("PADDMST", GETRS!RCOD & "", "CODE", "NAME")
        
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT NAME,ADDR FROM PADDMST WHERE CODE='" & Trim(GETRS!RCOD & "") & "' AND SRNO='1'", CN, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
           TXTDCOD = Trim(RS!NAME & "")
           TXTADDRESS = Trim(RS!ADDR & "")
        End If
        RS.Close
     End If
  End If
End Sub

Private Sub M_TXNM_GotFocus()
M_TXNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_TXNM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(M_TXNM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_TXNM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM TAXMST WHERE RECSTAT='A'", 0, M_TXNM.Text, "SELECT TAX FROM LIST")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "T"
            M_TXNM.Text = ""
            FrmSaleTaxMaster.Show
        Else
            M_TXCD = Key
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub M_TXNM_LostFocus()
  M_TXNM.BackColor = vbWhite
End Sub

Private Sub TXTBILLNO_GotFocus()
 TXTBILLNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTBILLNO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TXTBILLNO_LostFocus()
   TXTBILLNO.BackColor = vbWhite
End Sub

Public Sub btn_sts(Yes As Boolean)
    Frame1.Enabled = Not Yes
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
End Sub

Private Sub SetSaleType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND FYCD='" & FYCD & "' AND ACTIVE='Y'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbSaleType.AddItem Trim(PKTYPRS!NAME & "")
 PKTYPRS.MoveNext
Loop

If cmbSaleType.ListCount >= 1 Then cmbSaleType.ListIndex = 0

End Sub



Private Sub TXTBNET_GotFocus()
  TXTBNET.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTBNET_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTBNET, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTBNET_LostFocus()
  TXTBNET.BackColor = vbWhite
End Sub

Private Sub TXTCHLN_GotFocus()
 TXTCHLN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCHLN_LostFocus()
  TXTCHLN.BackColor = vbWhite
End Sub

Private Sub txtDCOD_Change()
  TXTADDRESS = Empty
End Sub

Private Sub TXTDCOD_GotFocus()
TXTDCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDCOD_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTDCOD = Empty
  ElseIf KeyCode = vbKeyF2 Or TXTDCOD = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False
     TXTDCOD = SearchList1("Select DISTINCT CODE,NAME From PADDMST WHERE RECSTAT='A'", 0, Empty, "Select Consinee Name ")
  End If
 Me.KeyPreview = True
End Sub

Private Sub TXTDCOD_LostFocus()
TXTDCOD.BackColor = vbWhite
End Sub

Private Sub TXTADDRESS_GotFocus()
TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   If TXTDCOD = Empty And TXTDCOD.Enabled Then TXTDCOD.SetFocus: Exit Sub
   TXTADDRESS.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Or (TXTADDRESS = Empty And KeyCode = vbKeyReturn) Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & TXTDCOD & "'", 0, Empty, "Select Consignee Address")
   End If
   
   Dim TEMPRS As New ADODB.Recordset
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT * FROM PADDMST WHERE NAME='" & TXTDCOD & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
        M_ADDRESS = TEMPRS!SRNO & ""
        M_DCOD = TEMPRS!CODE & ""
   End If
End Sub

Private Sub TXTADDRESS_LostFocus()
TXTADDRESS.BackColor = vbWhite
End Sub


Private Sub SetDBCD()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND NAME='" & cmbSaleType.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not PKTYPRS.EOF Then
   M_DBCD = Trim(PKTYPRS!CODE & "")
End If

End Sub

Private Sub TXTQNTY_GotFocus()
  TXTCHLN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTQNTY_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTQNTY, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTQNTY_LostFocus()
TXTCHLN.BackColor = vbWhite
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Function CHKSAVEDATA() As Boolean
  CHKSAVEDATA = True
  
  If Trim(cmbSaleType) = Empty Then
     MsgBox "Sale Type required.", vbCritical
     cmbSaleType.Enabled = True
     cmbSaleType.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If Trim(TXTBILLNO) = Empty Then
     MsgBox "Bill No. required.", vbCritical
     TXTBILLNO.Enabled = True
     TXTBILLNO.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If Val(TXTBNET) = 0 Then
     MsgBox "Bill Amount required.", vbCritical
     TXTBNET.Enabled = True
     TXTBNET.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
       
  Dim CHKRS As New ADODB.Recordset
  Set CHKRS = New ADODB.Recordset
  
  'Party  A/c Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * from ACCMST WHERE NAME='" & M_PNAM & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Party Name Not Define ", vbCritical
     M_PNAM.Enabled = True
     M_PNAM.SetFocus
     CHKSAVEDATA = False
     Exit Function
  Else
     M_PCOD = Trim(CHKRS!CODE & "")
     CPCD = Trim(CHKRS!CPCD & "")
     ARCD = Trim(CHKRS!ARCD & "")
     TTYP = Trim(CHKRS!TTYP & "")
  End If
  
  'Agent  A/c Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE from REFMST WHERE NAME='" & M_BRNM & "' AND CATA='B'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Agent Name Not Define ", vbCritical
     M_BRNM.Enabled = True
     M_BRNM.SetFocus
     CHKSAVEDATA = False
     Exit Function
  Else
     M_BRCD = Trim(CHKRS!CODE & "")
  End If
  
  'Consignee Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE,SRNO from PADDMST WHERE NAME='" & TXTDCOD & "' AND ADDR = '" & TXTADDRESS & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Consignee Name Not Define ", vbCritical
     TXTDCOD.Enabled = True
     TXTDCOD.SetFocus
     CHKSAVEDATA = False
     Exit Function
  Else
     M_DCOD = Trim(CHKRS!CODE & "")
     M_ADDRESS = Trim(CHKRS!SRNO & "")
  End If
   
  'Tax Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE,GRPCOD FROM TAXMST WHERE NAME='" & M_TXNM & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Tax Name Not Define ", vbCritical
     M_TXNM.Enabled = True
     M_TXNM.SetFocus
     CHKSAVEDATA = False
     Exit Function
  Else
     M_TXCD = Trim(CHKRS!CODE & "")
     TXGRPCD = Trim(CHKRS!GRPCOD & "")
  End If
    
  If SAVEFLAG Then
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' AND VTYP='OPC' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTBILLNO & "'", CN, adOpenDynamic, adLockOptimistic
  If Not CHKRS.EOF Then
     MsgBox "Duplicate Sale Bill No. !!!! ", vbCritical
     If TXTBILLNO.Enabled Then TXTBILLNO.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  End If
End Function


Private Sub SAVESAL()
On Error GoTo LAST

CN.BeginTrans

If SAVEFLAG Then
   CN.Execute "INSERT INTO BILLMAIN (COMP,UNIT,DVCD,VTYP,DBCD,DATE,VBNO,CHLN,SRNO,SRCH,CRAC,DRAC,PCOD,DCOD," & _
              "ADDRESS,BRCD,CPCD,ARCD,TXCD,TAXGRP,ITOT,BADJ,BNET,TQTY,PYRA,TTYP,[SYSR],[USER],BRMK,BSTS,RECSTAT) " & _
              "VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
              "','OPC','" & M_DBCD & "','" & Format(TXTVBDT, "MM/dd/yyyy") & "','" & TXTBILLNO & "','" & TXTCHLN & "','" & TXTBILLNO & _
              "','1','XXXXXX','" & M_PCOD & "','" & M_PCOD & "','" & M_DCOD & "','" & M_ADDRESS & _
              "','" & M_BRCD & "','" & CPCD & "','" & ARCD & "','" & M_TXCD & _
              "','" & TXGRPCD & "','" & Val(TXTBNET) & "',0,'" & Val(TXTBNET) & _
              "','" & Val(TXTBNET) & "','" & Val(TXTQNTY) & "','" & TTYP & "','N','" & cUName & "','" & BRMK & "','A','A')"
Else
   CN.Execute "UPDATE BILLMAIN SET DATE = '" & Format(TXTVBDT, "MM/dd/yyyy") & _
              "',DRAC='" & M_PCOD & "',PCOD='" & M_PCOD & "',DCOD='" & M_DCOD & _
              "',CHLN='" & TXTCHLN & "',ADDRESS='" & M_ADDRESS & "',BRCD='" & M_BRCD & "',CPCD='" & CPCD & _
              "',ARCD='" & ARCD & "',TXCD='" & M_TXCD & "',TAXGRP='" & TXGRPCD & _
              "',ITOT='" & Val(TXTBNET) & "',BNET='" & Val(TXTBNET) & _
              "',TQTY='" & Val(TXTQNTY) & "',PYRA='" & Val(TXTBNET) & "',TTYP='" & Trim(TTYP) & "',[SYSR]='U',BRMK='" & BRMK & "' WHERE COMP='" & compPth & _
              "' AND UNIT='" & UNCD & _
              "' AND VTYP='OPC' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTBILLNO & "'"
End If

  CN.CommitTrans
  
  Exit Sub
LAST:
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub RESETALL()
    M_PNAM = Empty
    M_TXNM = Empty
    TXTDCOD = Empty
    TXTADDRESS = Empty
    M_BRNM = Empty
    TXTBILLNO = Empty
    TXTBNET = Empty
    BRMK = Empty
    TXTCHLN = Empty
    TXTQNTY = Empty
End Sub

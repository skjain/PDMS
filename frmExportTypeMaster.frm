VERSION 5.00
Begin VB.Form frmExportTypeMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Type Master"
   ClientHeight    =   2835
   ClientLeft      =   3780
   ClientTop       =   3480
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6960
   Begin VB.TextBox totexpvalue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox totexpqty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ComboBox cmbqtyorvalue 
      Height          =   315
      ItemData        =   "frmExportTypeMaster.frx":0000
      Left            =   3840
      List            =   "frmExportTypeMaster.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox M_NAME 
      Height          =   285
      Left            =   1920
      MaxLength       =   150
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   6735
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5520
         TabIndex        =   14
         ToolTipText     =   "Click To Quit From This Module"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4440
         TabIndex        =   12
         ToolTipText     =   "Click to Delete Reference Entry"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3360
         TabIndex        =   10
         ToolTipText     =   "Click To Edit Reference"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   420
         Left            =   2280
         TabIndex        =   8
         ToolTipText     =   "Click to Cancel Editing"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Click to Save Reference"
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Click To Add New Reference"
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Total Export Value"
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
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Total Export Quantity"
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
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Reconsilation On Quantity / Value [Q/V]"
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
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "EXPORT TYPE :"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmExportTypeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Public M_CODE As String

Private Sub cmbqtyorvalue_LostFocus()
  If cmbqtyorvalue.Text = "Quantity" Then
    totexpqty.Enabled = True
    totexpvalue.Enabled = False
   Else
    totexpqty.Enabled = False
    totexpvalue.Enabled = True
  End If
End Sub

Private Sub cmdAdd_Click()
    Call ClsData
    Call btn_sts(False)
    M_NAME.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmbqtyorvalue.Locked = False
    totexpqty.Locked = False
    totexpvalue.Locked = False
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData
    totexpqty.Text = 0
    totexpvalue.Text = 0
End Sub

Private Sub cmdDelete_Click()
  If M_CODE = "" Then
     Exit Sub
  End If
    cmbqtyorvalue.Locked = True
    totexpqty.Locked = True
    totexpvalue.Locked = True
    Dim AYS
    
    AYS = MsgBox("Are You Sure ? Want to Delete This Export Type Master ?", vbYesNo + vbQuestion, "Are You Sure ?")
    
    If AYS = vbYes Then
      CN.BeginTrans
           CN.Execute "UPDATE EXPTYPMST SET RECSTAT='D' WHERE CODE='" & M_CODE & "' AND RECSTAT='A'"
           CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','EXP','XXXXXXXXXXXXX','" & M_NAME & "',NULL,'" & M_CODE & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
       CN.CommitTrans
    End If
    
    Call cmdCancel_Click
    cmdAdd.SetFocus
End Sub

Private Sub cmdEdit_Click()
On Error GoTo errLoadData
  SAVEFLAG = False
  NEW_VISIBLE = False
  Key = Empty
  M_DESC = Empty
  
  M_NAME = SearchList1("select DISTINCT CODE, NAME FROM EXPTYPMST WHERE RECSTAT='A'", 0, "", "List Of EXPORT TYPE MASTER")
  M_CODE = Key
  
  If M_CODE <> Empty Then
     'Call FILLFLEX
  End If
  
  If M_NAME.Enabled = True Then M_NAME.SetFocus

  If M_CODE = Empty Then Exit Sub
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM EXPTYPMST WHERE NAME='" & M_NAME & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    If RS!QORV = "Q" Then
      cmbqtyorvalue.ListIndex = 0
     Else
      cmbqtyorvalue.ListIndex = 1
    End If
  End If
  Call cmbqtyorvalue_LostFocus
  totexpqty.Text = RS!QUANTITY
  totexpvalue.Text = RS!Value
  cmbqtyorvalue.Locked = True
  totexpqty.Locked = True
  totexpvalue.Locked = True
  btn_sts (False)
  M_NAME.SetFocus
  Exit Sub
  
errLoadData:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveRec
    If cmbqtyorvalue.Text = "Quantity" Then
      If Not IsNumeric(totexpqty) Then
        MsgBox "Invalid Quantity"
        totexpqty.SetFocus
        Exit Sub
      End If
     Else
      If Not IsNumeric(totexpvalue) Then
        MsgBox "Invalid Value"
        totexpvalue.SetFocus
        Exit Sub
      End If
    End If
    
    totexpqty = Val(totexpqty)
    totexpvalue = Val(totexpvalue)
    
    
    
    If RS.State = 1 Then RS.Close
    
    
    If Trim(M_NAME) = Empty Then
        MsgBox "Please enter valid Manufacturer Name !!", vbInformation
        M_NAME.SetFocus
        Exit Sub
    End If
   
    
    If SAVEFLAG Then
       M_CODE = GENSIXCOD("SELECT ISNULL(MAX(CODE),'000000') AS CODE FROM EXPTYPMST WHERE CODE LIKE '0%'")
    End If
      
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM EXPTYPMST WHERE NAME='" & M_NAME & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockPessimistic
    If Not RS.EOF Then
        If RS!CODE = M_CODE Then
            'Nothing To Do
        Else
            MsgBox "Duplicate Name For Export Type Master.", vbInformation
            M_NAME.SetFocus
            Exit Sub
        End If
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM EXPTYPMST WHERE CODE='" & M_CODE & "'", CN, adOpenKeyset, adLockPessimistic
    CN.BeginTrans
    If RS.EOF Then
       RS.AddNew
       'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','MFG','XXXXXXXXXXXXX','" & M_NAME & "',NULL,'" & M_CODE & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','N')"
    Else
       'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','MFG','XXXXXXXXXXXXX','" & M_NAME & "',NULL,'" & M_CODE & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','M')"
    End If
        
        RS!CODE = M_CODE
        RS!NAME = M_NAME
        If cmbqtyorvalue.Text = "Quantity" Then
          RS!QORV = "Q"
          RS!QUANTITY = Val(totexpqty.Text)
          RS!Value = 0
         Else
          RS!QORV = "V"
          RS!QUANTITY = 0
          RS!Value = Val(totexpvalue.Text)
        End If
        RS!RECSTAT = "A"
        RS.Update
        
    CN.CommitTrans
    
    Call cmdCancel_Click
    
    cmdAdd.SetFocus

    Exit Sub
    
errSaveRec:
       Resume
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
    On Error Resume Next
    CN.RollbackTrans
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
cmbqtyorvalue.ListIndex = 0
totexpqty.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
     SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
On Error GoTo errLoad
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    M_NAME.Enabled = False
    Call CenterChild(frm_Main, Me)
    cmdExit.Cancel = True
    Me.KeyPreview = True
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    M_NAME.Enabled = Not bool
End Sub

Private Sub ClsData()
    M_NAME.Text = ""
End Sub

Private Sub M_NAME_GotFocus()
M_NAME.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_NAME_LostFocus()
M_NAME.BackColor = vbWhite
End Sub

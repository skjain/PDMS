VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmAddress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7320
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8281
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   8438015
      BackColor       =   12640511
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FramePlusCtl.FramePlus FramePlus2 
         Height          =   3975
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7011
         BackgroundPictureAlignment=   5
         BorderStyle     =   10
         BackColorGradient=   8438015
         BackColor       =   12640511
         Style           =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtRmk3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   6
            Top             =   3480
            Width           =   6735
         End
         Begin VB.TextBox txtRmk2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   5
            Top             =   2880
            Width           =   6735
         End
         Begin VB.TextBox txtRmk1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   4
            Top             =   2280
            Width           =   6735
         End
         Begin VB.TextBox txtAddDCom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   3
            Top             =   1680
            Width           =   6735
         End
         Begin VB.TextBox txtAddSup 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   2
            Top             =   1080
            Width           =   6735
         End
         Begin VB.TextBox txtAddACom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   1
            Top             =   480
            Width           =   6735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Remark3"
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
            TabIndex        =   15
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Remark2"
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
            TabIndex        =   14
            Top             =   2640
            Width           =   2895
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Remark1"
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
            Top             =   2040
            Width           =   6255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Address of Deputy Commissioner of Central Excise"
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
            TabIndex        =   12
            Top             =   1440
            Width           =   6735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Address of Superintendent of Central Excise"
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
            Top             =   840
            Width           =   6375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Address of Asst. Commissioner, Central Excise and Custom"
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
            TabIndex        =   10
            Top             =   240
            Width           =   5775
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   3855
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   7095
         End
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2640
         TabIndex        =   7
         Top             =   4080
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
         Image           =   "frmAddress.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   3840
         TabIndex        =   8
         Top             =   4080
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
         Image           =   "frmAddress.frx":0D8A
         cBack           =   -2147483633
      End
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEXIT_Click()
    Unload Me
End Sub
 
Private Sub cmdSave_Click()
On Error GoTo LAST

    CN.BeginTrans
        CN.Execute "UPDATE UNTMST SET ADD_ACOM='" & txtAddACom.Text & "', ADD_SUP='" & _
                    txtAddSup.Text & "', ADD_DCOM='" & txtAddDCom.Text & "', RMK1='" & _
                    txtRmk1.Text & "', RMK2='" & txtRmk2.Text & "', RMK3='" & txtRmk3.Text & _
                    "' WHERE COMP='" & compPth & "' AND CODE='" & UNCD & "' "
                    
    CN.CommitTrans
    
    MsgBox "DATA SAVED "
    
    Exit Sub
LAST:
    Resume
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  
  Dim TEMPRS As New ADODB.Recordset
  If TEMPRS.State = 1 Then TEMPRS.Close
  TEMPRS.Open "SELECT *FROM UNTMST WHERE COMP='" & compPth & "' AND CODE='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not TEMPRS.EOF Then
      txtAddACom.Text = Trim(TEMPRS!ADD_ACOM & "")
      txtAddSup.Text = Trim(TEMPRS!ADD_SUP & "")
      txtAddDCom.Text = Trim(TEMPRS!ADD_DCOM & "")
      txtRmk1.Text = Trim(TEMPRS!RMK1 & "")
      txtRmk2.Text = Trim(TEMPRS!RMK2 & "")
      txtRmk3.Text = Trim(TEMPRS!RMK3 & "")
  End If

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub txtAddACom_GotFocus()
    txtAddACom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtAddACom_LostFocus()
    txtAddACom.BackColor = vbWhite
End Sub

Private Sub txtAddDCom_GotFocus()
    txtAddDCom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtAddDCom_LostFocus()
    txtAddDCom.BackColor = vbWhite
End Sub

Private Sub txtAddSup_GotFocus()
    txtAddSup.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtAddSup_LostFocus()
    txtAddSup.BackColor = vbWhite
End Sub

Private Sub txtRmk1_GotFocus()
    txtRmk1.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtRmk1_LostFocus()
    txtRmk1.BackColor = vbWhite
End Sub

Private Sub txtRmk2_GotFocus()
    txtRmk2.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtRmk2_LostFocus()
    txtRmk2.BackColor = vbWhite
End Sub

Private Sub txtRmk3_GotFocus()
    txtRmk3.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtRmk3_LostFocus()
    txtRmk3.BackColor = vbWhite
End Sub

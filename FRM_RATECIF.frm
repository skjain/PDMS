VERSION 5.00
Begin VB.Form FRM_RATECIF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculation of CIF Value"
   ClientHeight    =   3195
   ClientLeft      =   6075
   ClientTop       =   4980
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4185
   Begin VB.TextBox ADVN 
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox CIF_VAL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox CIF_RAT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox FRT_VAL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox INS_VAL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox FOB_VAL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox FRT_RAT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox INS_RAT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox FOB_RAT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&O.k"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Advance In USD"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Value"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Unit Price"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "CIF VALUE"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "FREIGHT"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "INSURANCE"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "FOB"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line6 
      X1              =   2880
      X2              =   2880
      Y1              =   120
      Y2              =   2640
   End
   Begin VB.Line Line5 
      X1              =   1680
      X2              =   1680
      Y1              =   120
      Y2              =   2640
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4080
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "FRM_RATECIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CIF_RAT_Change()
 CIF_VAL = Val(CIF_RAT) * Val(frmSale.TXTTQTY)
End Sub

Private Sub cmdOk_Click()
  If IsNumeric(FOB_RAT) And IsNumeric(FOB_INS) And IsNumeric(INS_RAT) Then
    'O.k
   Else
    MsgBox "Invalid Rate "
    FOB_RAT.SetFocus
    Exit Sub
  End If
  If IsNumeric(FOB_VAL) And IsNumeric(FOB_VAL) And IsNumeric(INS_VAL) Then
    'O.k
   Else
    MsgBox "Invalid Value "
    FOB_VAL.SetFocus
    Exit Sub
  End If
  Me.Hide
End Sub

Private Sub FOB_RAT_Change()
  FOB_VAL = Val(FOB_RAT) * Val(frmSale.TXTTQTY)
  CIF_RAT = Val(FOB_RAT) + Val(FRT_RAT) + Val(INS_RAT)
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 FOB_RAT.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub FRT_RAT_Change()
  FRT_VAL = Val(FRT_RAT) * Val(frmSale.TXTTQTY)
  CIF_RAT = Val(FOB_RAT) + Val(FRT_RAT) + Val(INS_RAT)
End Sub

Private Sub INS_RAT_Change()
 INS_VAL = Val(INS_RAT) * Val(frmSale.TXTTQTY)
 CIF_RAT = Val(FOB_RAT) + Val(FRT_RAT) + Val(INS_RAT)
End Sub


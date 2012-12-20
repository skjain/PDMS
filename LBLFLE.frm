VERSION 5.00
Begin VB.Form LBLFLE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lable File"
   ClientHeight    =   4110
   ClientLeft      =   2535
   ClientTop       =   3705
   ClientWidth     =   6885
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6885
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   6615
      Begin VB.TextBox M_NOST 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   300
         Width           =   735
      End
      Begin VB.ComboBox M_FLNM 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton CMDPRV 
         Caption         =   "&Preview"
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "No of Sticker "
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "File Name"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Contants"
      Height          =   2535
      Left            =   4080
      TabIndex        =   19
      Top             =   480
      Width           =   2655
      Begin VB.TextBox M_LIN5 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox M_LIN4 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox M_LIN3 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox M_LIN2 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox M_LIN1 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Line-5"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Line-4"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Line-3"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Line-2"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Line-1"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Formating"
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   3735
      Begin VB.TextBox M_LINS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox M_SPCE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox M_MRGN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Lines"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Space"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Margin"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dimensions"
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   3735
      Begin VB.TextBox M_HGHT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox M_ACRS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox M_WIDT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Height"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Accross"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "COPS LABEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "LBLFLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As New ADODB.Recordset
Dim SQL As String
Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub CMDPRV_Click()
  
  If Not IsNumeric(M_WIDT) Then
    M_WIDT.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_HGHT) Then
    M_HGHT.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_ACRS) Then
    M_ACRS.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_MRGN) Then
    M_MRGN.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_LINS) Then
    M_LINS.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_SPCE) Then
    M_SPCE.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_NOST) Then
    M_NOST.SetFocus
    Exit Sub
  End If
  'Creating Text File for Cops Sticker
  M_LIN1 = Mid(M_LIN1 + Space(50), 1, 50)
  M_LIN2 = Mid(M_LIN2 + Space(50), 1, 50)
  M_LIN3 = Mid(M_LIN3 + Space(50), 1, 50)
  M_LIN4 = Mid(M_LIN4 + Space(50), 1, 50)
  M_LIN5 = Mid(M_LIN5 + Space(50), 1, 50)
  Call PRINTCHAR
  TXTNAM = "DISK" + cUName + "LBL.TXT"
  If Command() = Empty Then
    Open App.PATH & "\REPORTS\" & TXTNAM For Output As #1
   Else
    Open App.PATH & "\FASREPORT\" & TXTNAM For Output As #1
  End If
  Dim arow As Double
  Dim cntr As Double
  Dim prt_row As Double
  Dim prt_acrs As Double
  Dim prt_stg As String
  Dim PRT_ROW_NLM As Double
  arow = Val(M_NOST.Text) / Val(M_ACRS.Text)
  prt_row = 0
  PRT_ROW_NLM = 0
  Do While prt_row < arow
   prt_row = prt_row + 1
   PRT_ROW_NLM = PRT_ROW_NLM + 1
   If M_COMPBILL = "NLM" Then
     If PRT_ROW_NLM > 3 Then
       'Print #1, ""
       'Print #1, ""
       PRT_ROW_NLM = 1
     End If
   End If
   prt_acrs = 0
   prt_stg = "" + Space(Val(M_MRGN.Text))
   Do While prt_acrs < Val(M_ACRS.Text)
    prt_acrs = prt_acrs + 1
    If prt_acrs = Val(M_ACRS.Text) Then
      prt_stg = prt_stg + Mid(M_LIN1.Text, 1, Val(M_WIDT.Text))
     Else
      prt_stg = prt_stg + Mid(M_LIN1.Text, 1, Val(M_WIDT.Text)) + Space(Val(M_SPCE.Text))
    End If
   Loop
   Print #1, prt_stg
   If Val(M_HGHT.Text) >= 2 Then
     prt_acrs = 0
     prt_stg = "" + Space(Val(M_MRGN.Text))
     Do While prt_acrs < Val(M_ACRS.Text)
      prt_acrs = prt_acrs + 1
      If prt_acrs = Val(M_ACRS.Text) Then
        prt_stg = prt_stg + Mid(M_LIN2.Text, 1, Val(M_WIDT.Text))
       Else
        prt_stg = prt_stg + Mid(M_LIN2.Text, 1, Val(M_WIDT.Text)) + Space(Val(M_SPCE.Text))
      End If
     Loop
     Print #1, prt_stg
   End If
   If Val(M_HGHT.Text) >= 3 Then
     prt_acrs = 0
     prt_stg = "" + Space(Val(M_MRGN.Text))
     Do While prt_acrs < Val(M_ACRS.Text)
      prt_acrs = prt_acrs + 1
      If prt_acrs = Val(M_ACRS.Text) Then
        prt_stg = prt_stg + Mid(M_LIN3.Text, 1, Val(M_WIDT.Text))
       Else
        prt_stg = prt_stg + Mid(M_LIN3.Text, 1, Val(M_WIDT.Text)) + Space(Val(M_SPCE.Text))
      End If
     Loop
     Print #1, prt_stg
   End If
   If Val(M_HGHT.Text) >= 4 Then
     prt_acrs = 0
     prt_stg = "" + Space(Val(M_MRGN.Text))
     Do While prt_acrs < Val(M_ACRS.Text)
      prt_acrs = prt_acrs + 1
      If prt_acrs = Val(M_ACRS.Text) Then
        prt_stg = prt_stg + Mid(M_LIN4.Text, 1, Val(M_WIDT.Text))
       Else
        prt_stg = prt_stg + Mid(M_LIN4.Text, 1, Val(M_WIDT.Text)) + Space(Val(M_SPCE.Text))
      End If
     Loop
     Print #1, prt_stg
   End If
   If Val(M_HGHT.Text) >= 5 Then
     prt_acrs = 0
     prt_stg = "" + Space(Val(M_MRGN.Text))
     Do While prt_acrs < Val(M_ACRS.Text)
      prt_acrs = prt_acrs + 1
      If prt_acrs = Val(M_ACRS.Text) Then
        prt_stg = prt_stg + Mid(M_LIN5.Text, 1, Val(M_WIDT.Text))
       Else
        prt_stg = prt_stg + Mid(M_LIN5.Text, 1, Val(M_WIDT.Text)) + Space(Val(M_SPCE.Text))
      End If
     Loop
     Print #1, prt_stg
   End If
   Dim blk_lin As Double
   blk_lin = 0
   Do While blk_lin < Val(M_LINS)
    blk_lin = blk_lin + 1
    Print #1, ""
   Loop
  Loop
  Close #1
  TXTVIWER.Show
End Sub

Private Sub cmdSave_Click()
  If Not IsNumeric(M_WIDT) Then
    M_WIDT.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_HGHT) Then
    M_HGHT.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_ACRS) Then
    M_ACRS.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_MRGN) Then
    M_MRGN.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_LINS) Then
    M_LINS.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_SPCE) Then
    M_SPCE.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(M_NOST) Then
    M_NOST.SetFocus
    Exit Sub
  End If
  If RS.State = adStateOpen Then RS.Close
  SQL = "select * from lblfle where flnm='" & M_FLNM.Text & "'"
  RS.Open SQL, CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
    SQL = "delete from lblfle where flnm='" & M_FLNM.Text & "'"
    CN.BeginTrans
    CN.Execute SQL
    CN.CommitTrans
  End If
  SQL = "insert into lblfle (flnm,widt,hght,acrs,mrgn,lins,spce,lin1,lin2,lin3,lin4,lin5) values ('" & M_FLNM.Text & "','" & M_WIDT.Text & "','" & M_HGHT.Text & "','" & M_ACRS.Text & "','" & M_MRGN.Text & "','" & M_LINS.Text & "','" & M_SPCE.Text & "','" & M_LIN1.Text & "','" & M_LIN2.Text & "','" & M_LIN3.Text & "','" & M_LIN4.Text & "','" & M_LIN5.Text & "')"
  CN.BeginTrans
  CN.Execute SQL
  CN.CommitTrans
End Sub

Private Sub Form_Activate()
    If M_USRSECLEVL = "1" Then
       If ReadConfigReport("0064", 7, "R") = False Then ModuleDeniedMessage_Report: Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
   SendKeys "{Tab}"
 End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  If RS.State = adStateOpen Then RS.Close
  SQL = "select * from lblfle order by FLNM"
  RS.Open SQL, CN, adOpenKeyset, adLockPessimistic
  M_FLNM.Clear
  Do While Not RS.EOF
   M_FLNM.AddItem RS!FLNM
   RS.MoveNext
  Loop
  If M_FLNM.ListCount > 0 Then
    M_FLNM.ListIndex = 0
  End If
End Sub

Private Sub M_ACRS_GotFocus()
M_ACRS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_ACRS_LostFocus()
M_ACRS.BackColor = vbWhite
  If Not IsNumeric(M_ACRS) Then
    M_HGHT.SetFocus
  End If
End Sub

Private Sub M_FLNM_Click()
If RS.State = adStateOpen Then RS.Close
  SQL = "select * from lblfle where flnm='" & M_FLNM.Text & "'"
  RS.Open SQL, CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
    M_WIDT = RS!widt
    M_HGHT = RS!hght
    M_ACRS = RS!acrs
    M_MRGN = RS!MRGN
    M_LINS = RS!lins
    M_SPCE = RS!spce
    M_LIN1 = RS!lin1
    M_LIN2 = RS!lin2
    M_LIN3 = RS!lin3
    M_LIN4 = RS!lin4
    M_LIN5 = RS!lin5
  End If
End Sub

Private Sub M_FLNM_GotFocus()
M_FLNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_FLNM_LostFocus()
M_FLNM.BackColor = vbWhite
If RS.State = adStateOpen Then RS.Close
  SQL = "select * from lblfle where flnm='" & M_FLNM.Text & "'"
  RS.Open SQL, CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
    M_WIDT = RS!widt
    M_HGHT = RS!hght
    M_ACRS = RS!acrs
    M_MRGN = RS!MRGN
    M_LINS = RS!lins
    M_SPCE = RS!spce
    M_LIN1 = RS!lin1
    M_LIN2 = RS!lin2
    M_LIN3 = RS!lin3
    M_LIN4 = RS!lin4
    M_LIN5 = RS!lin5
  End If
End Sub

Private Sub M_HGHT_GotFocus()
M_HGHT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_HGHT_LostFocus()
 M_HGHT.BackColor = vbWhite
 If Not IsNumeric(M_HGHT) Then
    M_HGHT.SetFocus
  End If
End Sub

Private Sub M_LIN1_GotFocus()
M_LIN1.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_LIN1_LostFocus()
M_LIN1.BackColor = vbWhite
End Sub

Private Sub M_LIN2_GotFocus()
M_LIN2.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_LIN2_LostFocus()
M_LIN2.BackColor = vbWhite
End Sub

Private Sub M_LIN3_GotFocus()
M_LIN3.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_LIN3_LostFocus()
M_LIN3.BackColor = vbWhite
End Sub

Private Sub M_LIN4_GotFocus()
M_LIN4.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_LIN4_LostFocus()
M_LIN4.BackColor = vbWhite
End Sub

Private Sub M_LIN5_GotFocus()
M_LIN5.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_LIN5_LostFocus()
M_LIN5.BackColor = vbWhite
End Sub

Private Sub M_LINS_GotFocus()
M_LINS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_LINS_LostFocus()
M_LINS.BackColor = vbWhite
  If Not IsNumeric(M_LINS) Then
    M_MRGN.SetFocus
  End If
End Sub

Private Sub M_MRGN_GotFocus()
M_MRGN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_MRGN_LostFocus()
M_MRGN.BackColor = vbWhite
 If Not IsNumeric(M_MRGN) Then
    M_MRGN.SetFocus
  End If
End Sub

Private Sub M_NOST_GotFocus()
M_NOST.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_NOST_LostFocus()
 M_NOST.BackColor = vbWhite
 If Not IsNumeric(M_NOST) Then
   M_NOST.SetFocus
 End If
End Sub

Private Sub M_SPCE_GotFocus()
M_SPCE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_SPCE_LostFocus()
 M_SPCE.BackColor = vbWhite
 If Not IsNumeric(M_SPCE) Then
    M_SPCE.SetFocus
  End If
End Sub

Private Sub M_WIDT_GotFocus()
M_WIDT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_WIDT_LostFocus()
M_WIDT.BackColor = vbWhite
  If Not IsNumeric(M_WIDT) Then
    M_WIDT.SetFocus
  End If
End Sub

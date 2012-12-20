VERSION 5.00
Begin VB.Form frm_askunit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DIVISION SELECTION"
   ClientHeight    =   3090
   ClientLeft      =   4485
   ClientTop       =   3270
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3675
   Begin VB.TextBox TXTUNIT 
      Height          =   330
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   3600
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   2445
      TabIndex        =   3
      Top             =   2610
      Width           =   1185
   End
   Begin VB.CheckBox chkSelAll 
      Caption         =   "Select &All Division"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   2325
   End
   Begin VB.ListBox LSTNAME 
      Height          =   2085
      Left            =   15
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   375
      Width           =   3600
   End
   Begin VB.ListBox LSTUNIT 
      Height          =   735
      Left            =   135
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1305
      Visible         =   0   'False
      Width           =   3360
   End
End
Attribute VB_Name = "frm_askunit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totSel As Integer
Dim rstemp As Recordset
Private Sub cboCostCenter_Click()
    lstCostCode.ListIndex = cboCostCenter.ListIndex
End Sub
Private Sub Command1_Click()
    Me.Visible = False
End Sub
Private Sub chkSelAll_Click()
    If chkSelAll.Value = 1 Then
        For I = 0 To LSTUNIT.ListCount - 1
            LSTNAME.Selected(I) = True
        Next
    End If
End Sub
Private Sub cmdContinue_Click()
    
    sel_untnam = Empty
    If LSTUNIT.SelCount < 1 Then Exit Sub
    If LSTUNIT.ListCount = 0 Then Unload Me: Exit Sub
    sel_untcod = Empty
    totSel = 0
    Me.Visible = False

    For I = 0 To LSTUNIT.ListCount - 1
         If LSTUNIT.Selected(I) = True Then
            totSel = totSel + 1
            If sel_untcod <> Empty Then sel_untcod = sel_untcod & ","
            If sel_untnam <> Empty Then sel_untnam = sel_untnam & ","
            sel_untcod = sel_untcod & "'" & LSTUNIT.List(I) & "'"
            sel_untnam = sel_untnam & LSTNAME.List(I)
        End If
    Next
    If totSel = LSTUNIT.ListCount And InStr(1, sel_untnam, ",") > 0 Then sel_untnam = "All Unit"
End Sub

Private Sub Form_Load()
    
    Call CenterChild(Screen, Me)
    Me.Caption = "UNIT SELECTION"
    Set rstemp = New Recordset
    
    rstemp.Open "Select Code,Name From UNTMST WHERE COMP='" & compPth & "' Order By NAME", CN, adOpenDynamic
    
    Do While rstemp.EOF = False
        LSTUNIT.AddItem rstemp!code
        LSTNAME.AddItem rstemp!Name
        rstemp.MoveNext
    Loop
    
    If LSTNAME.ListCount > 0 Then LSTNAME.ListIndex = 0
    If LSTNAME.ListCount = 1 Then LSTNAME.Selected(0) = True: Me.Visible = False: CSCD = "'" & LSTUNIT.Text & "'"
    
    rstemp.Close
End Sub

Private Sub lstName_Click()
    LSTUNIT.ListIndex = LSTNAME.ListIndex
End Sub

Private Sub lstName_ItemCheck(Item As Integer)
    If LSTUNIT.ListCount < 1 Then Exit Sub
    LSTUNIT.Selected(LSTNAME.ListIndex) = LSTNAME.Selected(Item)
End Sub

Private Sub txtUNIT_Change()
    
    Set rstemp = New Recordset
    
    rstemp.Open "Select Code,Name From UNTMST WHERE COMP='" & compPth & "' AND Name Like '" & txtUNIT & "%' Order By CODE", CN, adOpenDynamic
    LSTUNIT.Clear
    LSTNAME.Clear
    Do While rstemp.EOF = False
        LSTUNIT.AddItem rstemp!code
        LSTNAME.AddItem rstemp!Name
        rstemp.MoveNext
    Loop
    If LSTNAME.ListCount > 0 Then LSTNAME.ListIndex = 0
    rstemp.Close

End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then LSTNAME.SetFocus
End Sub


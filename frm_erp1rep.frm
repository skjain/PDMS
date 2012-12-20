VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_erp1rep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E.R.1 Register"
   ClientHeight    =   1995
   ClientLeft      =   3120
   ClientTop       =   5295
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6750
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6495
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Pre&view"
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
         Image           =   "frm_erp1rep.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "frm_erp1rep.frx":0452
         cBack           =   -2147483633
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2880
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label13 
         Caption         =   "R&eport Zoom %"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.ComboBox cmbmnt 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ER-1 For the Month of"
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frm_erp1rep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
    
  Dim MM As mMonth
  Dim YY As Integer
  
  Dim RS As New ADODB.Recordset
  Set RS = New ADODB.Recordset
  
  Select Case cmbmnt.ListIndex
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8
    MM = cmbmnt.ListIndex + 4
   Case 9
    MM = 1
   Case 10
    MM = 2
   Case 11
    MM = 3
  End Select
  Select Case MM
    Case 4, 5, 6, 7, 8, 9, 10, 11, 12
      YY = Year(FSDT)
    Case 1, 2, 3
      YY = Year(FEDT)
  End Select
  Dim start_dt As Date
  Dim end_dt As Date
  
  start_dt = GetMinDate(MM, YY)
  end_dt = GetMaxDate(MM, YY)
    
    CRPT.Reset
    crptConnect CRPT
    ReportName = Empty
    rptsql = Empty
    
    ReportName = App.PATH & "\Reports\Form ER-1.rpt"
    RPTN = "Montly Return"
    
    rptsql = "{ER1.COMP}='" & compPth & "' AND {ER1.UNIT}='" & UNCD & "' AND {ER1.DATE}=DATE(" & Year(end_dt) & "," & Month(end_dt) & "," & Day(end_dt) & ")"
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    cmbmnt.SetFocus
    
    CRPT.ReportFileName = ReportName
    
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
    
        .SubreportToChange = ""
        .SubreportToChange = "MODVATDETAIL"
        .Connect = "DSN=" & ServerName & ";UID=sa;PWD= " & DefaultPassword_live & ";DSQ=" & CN.DefaultDatabase
        .SubreportToChange = ""
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True

         If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000084", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If

        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    
    Exit Sub
    
errPreview:
    
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal

End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  cmbmnt.Clear
  cmbmnt.AddItem "April"
  cmbmnt.AddItem "May"
  cmbmnt.AddItem "June"
  cmbmnt.AddItem "July"
  cmbmnt.AddItem "August"
  cmbmnt.AddItem "September"
  cmbmnt.AddItem "October"
  cmbmnt.AddItem "November"
  cmbmnt.AddItem "December"
  cmbmnt.AddItem "January"
  cmbmnt.AddItem "February"
  cmbmnt.AddItem "March"
  cmbmnt.ListIndex = 0
  Call CenterChild(frm_Main, Me)
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_yearendprocess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Y E A R    E N D   P R O C E S S "
   ClientHeight    =   3135
   ClientLeft      =   4080
   ClientTop       =   4185
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6300
   Begin VB.CommandButton CMDCLOSE 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CMDGO 
      Caption         =   "Procced"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker NEWFYST 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54460417
      CurrentDate     =   40473
   End
   Begin MSComCtl2.DTPicker NEWFYEN 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54460417
      CurrentDate     =   40473
   End
   Begin VB.Label Label3 
      Caption         =   "New FY-ENDING DATE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "New FY-STARTING DATE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      Height          =   1695
      Left            =   240
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CURRENT FIN-YEAR DD/MM/YYYY TO DD/MM/YYYY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frm_yearendprocess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub cmdGo_Click()
  Dim NEWFY As String
  Dim SQL As String, SUBSQL As String
  
  NEWFY = Mid((Year(NEWFYST.Value)), 3, 2) + Mid((Year(NEWFYEN.Value)), 3, 2)
  
  STFY = Format(NEWFYST, "YYYY/MM/DD")      'Mid(Year(FSDT), 1, 4)
  ENFY = Format(NEWFYEN, "YYYY/MM/DD")      'Mid(Year(FSDT), 1, 4)
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT FYCD FROM SERIALMASTER WHERE FYCD='" & NEWFY & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    MsgBox "Fy-Year Already exist"
    Exit Sub
  End If
  
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM SERIALMASTER WHERE STFY='" & Format(FSDT, "MM/DD/YYYY") & "' AND ENFY='" & Format(FEDT, "MM/DD/YYYY") & "'", CN, adOpenDynamic, adLockOptimistic
  Do While Not RS.EOF
  
      SQL = "INSERT INTO SERIALMASTER(COMP,UNIT,DVCD,VTYP,CODE,NAME,SRNO,FYCD,STFY,ENFY,PRFX) "
    
      SUBSQL = "VALUES('" & RS!COMP & "','" & RS!unit & "','" & RS!DVCD & "','" & RS!VTYP & _
               "','" & RS!CODE & "','" & RS!NAME & "','000000','" & NEWFY & "','" & STFY & _
               "','" & ENFY & "','" & RS!Prfx & "')"
               
      CN.Execute SQL & SUBSQL
    
  RS.MoveNext
  Loop
  
  CN.Execute "UPDATE SALMANMST SET LSRNO='0000000000'"
  CN.Execute "UPDATE PCKMST SET LBNO='0000000000',LPNO='0000000000'"
     
  MsgBox "Complete. Press Any Key To Restart an Application to Take Effect"
  Shell App.PATH & "\" + App.EXEName + ".exe"
  End
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Label1.Caption = "Current Fynancial Year is " + CStr(FSDT) + " To " + CStr(FEDT)
  NEWFYST.Value = CDate("01/04/" + CStr(Year(FSDT) + 1))
  NEWFYEN.Value = CDate("31/03/" + CStr(Year(FEDT) + 1))
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmVehicleEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Entry"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7035
   Begin VB.Frame framTransDetail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtVHCL 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtTransport 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txtDriver 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2160
         Width           =   4935
      End
      Begin VB.TextBox txtLicenceNo 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2520
         Width           =   4935
      End
      Begin VB.TextBox TXTRMRK 
         Height          =   285
         Left            =   1560
         MaxLength       =   150
         TabIndex        =   7
         Top             =   2880
         Width           =   4935
      End
      Begin MSMask.MaskEdBox txtIN 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker InDate 
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   49217537
         CurrentDate     =   40474
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   3360
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frmVehicleEntry.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4440
         TabIndex        =   9
         Top             =   3360
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frmVehicleEntry.frx":0D8A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5520
         TabIndex        =   10
         Top             =   3360
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frmVehicleEntry.frx":11DC
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Vehicle No."
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
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transport"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&In Time/Date"
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
         Left            =   360
         TabIndex        =   15
         Top             =   1695
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "D&river Name"
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
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Licence No."
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
         Left            =   360
         TabIndex        =   13
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Remar&ks "
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
         Left            =   360
         TabIndex        =   12
         Tag             =   "S"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000D&
         BorderWidth     =   2
         Height          =   3375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Entry Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmVehicleEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim Index As Long
Dim DIVCODE As String
Dim DIVNAME As String
Dim M_DBCD As String
Dim PKG_SCOD As String
Dim BOX_PKG_REQ As String
Dim FICD As String, MCCD As String, LOCCOD As String
Public CHALLAN As String

Private Sub cmdCancel_Click()
    ClsData (Me)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo LAST
    Dim SAVERS As ADODB.Recordset
    Set SAVERS = New ADODB.Recordset
    Dim M_VHCD As String, M_TRCD As String, M_CODE As String
    
    If txtVHCL.Text = Empty Then
        txtVHCL.SetFocus
    End If
    
    If txtTransport.Text = Empty Then
        txtTransport.SetFocus
    End If
    
    'TRANSPORT CODE
    If SAVERS.State = 1 Then SAVERS.Close
    SAVERS.Open "SELECT * FROM TRANSPORTMST WHERE NAME ='" & txtTransport & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
    If Not SAVERS.EOF Then
       M_TRCD = Trim(SAVERS!CODE & "")
    Else
       M_TRCD = Empty
    End If
    SAVERS.Close
    
    'VEHICLE CODE
    If SAVERS.State = 1 Then SAVERS.Close
    SAVERS.Open "SELECT * FROM VHCLMST WHERE NAME ='" & txtVHCL & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
    If Not SAVERS.EOF Then
       M_VHCD = Trim(SAVERS!CODE & "")
    Else
       M_VHCD = Empty
    End If
    SAVERS.Close
    
    If Not IsNumeric(Left(txtIN, 2)) Or Not IsNumeric(Right(txtIN, 2)) Then
        txtIN.SetFocus
        Exit Sub
    End If
    
    If Val(Left(txtIN, 2)) > 23 Or Val(Right(txtIN, 2)) > 59 Then
        txtIN.SetFocus
        Exit Sub
    End If
    
    
    CN.BeginTrans
    CN.Execute "DELETE FROM VHCLENTRY WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DATE='" & Format(InDate.Value, "MM/DD/YYYY") & "' AND VHCD='" & txtVHCL.Tag & "'"
    M_CODE = GENSIXCOD("SELECT ISNULL(MAX(CODE),0) AS CODE FROM VHCLENTRY")
    CN.Execute "INSERT INTO VHCLENTRY (COMP,UNIT,CODE,DATE,VHCD,TRCD,INTIME,DRIVER,LICENCE,REMARKS,OPERATOR) VALUES('" & compPth & "','" & UNCD & "','" & M_CODE & "','" & Format(InDate.Value, "MM/DD/YYYY") & _
                "','" & M_VHCD & "','" & M_TRCD & "','" & txtIN.Text & "','" & txtDriver.Text & "','" & txtLicenceNo.Text & "','" & TXTRMRK.Text & "','" & cUName & "')"
    CN.CommitTrans
    MsgBox "Vehicle Entry Saved Successfully", vbInformation, "Status"
    Call ClsData(Me)
    txtVHCL.SetFocus
                
Exit Sub
LAST:
    MsgBox ERR.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call CenterChild(frm_Main, Me)
    Call ColorComponent(Me)
    On Error GoTo errLoad

    InDate = Now
    InDate.MinDate = FSDT
    InDate.MaxDate = FEDT
    txtIN = Format(Now, "HH:MM")
    
errLoad:
      ErrNumber = ERR.Number
      ErrMessage = ERR.Description
      frm_ErrorHandler.Show vbModal
End Sub

Private Sub InDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtVHCL_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
   
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtVHCL = Empty
  ElseIf KeyCode = vbKeyF2 Or txtVHCL = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = True
     txtVHCL = SearchList1("Select DISTINCT CODE,NAME From VHCLMST WHERE RECSTAT='A'", 0, Empty, "Select Vehicle From List. ")
     
     If key_PressNew Then
        LOAD frmVehicleMaster
     Else
        txtVHCL.Tag = Key
     End If
     
     Call FindDetails
     txtTransport.Tag = GetCode("VHCLMST", txtVHCL.Tag, "CODE", "TRCD")
     txtTransport = GetCode("TRANSPORTMST", txtTransport.Tag, "CODE", "NAME")
  End If
  
 Me.KeyPreview = True
End Sub

Private Sub TXTVHCL_GotFocus(): txtVHCL.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTVHCL_LostFocus(): txtVHCL.BackColor = vbWhite: End Sub
Private Sub txtTransport_GotFocus(): txtTransport.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtTransport_LostFocus(): txtTransport.BackColor = vbWhite: End Sub
Private Sub txtDriver_GotFocus(): txtDriver.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtDriver_LostFocus(): txtDriver.BackColor = vbWhite: End Sub
Private Sub txtLicenceNo_GotFocus(): txtLicenceNo.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtLicenceNo_LostFocus(): txtLicenceNo.BackColor = vbWhite: End Sub
Private Sub TXTRMRK_GotFocus(): TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTRMRK_LostFocus(): TXTRMRK.BackColor = vbWhite: End Sub

Private Sub txtIN_GotFocus()
  txtIN.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtIN_LostFocus(): txtIN.BackColor = vbWhite: End Sub

Private Sub FindDetails()
    If Trim(txtVHCL.Tag) = Empty Then Exit Sub
    
    Dim TMPRS As ADODB.Recordset
    Set TMPRS = New ADODB.Recordset
    If TMPRS.State = 1 Then TMPRS.Close
    TMPRS.Open "SELECT * FROM VHCLENTRY WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VHCD='" & txtVHCL.Tag & "' ORDER BY CODE DESC", CN, adOpenDynamic, adLockOptimistic
    If Not TMPRS.EOF Then
      txtIN = TMPRS!INTIME & ""
      InDate = Format(TMPRS!Date, "DD/MM/YYYY")
      txtDriver = Trim(TMPRS!DRIVER) & ""
      txtLicenceNo = Trim(TMPRS!LICENCE) & ""
      TXTRMRK = Trim(TMPRS!REMARKS) & ""
    Else
        txtIN = Format(Now, "HH:MM")
        txtDriver = ""
        txtLicenceNo = ""
        TXTRMRK = ""
        InDate = Format(Now, "DD/MM/YYYY")
        
    End If
    TMPRS.Close
End Sub

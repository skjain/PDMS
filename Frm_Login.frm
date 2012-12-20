VERSION 5.00
Object = "{8BD302C0-15C7-44FF-8891-BE3F03425023}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Frm_Login 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "User Login"
   ClientHeight    =   2775
   ClientLeft      =   3720
   ClientTop       =   4545
   ClientWidth     =   5355
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "Frm_Login.frx":0000
   Picture         =   "Frm_Login.frx":A6966
   ScaleHeight     =   2775
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   5580
   End
   Begin VB.TextBox txtUName 
      Height          =   330
      Left            =   2400
      TabIndex        =   0
      Text            =   "ADMIN"
      ToolTipText     =   "Enter the User Name"
      Top             =   1080
      Width           =   2505
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "="
      TabIndex        =   1
      Text            =   "ADMIN"
      ToolTipText     =   "Enter the Password"
      Top             =   1440
      Width           =   2505
   End
   Begin VB.ComboBox cboConnectTo 
      Height          =   315
      ItemData        =   "Frm_Login.frx":D71A0
      Left            =   2400
      List            =   "Frm_Login.frx":D71A2
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   2505
   End
   Begin OsenXPCntrl.OsenXPButton cmdOk 
      Height          =   435
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   767
      BTYPE           =   14
      TX              =   "&Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210816
      FCOLO           =   4210816
      MCOL            =   4210816
      MPTR            =   0
      MICON           =   "Frm_Login.frx":D71A4
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Height          =   435
      Left            =   3630
      TabIndex        =   4
      Top             =   1920
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   767
      BTYPE           =   14
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210816
      FCOLO           =   4210816
      MCOL            =   4210816
      MPTR            =   0
      MICON           =   "Frm_Login.frx":D71C0
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(150), Y(150), ptimer 'snow
Dim var As Long

Private Sub cboConnectTo_Change()
    If cboConnectTo.ListIndex = -1 Then cmdOk.Default = False Else cmdOk.Default = True
End Sub

Private Sub cboConnectTo_Click()
    If Trim(txtUName) <> Empty And Trim(txtPass) <> Empty And cboConnectTo.ListIndex <> -1 Then cmdOk.Default = True Else cmdOk.Default = False
End Sub

Private Sub cboConnectTo_GotFocus()
    Msg "Choose Module Connect To !!"
End Sub

Private Sub cboConnectTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboConnectTo_LostFocus()
    Msg ""
End Sub

Private Sub cmdCancel_Click()
    Dim SQL As String

    If MsgBox("Would You Like To Quit !!", vbQuestion + vbYesNo, "Exit From Application ?") = vbYes Then
        IsLoggingOf = True
        CN.Execute "UPDATE USERMAST SET EXTRA1=NULL WHERE UID='" & txtUName & "'"
        CN.Execute "UPDATE USERMAST SET EXTRA1=NULL WHERE UID='" & cUName & "'"
        If txtUName = "ADMIN" Then
          CN.Execute "UPDATE USERMAST SET EXTRA1=NULL"
        End If
        End
    Else
        Exit Sub
    End If

    If NOOFFY > 1 Then
        frm_FYrSelection.Show 1
    Else
UnloadAll:
        For Each LastFrm In Forms
            If LastFrm.NAME <> Me.NAME Then
                Unload LastFrm
            End If
        Next
        
        If IsLoggingOf Then
            Unload frm_Main
            Unload Frm_Login
            Exit Sub
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub cmdCancel_GotFocus()
    Msg "Click To Exit From Program !!"
End Sub

Private Sub CMDOK_Click()
Dim uid As String, PASS As String
    
    If RS.State = adStateOpen Then RS.Close
    
    If InStr(1, txtUName, "'", vbTextCompare) > 0 Or InStr(1, txtPass, "'", vbTextCompare) > 0 Then
        MsgBox "Invalid User Name Or Password", vbInformation, "User / Password Mismatch"
        txtUName.SetFocus
        Exit Sub
    End If
    
    If RS.State = adStateOpen Then RS.Close
    Set RS = New Recordset
    RS.Open "Select * from UserMast where Uid ='" & VBA.Trim(txtUName.Text) & "'", CN, adOpenDynamic, adLockOptimistic
    
    If RS.EOF = True Then
        MsgBox "Invalid Username !! Please Enter Valid User Name", vbInformation, App.TITLE
        txtUName.Text = "": txtPass.Text = "": txtUName.SetFocus
        RS.Close
        Exit Sub
    Else
        If Not IsNull(RS!EXTRA1) Then
          If RS!EXTRA1 = "Y" Then
            'MsgBox "User Already Login with this user", vbCritical
            'End
          End If
        End If
        If RS!USER_ACST = 1 Then
            MsgBox "User Account Is Disabled Or Discontinued !!", vbCritical, "Account Is Disabled"
            Exit Sub
        End If
        
        frm_Main.StsMsg.Panels(4).Text = Me.txtUName.Text
        
        M_USRSECLEVL = RS!USER_LEVEL
        
        If VBA.UCase(txtUName.Text) = VBA.UCase(RS!uid) And txtPass.Text = RS!PASS Then
            'If Not IsNull(RS!Create_Comp) Then If RS!Create_Comp = "n" Then Comp_Cr = "N" Else Comp_Cr = "y" Else Comp_Cr = "n"
            'Comp_Cr = RS!Create_Comp
            RS!EXTRA1 = "Y"
            RS.Update
            Comp_Cr = "Y"
            cUName = txtUName.Text
            
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("0014", 7, "M") = False Then
                    Can_MngUsers = False
                Else
                    Can_MngUsers = True
                End If
            Else
                Can_MngUsers = True
            End If
            
            'If Can_MngUsers = True Then frm_Main.umnuusrmgmt.Visible = True Else frm_Main.umnuusrmgmt.Visible = False
    
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("0015", 4, "M") = False Then
                    Can_CreateComp = False
                Else
                    Can_CreateComp = True
                End If
            End If
            
            If RS.State = 1 Then RS.Close
            

            Call CREATE_LDT
            LOAD Frm_Selection
            Me.Hide
            frm_Main.picInfo.Visible = False
            Unload Me
REASKUNIT:
            Frm_Selection.Show 1
            frm_UnitSelction.Show 1
            
            Call SetMenuVisibility
            
            If UnitFound And UNCD = Empty Then
                GoTo REASKUNIT
            ElseIf UnitFound = False Then
                On Error Resume Next
                Dim Ctrl As Control
                For Each Ctrl In frm_Main
                    If TypeOf Ctrl Is Menu Then
                        Ctrl.Enabled = False
                    End If
                Next
                Exit Sub
            End If
            
            frm_Main.Enabled = True
            frm_Main.picInfo.Visible = True
            frm_Main.lblUnitName = UntNm
            frm_Main.lblCompanycode = compPth
            
            'If Not TipShown Then
            '    If GetSetting(App.EXEName, "Options1", "Show Tips at Startup", 1) <> 0 Then
            '        'LOAD frmTip
            '        'frmTip.Show
            '    End If
            'End If
            
            IsLoggingOf = False
            Unload Me
            Exit Sub
        Else
            MsgBox "Invalid Username or Password!!", vbInformation, App.TITLE
            txtPass.Text = "": txtPass.SetFocus
            RS.Close
        End If
    End If
    
End Sub

Private Sub cmdOk_GotFocus()
    Msg ("LIVEWIRE SOFTWARE LIMITED")
End Sub

Private Sub Form_Activate()
    
    Me.Top = 3000
    Me.Left = Me.Left - 200
    cmdOk.Default = False
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim RS As Object
    
    Screen.MousePointer = vbDefault
    
    If KeyCode = vbKeyEscape Then
        Unload Me
        'Unload Frm_Main
        If NOOFFY > 1 Then
            frm_FYrSelection.Show 1
        Else
            If Trim(M_MEDIA) = "1" Then
                If MsgBox("Want To Remove Data From SQL Catalog ??", vbQuestion + vbYesNo) = vbYes Then
                    'SIMPLY DETACHMENT OF DATABASE AND THEN END PROGRAM
                    CN.Close
                    Set CN = Nothing
                    Set CN = New Connection
                    For Each RS In Me
                        If TypeOf RS Is Recordset Then RS.Close
                    Next
                    CN.ConnectionTimeout = 500
                    CN.Open "Provider=SQLOLEDB.1;Data Source=" & ServerName & ";User Id=sa;PWD= " & DefaultPassword_live & ";Initial Catalog=MASTER"
                    On Error GoTo ErrDetach
                    CN.Execute "sp_detach_db '" & M_DBNM & "','True'"
                    Unload Me
                    End
                Else
                    'SIMPLY EXIT NO DETACHMENT OF DATABASE
                    Unload Me
                    End
                End If
            Else
                Unload Me
                End
            End If
        End If
    End If
    
    Exit Sub

ErrDetach:

If ERR.Number = -2147217865 Then
    MsgBox "Application Failed While Removing Temporary Object !!" & vbCrLf & "Please Close All Application And Then Try Again", vbInformation, ""
    Resume
Else
    Exit Sub
End If

End Sub

Private Sub Form_Load()
'Call ColorComponent(Me)
Dim a As Variant
    a = CenterChild(frm_Main, Me)
   
   For var = 0 To 150
        X(var) = Frm_Login.ScaleHeight * Rnd
        Y(var) = Frm_Login.ScaleWidth * Rnd
        Frm_Login.Circle (X(var), Y(var)), 10, QBColor(15)
    Next var
End Sub

Private Sub txtPass_GotFocus()
txtPass.BackColor = RGB(223, 223, 223)
    Msg ("Enter the Password. Or Press {Esc} To Quit")
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtUName) <> Empty And cboConnectTo.ListIndex <> -1 Then CMDOK_Click
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPass_LostFocus()
txtPass.BackColor = vbWhite
    If Trim(txtUName) <> Empty And Trim(txtPass) <> Empty And cboConnectTo.ListIndex <> -1 Then cmdOk.Default = True Else cmdOk.Default = False
End Sub

Private Sub txtUName_GotFocus()
txtUName.BackColor = RGB(223, 223, 223)
    Msg ("Enter the User Name. Or Press {Esc} To Quit")
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtUName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPass.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtUName_LostFocus()
txtUName.BackColor = vbWhite
    If Trim(txtUName) <> Empty And Trim(txtPass) <> Empty And cboConnectTo.ListIndex <> -1 Then cmdOk.Default = True Else cmdOk.Default = False
End Sub

Public Sub SetMenuVisibility()
    With frm_Main
          .mnuReportOPStkRpt(0).Visible = True
         '.mnuReportOPStkRpt(4).Visible = False
         '.mnuReportOPStkRpt(5).Visible = False
         '.mnuReportOPStkRpt(6).Visible = False
         '.mnuReportOPStkRpt(7).Visible = False
         .mnuOrderBooking(8).Visible = True
         
    End With
End Sub

Private Sub CREATE_LDT()
  Dim SRNO As String
  If RS.State = 1 Then RS.Close
  RS.Open "select * from storeman", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    SRNO = STR(Year(Now)) + STR(Day(Now)) + STR(Month(Now)) + "075"
    CN.BeginTrans
    CN.Execute "insert into storeman (comp,unit,vtyp,srno,srch,dbcd,date,vbno,pcod) values ('0001','000001','LDT','" & Mid(SRNO, 1, 13) & "','01','000001','" & Format(CDate(Now + 75), "MM/DD/YYYY") & "','0000000001','000001')"
    CN.CommitTrans
  End If
End Sub

Private Sub Timer1_Timer()
Dim s As Long
Dim r As Long
If Timer < ptimer + 0.01 Then Exit Sub
Frm_Login.Cls

For var = 0 To 150

s = Int(Rnd * 50)
X(var) = Val(X(var) + s)
r = Int(Rnd * 50)
Y(var) = Val(Y(var) + r)
If X(var) > Frm_Login.ScaleHeight Then
X(var) = -1
End If
If Y(var) > Frm_Login.ScaleWidth Then
Y(var) = -1
End If
Frm_Login.Circle (X(var), Y(var)), 15, QBColor(15)
ptimer = Timer
Next
End Sub

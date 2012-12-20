VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmOrderClearing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Clearing "
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11280
   Begin VB.Frame FrmData 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   11055
      Begin MSComCtl2.DTPicker INVDate 
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   50659329
         CurrentDate     =   40875
      End
      Begin MSComctlLib.ListView lstOrdClr 
         Height          =   3495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2207
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order No."
            Object.Width           =   2207
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agent Name"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Name"
            Object.Width           =   3563
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "BF"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "GSM"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SIZE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Quantity"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "RATE"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "AGRATE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Dispatch Qty"
            Object.Width           =   2645
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Cancel Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Balance Qty"
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Party Order No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "TAX"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "RAT"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "OSRC"
            Object.Width           =   0
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdUpdate 
         Height          =   375
         Left            =   8640
         TabIndex        =   17
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Update"
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
         Image           =   "frmOrderClearing.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdClose 
         Height          =   375
         Left            =   9840
         TabIndex        =   18
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Close"
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
         Image           =   "frmOrderClearing.frx":0D8A
         cBack           =   -2147483633
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellation Date"
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
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   1695
      End
   End
   Begin VB.Frame Frmfilt 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.TextBox TXTPTY 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox TXTGRAD 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   8655
      End
      Begin VB.TextBox txtSM 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   6495
      End
      Begin WelchButton.lvButtons_H CMDSRCH 
         Height          =   375
         Left            =   9720
         TabIndex        =   5
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Search"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmOrderClearing.frx":11DC
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4200
         TabIndex        =   2
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
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
         Format          =   50659329
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
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
         Format          =   50659329
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker INVDT 
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   8640
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
         Format          =   50659329
         CurrentDate     =   40289
      End
      Begin VB.Label Label2 
         Caption         =   "Sales Man   "
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
         Left            =   240
         TabIndex        =   16
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Name of Party"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   11
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Grade"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOrderClearing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ORDBOK As String
Public ORDDBCD As String
Dim STRSQL As String
Dim RAT As Double
Dim lstItem As ListItem
Dim RMK As String
'GLOBAL CONSTANT
Dim M_BRCD As String, SCOMP As String, SUNIT As String, SDVCD  As String, SITM  As String, STAX As String, SBF As String, SGSM As String, SSIZE As String, SUBGRD As String, RATECOD As String, SPARTY As String, SGRD As String
Dim SAVEFLAG As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CMDSRCH_Click()
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
Dim PTYCOD As String
Dim SHDRQ As String
Dim BALQTY As Double
Dim O As Integer
        
    Screen.MousePointer = vbHourglass
    'cmdGo.Default = False
    lstOrdClr.ListItems.Clear
    
        If TEMPRS.State = 1 Then TEMPRS.Close
        STRSQL = "Select ORDMAN.*,ACCMST.NAME,FINITMMST.NAME AS ITEM,REFMST.NAME AS BRNM,  " & _
        "TAXMST.NAME AS TAX,RATEMST.NAME AS RAT from ORDMAN " & _
        "INNER JOIN ACCMST on ORDMAN.PCOD = ACCMST.CODE " & _
        "INNER JOIN FINITMMST on ORDMAN.COMP = FINITMMST.COMP AND ORDMAN.UNIT = FINITMMST.UNIT AND " & _
        "ORDMAN.DCOD = FINITMMST.DVCD AND ORDMAN.ICOD = FINITMMST.CODE " & _
        "INNER JOIN TAXMST on ORDMAN.TXCD = TAXMST.CODE " & _
        "INNER JOIN RATEMST on ORDMAN.RTCD = RATEMST.CODE " & _
        "INNER JOIN REFMST on ORDMAN.BRCD = REFMST.CODE where ORDMAN.DBCD='" & ORDDBCD & _
        "' AND ORDMAN.RECSTAT<>'D' AND ORDMAN.ORDT >= '" & Format(txtFrDate.Value, "MM/dd/YYYY") & _
        "' and ORDMAN.ORDT <= '" & Format(txtToDate.Value, "MM/dd/YYYY") & "' AND ORDMAN.COMP='" & compPth & _
        "' and oflg<>'Y' AND FIN_APRV = 'O' AND (ORDMAN.QNTY - ORDMAN.DOQTY - ORDMAN.DISPATCHQTY - ORDMAN.CANCELQTY) > 0 " 'FOR BALANCE
                
        If Not TXTPTY = Empty Then
          STRSQL = STRSQL & " AND ACCMST.NAME='" & Trim(TXTPTY.Text) & "'"
        End If
        
        STRSQL = STRSQL & " ORDER BY ORDT,ORDMAN.ORDN ASC"
                       
        If TEMPRS.State = 1 Then TEMPRS.Close
        TEMPRS.Open STRSQL, CN, adOpenDynamic, adLockOptimistic
        STRSQL = Empty
        If TEMPRS.EOF = True Then
            MsgBox "There are no Record found.", vbInformation, App.Title
            'cmdOk.Enabled = False
        Else
            Do While Not TEMPRS.EOF
                Set lstItem = lstOrdClr.ListItems.ADD
                lstItem.Text = Format(TEMPRS!ORDT, "dd/MM/yyyy")
                lstItem.Checked = False
                lstItem.SubItems(1) = TEMPRS!ORDN & ""
                lstItem.SubItems(2) = TEMPRS![NAME] & ""
                lstItem.SubItems(3) = TEMPRS![BRNM] & ""
                lstItem.SubItems(4) = TEMPRS!Item & ""
                lstItem.SubItems(17) = TEMPRS!OSRC & ""
        
                
                If Not IsNull(TEMPRS!QNTY) Then lstItem.SubItems(8) = Format(TEMPRS!QNTY, "#####.00") Else lstItem.SubItems(8) = 0
                If Not IsNull(TEMPRS!RATE) Then lstItem.SubItems(9) = Format(TEMPRS!RATE, "#####.00") Else lstItem.SubItems(9) = 0
                
                If Not IsNull(TEMPRS!DISPATCHQTY) Then lstItem.SubItems(11) = Format(TEMPRS!DISPATCHQTY, "#####.00") Else lstItem.SubItems(11) = 0
                If Not IsNull(TEMPRS!CANCELQTY) Then lstItem.SubItems(12) = Format(TEMPRS!CANCELQTY, "#####.00") Else lstItem.SubItems(12) = 0
                
                BALQTY = TEMPRS!QNTY - TEMPRS!DOQTY - TEMPRS!DISPATCHQTY - TEMPRS!CANCELQTY
                
                If Not IsNull(BALQTY) Then lstItem.SubItems(13) = Format(BALQTY, "#####.00") Else lstItem.SubItems(13) = 0
                lstItem.SubItems(14) = TEMPRS!RMRK & ""
                lstItem.SubItems(15) = TEMPRS!TAX
                lstItem.SubItems(16) = TEMPRS!RAT
                
                TEMPRS.MoveNext
            Loop
            'cmdOk.Enabled = True
            'cmdOk.Default = True
            lstOrdClr.SetFocus
        End If
        Screen.MousePointer = vbNormal
    

End Sub

Private Sub cmdupdate_Click()
    On Error GoTo LAST
    Dim I As Long
    Dim L As Long
    Dim ORDN As String
    Dim FLAG As Boolean: FLAG = False
    
    I = 0
    For I = 1 To lstOrdClr.ListItems.COUNT
        If lstOrdClr.ListItems(I).Checked = True Then
           FLAG = True 'CHECKING IF USER SELECT ONE ITEM FROM LIST
           'Exit Sub
           Exit For
        End If
    Next I
    
    'If Not FLAG Then Exit Sub
        
    CN.BeginTrans
        
    I = 0
    For I = 1 To lstOrdClr.ListItems.COUNT
    If lstOrdClr.ListItems(I).Checked = True Then
       Call SetGlobal(I)
              
        Dim INVNO As String
        INVNO = GenDONO(I)
        'TRY
        STRSQL = "INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO," & _
        "QNTY,DELQNTY,RATE,ARAT,ORDN,OSRC,ORDQTY,ORDRATE,ORDDATE,BRMK,PRDL,ICOD,TXRT,TXCD,RTCD,FREIGHT_PERKG," & _
        "FREIGHT_FACTOR,DFLG,DOSTAT,DOAPRVBY,DOAPRVDATE,GRAD) VALUES ('" & SCOMP & _
        "','" & SUNIT & "','" & SDVCD & "','DOS','" & ORDDBCD & "','" & INVNO & "','" & Format(INVDate.Value, "MM/DD/YYYY") & _
        "','" & SPARTY & "','','','" & M_BRCD & "','','" & Val(lstOrdClr.ListItems(I).SubItems(13)) & _
        "','" & Val(lstOrdClr.ListItems(I).SubItems(13)) & "','0','0','" & lstOrdClr.ListItems(I).SubItems(1) & "','1','" & Val(lstOrdClr.ListItems(I).SubItems(8)) & "','" & RAT & _
        "','" & Format(lstOrdClr.ListItems(I).Text, "MM/DD/YYYY") & "','" & RMK & "','','" & SITM & "','','" & STAX & "','" & RATECOD & _
        "','0','0','Y','Y','" & Trim(cUName) & "','" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "','" & SGRD & "')"
        
        CN.Execute STRSQL
        
        STRSQL = "UPDATE ORDMAN SET CANCELQTY = CANCELQTY + " & Val(lstOrdClr.ListItems(I).SubItems(13)) & " WHERE COMP='" & SCOMP & _
        "' AND UNIT='" & SUNIT & "' AND DCOD='" & SDVCD & "' AND DBCD='" & ORDDBCD & _
        "' AND ORDN = '" & lstOrdClr.ListItems(I).SubItems(1) & "' AND ICOD = '" & SITM & "' "
        Dim UPDREC As Double
        
        CN.Execute STRSQL, UPDREC
        
        Call DAILYSTATUS("CNL", SPARTY, ORDDBCD, Val(lstOrdClr.ListItems(I).SubItems(8)), lstOrdClr.ListItems(I).SubItems(1), 0, cUName, "N", Now, lstOrdClr.ListItems(I).Text)
        
    End If
       
    Next
    CN.CommitTrans
    
    MsgBox "Your Order Cancellation No. is : " & INVNO
    
    Call CMDSRCH_Click
    Exit Sub
LAST:
'Resume
     MsgBox ERR.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  
  INVDate.Value = Now
  txtFrDate.Value = GetMinDate
  txtToDate.Value = GetMaxDate
    
  NEW_VISIBLE = False: CANCEL_VISIBLE = False:  M_DESC = Empty:  Key = Empty
  
  '-------SALESMAN MASTER
  ORDBOK = Empty: ORDDBCD = Empty
  ORDBOK = SearchList1("SELECT TOP 20 CODE,NAME FROM SALMANMST", 0, ORDBOK, "SELECT SALESMAN FROM LIST")
  If Key = Empty Then Exit Sub
  ORDDBCD = Key
  txtSM.Text = ORDBOK
  
  'Me.Caption = Me.Caption + " BOOKED BY SALESMAN : " + ORDBOK

End Sub

Private Sub INVDate_GotFocus()
    INVDate.CalendarBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub INVDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub INVDate_LostFocus()
    txtFrDate.CalendarBackColor = vbWhite
End Sub

Private Sub lstOrdClr_GotFocus()
    lstOrdClr.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstOrdClr_LostFocus()
    lstOrdClr.BackColor = vbWhite
End Sub

Private Sub txtFrDate_GotFocus()
    txtFrDate.CalendarBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtToDate.SetFocus
    
End Sub

Private Sub txtFrDate_LostFocus()
    txtFrDate.CalendarBackColor = vbWhite
End Sub

Private Sub TXTPTY_GotFocus()
    TXTPTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPTY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    TXTPTY.Text = SearchList1("SELECT TOP 20 Code,NAME FROM ACCMST", 0, TXTPTY.Text, "SELECT A/C PARTY")
  End If
  If KeyCode = vbKeyDelete Then
    TXTPTY.Text = Empty
  End If
End Sub

Private Sub TXTPTY_LostFocus()
    TXTPTY.BackColor = vbWhite
End Sub

Private Sub txtSM_Change()
    Me.Caption = ""
    Me.Caption = "Order Clearing -" + " BOOKED BY SALESMAN : " + txtSM.Text
End Sub

Private Sub txtSM_GotFocus()
    txtSM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtSM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtSM = SearchList1("Select  TOP 20 Code,Name From SALMANMST", 0, Empty, "Select Sales Man From List")
        ORDDBCD = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtSM = Empty
        ORDDBCD = Empty
    End If
End Sub

Private Function GenDONO(I As Long) As String
    Dim DORS As ADODB.Recordset
    Set DORS = New ADODB.Recordset
    Dim NO As Double
    
    If DORS.State = 1 Then DORS.Close
    DORS.Open "SELECT ISNULL(MAX(RIGHT(DONO,4)),0) AS DONUM FROM ORDTRN WHERE ORDN='" & lstOrdClr.ListItems(I).SubItems(1) & "'", CN, adOpenDynamic, adLockOptimistic
    
    NO = Val(DORS!DONUM)
    NO = NO + 1
    DORS.Close
            
    If NO < 10 Then
       GenDONO = "000" + Trim(nstr(NO, 1, 0))
    ElseIf NO < 100 Then
       GenDONO = "00" + Trim(nstr(NO, 1, 0))
    ElseIf NO < 1000 Then
       GenDONO = "0" + Trim(nstr(NO, 1, 0))
    ElseIf NO < 10000 Then
       GenDONO = Trim(nstr(NO, 1, 0))
    End If
          
       GenDONO = Mid$(CStr(lstOrdClr.ListItems(I).SubItems(1)), 1, 6) & GenDONO

End Function

Private Sub SetGlobal(I As Long)
    Dim GRRS As ADODB.Recordset
    Set GRRS = New ADODB.Recordset
    
    If GRRS.State = 1 Then GRRS.Close
    GRRS.Open "SELECT * FROM ORDMAN WHERE COMP= '" & compPth & "' AND DBCD='" & ORDDBCD & "' AND ORDN ='" & lstOrdClr.ListItems(I).SubItems(1) & "'", CN, adOpenDynamic, adLockOptimistic
    If Not GRRS.EOF Then
       SCOMP = GRRS!COMP
       SUNIT = GRRS!unit
       SDVCD = GRRS!DCOD
       SGRD = Trim(GRRS!TRCD)
       M_BRCD = GetCode("REFMST", lstOrdClr.ListItems(I).SubItems(3), "NAME", "CODE")
       STAX = GetCode("TAXMST", lstOrdClr.ListItems(I).SubItems(15), "NAME", "CODE")
       RATECOD = GetCode("RATEMST", lstOrdClr.ListItems(I).SubItems(16), "NAME", "CODE")
       
       SITM = FindItemCode(lstOrdClr.ListItems(I).SubItems(1), lstOrdClr.ListItems(I).SubItems(4))
              
       
       RAT = GRRS!RATE
       RMK = GRRS!RMRK
    End If
    
    GRRS.Close
    
    
    SPARTY = GetCode("ACCMST", lstOrdClr.ListItems(I).SubItems(2), "NAME", "CODE")

End Sub

Private Function FindItemCode(ORDN As String, INAM As String) As String
    Dim ITRS As ADODB.Recordset
    Set ITRS = New ADODB.Recordset
    Dim GRRS As ADODB.Recordset
    Set GRRS = New ADODB.Recordset
    
    
    If GRRS.State = 1 Then GRRS.Close
    GRRS.Open "SELECT * FROM ORDMAN WHERE DBCD='" & ORDDBCD & "' AND ORDN ='" & ORDN & "'", CN, adOpenDynamic, adLockOptimistic
    If Not GRRS.EOF Then
    
    If ITRS.State = 1 Then ITRS.Close
    ITRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & GRRS!COMP & "' AND UNIT='" & GRRS!unit & "' AND DVCD='" & GRRS!DCOD & "' AND NAME ='" & INAM & "'", CN, adOpenDynamic, adLockOptimistic
    If Not ITRS.EOF Then
       FindItemCode = ITRS!CODE
    Else
       FindItemCode = Empty
    End If
    ITRS.Close
    
    End If
    GRRS.Close

End Function

Private Sub txtSM_LostFocus()
    txtSM.BackColor = vbWhite
End Sub

Private Sub txtToDate_GotFocus()
    txtToDate.CalendarBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtSM.SetFocus
End Sub

Private Sub txtToDate_LostFocus()
    txtToDate.CalendarBackColor = vbWhite
End Sub

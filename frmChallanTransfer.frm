VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "welchbutton.ocx"
Begin VB.Form frmChallanTransfer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Challan Transfer"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10425
   Begin VB.TextBox TXTSTATION 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5880
      Width           =   3585
   End
   Begin MSComCtl2.DTPicker TXTDATE 
      Height          =   350
      Left            =   6480
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      Format          =   87687169
      CurrentDate     =   40413
   End
   Begin VB.TextBox DSTCOMP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":0156
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5160
      Width           =   3105
   End
   Begin VB.TextBox TXTSERVER 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   30
      MouseIcon       =   "frmChallanTransfer.frx":02AC
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4440
      Width           =   1425
   End
   Begin VB.TextBox TXTDATABASE 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      MaxLength       =   30
      MouseIcon       =   "frmChallanTransfer.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4440
      Width           =   2025
   End
   Begin VB.TextBox DSTDBCD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":0558
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5520
      Width           =   3585
   End
   Begin VB.TextBox DSTUNIT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":06AE
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5520
      Width           =   3105
   End
   Begin VB.TextBox DSTDVCD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":0804
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5160
      Width           =   3585
   End
   Begin VB.Frame FramCont 
      Height          =   2505
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   10185
      Begin MSComctlLib.ListView lstBox 
         Height          =   2265
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Challan No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Challan Date."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Del. Party"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pcs"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Quantity"
            Object.Width           =   2892
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "SRNO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox SRCPARTY 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":095A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   720
      Width           =   3705
   End
   Begin VB.TextBox SRCDVCD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":0AB0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   720
      Width           =   3105
   End
   Begin VB.TextBox SRCUNIT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":0C06
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   3105
   End
   Begin VB.TextBox SRCDBCD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      MouseIcon       =   "frmChallanTransfer.frx":0D5C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   360
      Width           =   3705
   End
   Begin WelchButton.lvButtons_H cmdTransfer 
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Transfer"
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
      Image           =   "frmChallanTransfer.frx":0EB2
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdConnected 
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Connect"
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
      Image           =   "frmChallanTransfer.frx":144C
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "frmChallanTransfer.frx":19E6
      cBack           =   -2147483633
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Station"
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
      Left            =   4920
      TabIndex        =   31
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Challan Dt."
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
      Left            =   5040
      TabIndex        =   29
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label LBLCOMP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name: "
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
      TabIndex        =   28
      Top             =   5160
      Width           =   1440
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Destination Database Configuration Details"
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
      TabIndex        =   27
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   120
      Top             =   4200
      Width           =   10215
   End
   Begin VB.Label LBLSERVER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server/PC Name"
      Enabled         =   0   'False
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
      Left            =   360
      TabIndex        =   26
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Label LBLDB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DataBase Name "
      Enabled         =   0   'False
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
      Left            =   3600
      TabIndex        =   25
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Label LBLDAYBOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Type"
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
      Left            =   4920
      TabIndex        =   24
      Top             =   5520
      Width           =   1185
   End
   Begin VB.Label LBLUNIT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Name: "
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
      TabIndex        =   23
      Top             =   5520
      Width           =   1020
   End
   Begin VB.Label LBLDVCD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division Name "
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
      Left            =   4920
      TabIndex        =   22
      Top             =   5160
      Width           =   1290
   End
   Begin VB.Label LBLDST 
      BackColor       =   &H8000000A&
      Caption         =   "Destination Transfer Details"
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
      TabIndex        =   21
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   4920
      Width           =   10215
   End
   Begin VB.Label LBLSRC 
      BackColor       =   &H8000000A&
      Caption         =   "Source Transfer Details"
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
      TabIndex        =   19
      Top             =   0
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name : "
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
      Left            =   5040
      TabIndex        =   18
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division Name "
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
      Left            =   360
      TabIndex        =   17
      Top             =   720
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Name: "
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
      Left            =   360
      TabIndex        =   16
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label lblDayBK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Type: "
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
      Left            =   5040
      TabIndex        =   15
      Top             =   360
      Width           =   1365
   End
End
Attribute VB_Name = "frmChallanTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TCOMP As String
Dim SUNIT As String: Dim SDVCD As String: Dim SDBCD As String: Dim SPARTY As String
Dim TUNIT As String, TDVCD As String, TDBCD As String, TPKGSTCD As String
Dim SRN As String

Private Sub cmdConnected_Click()
On Error GoTo ERR_CONNECTION
 If OpenDESTDB("Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Password='" & DefaultPassword_live & "';Initial Catalog='" & TXTDATABASE & "';Server=" & TXTSERVER & "", "") = False Then
    DSTCOMP.Enabled = False
    DSTUNIT.Enabled = False
    DSTDVCD.Enabled = False
    DSTDBCD.Enabled = False
    Exit Sub
 Else
    DSTCOMP.Enabled = True
    DSTUNIT.Enabled = True
    DSTDVCD.Enabled = True
    DSTDBCD.Enabled = True
 End If
 
 If SRCUNIT <> Empty Then MsgBox "Connected Successfully."
 
 If DSTCOMP <> Empty And SRCUNIT <> Empty Then
    DSTUNIT.Enabled = True: DSTUNIT.SetFocus
 ElseIf SRCUNIT <> Empty Then
    DSTCOMP.Enabled = True: DSTCOMP.SetFocus
 End If
 Exit Sub
ERR_CONNECTION:
MsgBox "Error in Connection. Check LAN/LOCAL Connection."
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdtransfer_Click()
On Error GoTo LAST

'CHECK AUTHORITY OF TRANSFERING
If Not IsUnitAllowTransfer Then
  MsgBox "Access Denied!!!!, Data Can't be Transfered from this unit to another."
  Exit Sub
End If

    Dim MOVEFLAG As Boolean: MOVEFLAG = False
    If Not CHKSAVEDATA Then Exit Sub
        
    'CHECK SELECTION-----------------------------------
    If lstBox.ListItems.COUNT > 0 Then
        Dim I
        For I = 1 To lstBox.ListItems.COUNT
            If lstBox.ListItems(I).Checked Then MOVEFLAG = True: Exit For
        Next I
    End If
    
    If Not MOVEFLAG Then Exit Sub
    '-----------------------------------------------------
    
    'BOXREGISTER------------------------
    
    Dim SQL As String
    Dim nTQTY As Double, nTPCS As Double
                   
    Dim MAINRS As ADODB.Recordset
    Set MAINRS = New ADODB.Recordset
    
    Dim FINISHITEMCODE As String
    
    CN.BeginTrans
        
    For I = 1 To lstBox.ListItems.COUNT
    If lstBox.ListItems(I).Checked Then
      DEST_CN.BeginTrans
      If MAINRS.State = 1 Then MAINRS.Close
             
      MAINRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & _
      "' AND VTYP='DPF' AND RDBC='" & SDBCD & "' AND RVBNO='" & lstBox.ListItems(I) & _
      "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
      
      Do While Not MAINRS.EOF
               
       'FIND SRNO
       If SRN = Empty Then
          SRN = GenTVNO(TUNIT, "DPF", TDBCD)
       End If
                     
       If Trim(lstBox.ListItems(I).SubItems(2)) <> "" Then
          FINISHITEMCODE = GetFinishItemCode(lstBox.ListItems(I).SubItems(2))
       End If
        
       Dim CHKBOX As New ADODB.Recordset
       Set CHKBOX = New ADODB.Recordset
       
       If CHKBOX.State = 1 Then CHKBOX.Close
       CHKBOX.Open "SELECT VBNO FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & TUNIT & _
       "' AND PKG_STCOD ='" & Trim(MAINRS!PKG_STCOD & "") & "' AND VBNO='" & MAINRS!VBNO & "' AND RECSTAT='A'", DEST_CN, adOpenDynamic, adLockOptimistic
       
       If CHKBOX.EOF = True Then
            'Insert data Only when the box does not exist
            
            SQL = "INSERT INTO BOXREGISTER(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,VBDT,CHLN,PKG_STCOD,PKGNG_COD,"
            SQL = SQL & "LOCCOD,PCOD,ISRETURNABLE,LOTNO,ICOD,GRAD,SUBGRD,MCCD,COPS,BOXWGT,COPSWGT,GRSWGT,TRWGT,"
            SQL = SQL & "NTWGT,RMRK,RECSTAT)VALUES('" & compPth & "','" & TUNIT & "','" & TDVCD & "','" & TDBCD & _
            "','PPF','" & MAINRS!VBNO & "','" & Format(Trim(MAINRS!VBDT), "MM/DD/YYYY") & "','" & MAINRS!chln & _
            "','" & TPKGSTCD & "','" & Trim(MAINRS!PKGNG_COD & "") & "','" & Trim(MAINRS!LOCCOD & "") & "','" & Trim(MAINRS!PCOD & "") & _
            "','" & Trim(MAINRS!ISRETURNABLE & "") & "','" & Trim(MAINRS!LOTNO) & "','" & FINISHITEMCODE & "','" & Trim(MAINRS!grad) & _
            "','" & Trim(MAINRS!SUBGRD) & "','000001','" & Val(MAINRS!COPS) & _
            "','" & Val(MAINRS!BOXWGT) & "','" & Val(MAINRS!COPSWGT) & "','" & Val(MAINRS!GRSWGT) & _
            "','" & Val(MAINRS!TRWGT) & "','" & Val(MAINRS!NTWGT) & "','" & Trim(MAINRS!RMRK & "") & "','A')"
                            
            DEST_CN.Execute SQL
       
       Else
            SQL = "UPDATE BOXREGISTER SET CHLN='" & Trim(MAINRS!chln & "") & "',VBDT= '" & Format(Trim(MAINRS!VBDT), "MM/DD/YYYY") & _
            "' WHERE COMP='" & compPth & "' AND UNIT='" & TUNIT & "' AND PKG_STCOD='" & TPKGSTCD & _
            "' AND VBNO='" & Trim(MAINRS!VBNO) & "' AND RECSTAT='A'"
                
            DEST_CN.Execute SQL
       End If
         
              
       If Trim(MAINRS!LOTNO & "") <> "" Then
          Call SetLot(Trim(MAINRS!LOTNO), FINISHITEMCODE)
       End If
       
       If Trim(MAINRS!SUBGRD & "") <> "" And Trim(MAINRS!grad & "") <> "" Then
          Call SetSubGrade(Trim(MAINRS!grad), Trim(MAINRS!SUBGRD))
       End If
       
       If Trim(MAINRS!MCCD & "") <> "" Then
          Call SetMachine(Trim(MAINRS!MCCD & ""))
       End If
       
       'FOR GRADE TRANSFER
       If UCase(TXTSERVER) <> UCase(ServerName) Or UCase(TXTDATABASE) <> UCase(CN.DefaultDatabase) Then
          Call SetGrade(Trim(MAINRS!grad))
       End If
       '-------------------
                        
       MAINRS.MoveNext
       Loop
       MAINRS.Close
       
       SQL = "UPDATE SPTRAN SET TRANSFER='Y' WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & _
             "' AND VTYP='DPF' AND DBCD='" & SDBCD & "' AND VBNO='" & lstBox.ListItems(I) & "' AND RECSTAT<>'D'"
            
       CN.Execute SQL
       End If
        
        '-----------------------------------------------------
    DEST_CN.CommitTrans
    SRN = Empty
    Next I
    
    'Call SetSrcChallan
    CN.CommitTrans
    '---------------------------------------------------------------------
    MsgBox "Transfer Successfully"
    
    Call CLEARDATA
Exit Sub
LAST:
  MsgBox ERR.Description
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtdate_LostFocus()
  Call GetChallan
End Sub

Private Sub DSTCOMP_GotFocus()
  DSTCOMP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DSTCOMP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (DSTCOMP = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   DSTCOMP.Text = SearchListGlobal("SELECT TOP 20 COMP_PATH,COMP_NAME FROM COMPMAST", 0, "", "List Of Transfer Company")
   TCOMP = Key
End If
End Sub

Private Sub DSTCOMP_LostFocus()
  DSTCOMP.BackColor = vbWhite
End Sub

Private Sub DSTDBCD_GotFocus()
  DSTDBCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DSTDBCD_KeyDown(KeyCode As Integer, Shift As Integer)
If DSTCOMP = Empty Then DSTCOMP.Enabled = True: DSTCOMP.SetFocus: Exit Sub
If DSTUNIT = Empty Then DSTUNIT.Enabled = True: DSTUNIT.SetFocus: Exit Sub

If KeyCode = vbKeyF2 Or (DSTDBCD = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   DSTDBCD.Text = SearchListGlobal("SELECT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & _
                  "' AND UNIT='" & TUNIT & "' AND VTYP='PPF' AND FYCD='" & FYCD & "'", 0, "", "List Of Packing Type")
   TDBCD = Key
End If
End Sub

Private Sub DSTDBCD_LostFocus()
  DSTDBCD.BackColor = vbWhite
End Sub

Private Sub DSTDVCD_GotFocus()
  DSTDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DSTDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
If DSTCOMP = Empty Then DSTCOMP.Enabled = True: DSTCOMP.SetFocus: Exit Sub
If DSTUNIT = Empty Then DSTUNIT.Enabled = True: DSTUNIT.SetFocus: Exit Sub
If KeyCode = vbKeyF2 Or (DSTDVCD = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   DSTDVCD.Text = SearchListGlobal("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & TCOMP & "' AND UNIT='" & TUNIT & "' AND RECSTAT<>'D'", 0, "", "List Of Division")
   TDVCD = Key
End If
End Sub

Private Sub DSTDVCD_LostFocus()
  DSTDVCD.BackColor = vbWhite
End Sub

Private Sub DSTUNIT_GotFocus()
  DSTUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DSTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
If DSTCOMP = Empty Then DSTCOMP.Enabled = True: DSTCOMP.SetFocus: Exit Sub
If KeyCode = vbKeyF2 Or (DSTUNIT = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   DSTUNIT.Text = SearchListGlobal("SELECT TOP 20 CODE, NAME FROM UNTMST WHERE COMP='" & TCOMP & "'", 0, "", "List Of Transfer Unit")
   TUNIT = Key
End If
End Sub

Private Sub Form_Activate()
    Call ColorComponent(Me)
    Me.BackColor = RGB(RED, GREEN, BLUE)
    LBLSRC.BackColor = &H8000000A
    LBLDST.BackColor = &H8000000A
    FRMPARA = "SAL"
    Call CHK_UNIT_CONFIG
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If UCase(ActiveControl.NAME) = UCase("lstBox") Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    If ActiveControl.Text = Empty Then Exit Sub
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
On Error GoTo LAST
 Call ColorComponent(Me)
 Me.KeyPreview = True
 txtdate = Now
 Call CenterChild(frm_Main, Me)
 
Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Sub Form_QueryUnload(CANCEL As Integer, UnloadMode As Integer)
   Unload Me
End Sub

Private Sub lstBox_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub SRCDBCD_KeyDown(KeyCode As Integer, Shift As Integer)
If SRCUNIT = Empty Then SRCUNIT.Enabled = True: SRCUNIT.SetFocus: Exit Sub
If SRCDVCD = Empty Then SRCDVCD.Enabled = True: SRCDVCD.SetFocus: Exit Sub

If KeyCode = vbKeyF2 Or (SRCDBCD = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   SRCDBCD.Text = SearchList1("SELECT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & _
   "' AND UNIT='" & SUNIT & "' AND VTYP='DPF' AND CODE NOT IN ('000003','000004') AND " & _
   " NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "'", 0, "", "List Of Day Book")
   SDBCD = Key
End If
If SRCUNIT <> Empty And SRCDVCD <> Empty And SRCDBCD <> Empty And SRCPARTY <> Empty Then Call GetChallan
End Sub

Private Sub SRCDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
If SRCUNIT = Empty Then SRCUNIT.Enabled = True: SRCUNIT.SetFocus: Exit Sub
If KeyCode = vbKeyF2 Or (SRCDVCD = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   SRCDVCD.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & "' AND RECSTAT<>'D'", 0, "", "List Of Division")
   SDVCD = Key
End If
If SRCUNIT <> Empty And SRCDVCD <> Empty And SRCDBCD <> Empty And SRCPARTY <> Empty Then Call GetChallan
End Sub

Private Sub SRCPARTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (SRCPARTY = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   SRCPARTY.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST", 0, "", "List Of A/c Party")
   SPARTY = Key
End If

End Sub

Private Sub SRCUNIT_GotFocus()
  SRCUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SRCUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (SRCUNIT = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   SRCUNIT.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM UNTMST WHERE COMP='" & compPth & "'", 0, "", "List Of Unit")
   SUNIT = Key
End If
End Sub

Private Sub SRCUNIT_LostFocus()
  SRCUNIT.BackColor = vbWhite
End Sub

Private Sub DSTUNIT_LostFocus()
  DSTUNIT.BackColor = vbWhite
End Sub

Private Sub SRCDVCD_LostFocus()
  SRCDVCD.BackColor = vbWhite
End Sub

Private Sub SRCDVCD_GotFocus()
  SRCDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SRCDBCD_LostFocus()
  SRCDBCD.BackColor = vbWhite
End Sub

Private Sub SRCDBCD_GotFocus()
  SRCDBCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SRCPARTY_GotFocus()
  SRCPARTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SRCPARTY_LostFocus()
  SRCPARTY.BackColor = vbWhite
End Sub

Private Sub GetChallan()
On Error GoTo LAST
    Dim SQL As String
    Dim lstItm As ListItem, I As Integer
    
    Dim MAINRS As ADODB.Recordset
    Set MAINRS = New ADODB.Recordset
    
    If SRCDBCD = Empty Then
        MsgBox "Day Book is not selected.", vbCritical, App.Title
        SRCDBCD.SetFocus
        Exit Sub
    End If
    
    lstBox.ListItems.Clear
    
    If MAINRS.State = 1 Then MAINRS.Close
        
    SQL = "SELECT DISTINCT SPTRAN.DBCD,SPTRAN.VBNO,SPTRAN.DATE,SUM(SPTRAN.PCES) AS TPCS,SUM(QNTY) AS TQTY,"
    SQL = SQL & "FINITMMST.NAME AS ITEM,SPTRAN.DCOD From SPTRAN " & _
    "INNER JOIN FINITMMST ON SPTRAN.COMP=FINITMMST.COMP AND SPTRAN.UNIT=FINITMMST.UNIT " & _
    "AND SPTRAN.DVCD=FINITMMST.DVCD AND SPTRAN.ICOD=FINITMMST.CODE "
    SQL = SQL & "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & SUNIT & "' AND SPTRAN.DVCD='" & SDVCD & _
    "' AND SPTRAN.DATE='" & Format(txtdate.Value, "MM/DD/YYYY") & "' AND SPTRAN.DBCD='" & SDBCD & _
    "' AND SPTRAN.PCOD='" & SPARTY & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.RECSTAT<>'D' " & _
    "AND (SPTRAN.TRANSFER IS NULL)"
    
    'DOESN'T COME IN LIST AFTER TRANSFER
        
    SQL = SQL & "GROUP BY SPTRAN.DBCD,SPTRAN.VBNO,SPTRAN.DATE,FINITMMST.NAME,SPTRAN.DCOD "
    SQL = SQL & "ORDER BY SPTRAN.VBNO"
    
    MAINRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    
    If MAINRS.EOF = True Then
        SRCPARTY.SetFocus
        MsgBox "No Record Found for Given Criteria", vbCritical, App.Title
        Exit Sub
    End If
    
    MAINRS.MoveFirst
    Do While Not MAINRS.EOF
        Screen.MousePointer = vbHourglass
        Set lstItm = lstBox.ListItems.ADD()
        lstItm.Text = MAINRS![VBNO]
        lstItm.SubItems(1) = MAINRS![Date]
        lstItm.SubItems(2) = Trim(MAINRS!Item & "")
        lstItm.SubItems(3) = GetCode("PADDMST", Trim(MAINRS![DCOD] & ""), "CODE", "NAME")
        lstItm.SubItems(4) = MAINRS![TPCS]
        lstItm.SubItems(5) = MAINRS![TQTY]
        lstItm.SubItems(6) = MAINRS![dbcd]
        lstItm.Checked = True
        MAINRS.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    Exit Sub
LAST:
  MsgBox ERR.Description
  Resume
End Sub

Private Function CHKSAVEDATA() As Boolean
CHKSAVEDATA = True

If SRCUNIT = Empty Then CHKSAVEDATA = False: SRCUNIT.Enabled = True: SRCUNIT.SetFocus: Exit Function
If SRCDVCD = Empty Then CHKSAVEDATA = False: SRCDVCD.Enabled = True: SRCDVCD.SetFocus: Exit Function
If SRCDBCD = Empty Then CHKSAVEDATA = False: SRCDBCD.Enabled = True: SRCDBCD.SetFocus: Exit Function
If SRCPARTY = Empty Then CHKSAVEDATA = False: SRCPARTY.Enabled = True: SRCPARTY.SetFocus: Exit Function

If DSTUNIT = Empty Then CHKSAVEDATA = False: DSTUNIT.Enabled = True: DSTUNIT.SetFocus: Exit Function
If DSTDVCD = Empty Then CHKSAVEDATA = False: DSTDVCD.Enabled = True: DSTDVCD.SetFocus: Exit Function
If DSTDBCD = Empty Then CHKSAVEDATA = False: DSTDBCD.Enabled = True: DSTDBCD.SetFocus: Exit Function
If TXTSTATION = Empty Then CHKSAVEDATA = False: TXTSTATION.Enabled = True: TXTSTATION.SetFocus: Exit Function

End Function



Private Sub CLEARDATA()
SRCUNIT = Empty
SRCDVCD = Empty
SRCDBCD = Empty
SRCPARTY = Empty

DSTUNIT = Empty
DSTDVCD = Empty
DSTDBCD = Empty
lstBox.ListItems.Clear
SRCUNIT.Enabled = True
SRCUNIT.SetFocus
End Sub

Private Sub SetLot(LOTNO As String, FICD As String)
On Error GoTo LAST
Dim RSLOT As ADODB.Recordset
Set RSLOT = New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM TXULOT WHERE COMP='" & compPth & "' AND DVCD='" & TDVCD & "' AND LTNO='" & Trim(LOTNO) & "' AND UNIT='" & TUNIT & "'", DEST_CN, adOpenDynamic, adLockOptimistic
    If RS.EOF Then 'NEW LOT
       If RSLOT.State = 1 Then RSLOT.Close
       RSLOT.Open "SELECT * FROM TXULOT WHERE COMP='" & compPth & "' AND DVCD='" & SDVCD & "' AND LTNO='" & Trim(LOTNO) & "' AND UNIT='" & SUNIT & "'", CN, adOpenDynamic, adLockOptimistic
       Do While Not RSLOT.EOF  'INSERT
          
          DEST_CN.Execute "INSERT INTO TXULOT (COMP,UNIT,DVCD,LTNO,SRCH,FICD,RICD,PERC,ACTIVE,RECSTAT,SHCD) VALUES ('" & compPth & _
          "','" & TUNIT & "','" & TDVCD & "','" & Trim(LOTNO) & "','" & Trim(RSLOT!SRCH) & "','" & FICD & _
          "','" & Trim(RSLOT!RICD) & "','" & Trim(RSLOT!PERC) & "','" & Trim(RSLOT!ACTIVE) & "','A','" & Trim(RSLOT!SHCD) & "')"
                              
       RSLOT.MoveNext
       Loop
       RSLOT.Close
    End If  'NEW LOT

RS.Close
Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Function GetFinishItemCode(INAM As String) As String
On Error GoTo LAST
Dim FITMCODE As String
Dim RSFINITM As ADODB.Recordset
Set RSFINITM = New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & TUNIT & "' AND NAME='" & Trim(INAM) & "'", DEST_CN, adOpenDynamic, adLockOptimistic
    If RS.EOF Then 'NEW FINISH ITEM
    
       If RSFINITM.State = 1 Then RSFINITM.Close
       RSFINITM.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & "' AND NAME='" & Trim(INAM) & _
       "'", CN, adOpenDynamic, adLockOptimistic
       If Not RSFINITM.EOF Then 'INSERT
          
          FITMCODE = GENICODE
          
          DEST_CN.Execute "INSERT INTO FINITMMST(COMP,UNIT,DVCD,CODE,NAME,DENI,UOM,QORP) VALUES ('" & compPth & _
          "','" & TUNIT & "','" & TDVCD & "','" & FITMCODE & "','" & Trim(RSFINITM!NAME) & "','" & Trim(RSFINITM!DENI) & _
          "','" & Trim(RSFINITM!UOM) & "','" & Trim(RSFINITM!QORP) & "')"
          
          GetFinishItemCode = FITMCODE
                              
       End If
       
    Else
        GetFinishItemCode = RS!CODE & ""
    End If  'NEW FINISH ITEM

RS.Close
Exit Function
LAST:
MsgBox ERR.Description
End Function

Private Sub SetGrade(grad As String)

Dim RSLOT As ADODB.Recordset
Set RSLOT = New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM GRDMST WHERE CODE='" & grad & "'", DEST_CN, adOpenDynamic, adLockOptimistic
    If RS.EOF Then 'NEW GRADE
       If RSLOT.State = 1 Then RSLOT.Close
         RSLOT.Open "SELECT * FROM GRDMST WHERE CODE='" & grad & "'", CN, adOpenDynamic, adLockOptimistic
         If Not RSLOT.EOF Then 'INSERT
          
          DEST_CN.Execute "INSERT INTO GRDMST (CODE,GRAD,SEQC) VALUES ('" & grad & "','" & Trim(RSLOT!grad) & _
                          "','" & Trim(RSLOT!SEQC) & "')"
         End If 'INSERT
       RSLOT.Close
     End If  'NEW GRADE

RS.Close
End Sub

Private Sub SetSubGrade(grad As String, SUBGRAD As String)

Dim RSLOT As ADODB.Recordset
Set RSLOT = New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & TUNIT & "' AND DVCD='" & TDVCD & _
"' AND GRAD='" & grad & "' AND SUBGRD='" & SUBGRAD & "'", DEST_CN, adOpenDynamic, adLockOptimistic
    If RS.EOF Then 'NEW GRADE
       If RSLOT.State = 1 Then RSLOT.Close
         RSLOT.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & _
         "' AND DVCD='" & SDVCD & "' AND GRAD='" & grad & "' AND SUBGRD='" & SUBGRAD & "'", CN, adOpenDynamic, adLockOptimistic
         Do While Not RSLOT.EOF  'INSERT
          
          DEST_CN.Execute "INSERT INTO SUBGRDMST (COMP,UNIT,DVCD,GRAD,SUBGRD,NAME,SWGT,EWGT,RDIFF,SEQNO,STATUS,RECSTAT) VALUES ('" & compPth & _
          "','" & TUNIT & "','" & TDVCD & "','" & grad & "','" & SUBGRAD & "','" & Trim(RSLOT!NAME) & _
          "','" & Trim(RSLOT!SWGT) & "','" & Trim(RSLOT!EWGT) & "','" & Trim(RSLOT!RDIFF) & "','" & Trim(RSLOT!SEQNO) & _
          "','" & Trim(RSLOT!Status) & "','" & Trim(RSLOT!RECSTAT) & "')"
         RSLOT.MoveNext 'INSERT
         Loop
         
       RSLOT.Close
    End If  'NEW GRADE
RS.Close
End Sub


Private Sub SetMachine(MCCD As String)

Dim RSMAC As ADODB.Recordset
Set RSMAC = New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & TUNIT & "' AND DVCD='" & TDVCD & _
"' AND CODE='" & MCCD & "'", DEST_CN, adOpenDynamic, adLockOptimistic
    If RS.EOF Then 'NEW MACHINE
       If RSMAC.State = 1 Then RSMAC.Close
         RSMAC.Open "SELECT * FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & _
         "' AND DVCD='" & SDVCD & "' AND CODE='" & MCCD & "' ", CN, adOpenDynamic, adLockOptimistic
         
         Do While Not RSMAC.EOF  'INSERT
          
          DEST_CN.Execute "INSERT INTO MACMST (COMP,UNIT,DVCD,CODE,NAME,SPDL) VALUES ('" & compPth & _
          "','" & TUNIT & "','" & TDVCD & "','" & MCCD & "','" & Trim(RSMAC!NAME) & "','" & Val(RSMAC!SPDL) & "')"
          
         RSMAC.MoveNext 'INSERT
         Loop
         
       RSMAC.Close
    End If  'NEW GRADE
RS.Close
End Sub


Public Function OpenDESTDB(FILENAME As String, PATH As String) As Boolean
On Error GoTo LAST
    If DEST_CN Is Nothing Then Set DEST_CN = New Connection
    If DEST_CN.State = adStateOpen Then DEST_CN.Close
    With DEST_CN
        .CursorLocation = adUseClient
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .Open FILENAME
    End With
    'compPth = PATH
    OpenDESTDB = True
    
    Exit Function
LAST:
    
    If ERR.Number = -2147467259 Then
        MsgBox "Connect To Server " & ServerName & " Failed...." & vbCrLf & "Server Configuration Required", vbInformation, "Configure Server"
        OpenDESTDB = False
        Exit Function
        If MsgBox("Do You Want To Re Configure Your Server ?", vbYesNo + vbQuestion + vbDefaultButton2, "Re Configure Server?") = vbYes Then
            'If MsgBox("Are You Sure ? ", vbYesNo + vbQuestion + vbDefaultButton2, "Sure ?") = vbYes Then Call ConfigureServer
        Else
            End
        End If
    Else
        MsgBox ERR.Description
        MsgBox "Please Contact Your Software Vendor.", vbInformation, App.Title
        'Resume
        OpenDESTDB = False
        End
    End If
    
End Function


Private Sub TXTDATABASE_GotFocus()
TXTDATABASE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDATABASE_LostFocus()
  TXTDATABASE.BackColor = vbWhite
End Sub

Private Sub txtServer_GotFocus()
TXTSERVER.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSERVER_LostFocus()
TXTSERVER.BackColor = vbWhite
End Sub

Private Sub CHK_UNIT_CONFIG()
  TXTSERVER = ServerName: TXTDATABASE = CN.DefaultDatabase
  DSTCOMP = compNm
  TCOMP = compPth
  DSTCOMP.Enabled = False
  Call cmdConnected_Click
End Sub

Public Function pubGenSrNoPDMS_REMOTE(PDT As Date, TYP As String) As String
'use in delivery challan,box packing, hank despath
Dim Ptrn As String, TEMPRS As New ADODB.Recordset, ctr As Integer
    Ptrn = Format(PDT, "yyyy") + Format(PDT, "mm") + Format(PDT, "dd")
    Ptrn = (Mid(UNCD, 5, 2)) + Format(PDT, "YY") + Format(PDT, "mm") + Format(PDT, "dd") + UCase(Mid(ServerName, 1, 1))
    If TEMPRS.State = 1 Then TEMPRS.Close
    If TYP = Empty Then
        'TEMPRS.Open "Select MAX(SRNO) AS MSRNO from SPMAIN where COMP='" & TCOMP & "' AND SRNO like ('" & Ptrn & "%')", DEST_CN, adOpenDynamic, adLockOptimistic
    Else
        'TEMPRS.Open "Select MAX(SRNO) AS MSRNO from SPMAIN where COMP='" & TCOMP & "' AND SRNO like ('" & Ptrn & "%') AND VTYP='" & TYP & "'", DEST_CN, adOpenDynamic, adLockOptimistic
    End If
    
    If IsNull(TEMPRS!MSRNO) = True Then
        ctr = 1
    Else
        'TEMPRS.MoveLast
        ctr = VBA.Right(TEMPRS!MSRNO, 4)
        ctr = ctr + 1
        Debug.Print ctr
    End If
    If ctr < 10 Then
        pubGenSrNoPDMS_REMOTE = Ptrn + "000" + CStr(ctr)
        Debug.Print pubGenSrNoPDMS_REMOTE
    ElseIf ctr >= 10 And ctr < 100 Then
        pubGenSrNoPDMS_REMOTE = Ptrn + "00" + CStr(ctr)
    ElseIf ctr >= 100 And ctr < 1000 Then
        pubGenSrNoPDMS_REMOTE = Ptrn + "0" + CStr(ctr)
    ElseIf ctr >= 100 And ctr < 1000 Then
        pubGenSrNoPDMS_REMOTE = Ptrn + CStr(ctr)
    Else
        pubGenSrNoPDMS_REMOTE = Ptrn + CStr(ctr)
    End If
End Function

Public Function GenTVNO(m_unit As String, VTYP As String, dbcd As String) As String
Dim NO As Double: NO = 0
Dim GENRS As ADODB.Recordset
Set GENRS = New ADODB.Recordset

If GENRS.State = 1 Then GENRS.Close
GENRS.Open "SELECT LEFT(SRNO,6) AS LVNO FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT = '" & m_unit & _
"' AND VTYP='" & VTYP & "' AND CODE='" & dbcd & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GENRS.EOF Then
   NO = Val(GENRS!LVNO)
   NO = NO + 1
End If
GENRS.Close
   
   If NO < 10 Then
     GenTVNO = "00000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 100 Then
     GenTVNO = "0000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 1000 Then
     GenTVNO = "000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 10000 Then
     GenTVNO = "00" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 100000 Then
     GenTVNO = "0" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 1000000 Then
     GenTVNO = Trim(nstr(NO, 1, 0))
   End If
      
   GenTVNO = GenTVNO & FYCD
End Function


Private Function GENICODE() As String
  
  Dim ITMRS As New ADODB.Recordset
  Set ITMRS = New ADODB.Recordset
           
  If ITMRS.State = 1 Then ITMRS.Close
  ITMRS.Open "Select IsNull(Max(CODE),0) AS CODE From FINITMMST WHERE COMP='" & compPth & _
  "' AND UNIT='" & TUNIT & "'", CN, adOpenDynamic, adLockOptimistic
        
  If Trim(ITMRS!CODE) = "0" Then  'C1
   GENICODE = "0000000001"
  Else
  
   GENICODE = Val(ITMRS!CODE) + 1
   ITMRS.Close
   
   If GENICODE < 10 Then
      GENICODE = "000000000" & GENICODE
   ElseIf GENICODE < 100 Then
      GENICODE = "00000000" & GENICODE
   ElseIf GENICODE < 1000 Then
      GENICODE = "0000000" & GENICODE
   ElseIf GENICODE < 10000 Then
      GENICODE = "000000" & GENICODE
   ElseIf GENICODE < 100000 Then
      GENICODE = "00000" & GENICODE
   ElseIf GENICODE < 1000000 Then
      GENICODE = "0000" & GENICODE
   ElseIf GENICODE < 10000000 Then
      GENICODE = "000" & GENICODE
   ElseIf GENICODE < 100000000 Then
      GENICODE = "00" & GENICODE
   ElseIf GENICODE < 1000000000 Then
     GENICODE = "0" & GENICODE
   Else
      GENICODE = GENICODE
   End If
 End If    'C1
End Function

Private Sub TXTSTATION_GotFocus()
  TXTSTATION.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSTATION_KeyDown(KeyCode As Integer, Shift As Integer)
If DSTCOMP = Empty Then DSTCOMP.Enabled = True: DSTCOMP.SetFocus: Exit Sub
If DSTUNIT = Empty Then DSTUNIT.Enabled = True: DSTUNIT.SetFocus: Exit Sub

If KeyCode = vbKeyF2 Or (TXTSTATION = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False: M_DESC = Empty:  Key = Empty
   TXTSTATION.Text = SearchListGlobal("SELECT CODE,NAME FROM PCKMST WHERE COMP='" & TCOMP & "' AND UNIT='" & TUNIT & "'", 0, "", "List Of Packing Station")
   TPKGSTCD = Key
End If
End Sub

Private Sub TXTSTATION_LostFocus()
TXTSTATION.BackColor = vbWhite
End Sub

Private Function IsUnitAllowTransfer() As Boolean
On Error GoTo LAST
'DEFAULT
IsUnitAllowTransfer = False
'-----------------------

If RS.State = 1 Then RS.Close
RS.Open "SELECT TRANSFER_REQ FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   If Trim(RS!TRANSFER_REQ) = "Y" Then IsUnitAllowTransfer = True
End If
RS.Close

Exit Function
LAST:
IsUnitAllowTransfer = False
MsgBox ERR.Description
CN.RollbackTrans
End Function

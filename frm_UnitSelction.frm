VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8BD302C0-15C7-44FF-8891-BE3F03425023}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form frm_UnitSelction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Selection"
   ClientHeight    =   4590
   ClientLeft      =   1440
   ClientTop       =   2115
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7950
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6CEA8&
      Height          =   3510
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   7965
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   90
         Top             =   2835
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UnitSelction.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UnitSelction.frx":629A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UnitSelction.frx":66EC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstUnit 
         Height          =   3105
         Left            =   135
         TabIndex        =   2
         Top             =   270
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   5477
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   14352123
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Company / Unit"
            Object.Width           =   13583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CompBill"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin OsenXPCntrl.OsenXPButton cmdOk 
      Height          =   435
      Left            =   5160
      TabIndex        =   3
      Top             =   4080
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   767
      BTYPE           =   14
      TX              =   "&O.k"
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
      MICON           =   "frm_UnitSelction.frx":6FC6
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
      Left            =   6480
      TabIndex        =   4
      Top             =   4080
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
      MICON           =   "frm_UnitSelction.frx":6FE2
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Company  Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7965
   End
End
Attribute VB_Name = "frm_UnitSelction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
    Frm_Selection.Show 1
End Sub

Private Sub CMDOK_Click()
On Error GoTo errOkClick
    'Dim RS As New ADODB.Recordset
    'Dim URS As New ADODB.Recordset
   'If cUName <> "ADMIN" Then
   'CN.Execute "UPDATE USERMAST SET EXTRA2 = '" & UNCD & "'  WHERE COMP = '" & compPth & "' AND UID = '" & cUName & "' AND EXTRA2 IS NULL"
   'End If

'If cUName <> "ADMIN" Then
 '   If RS.State = 1 Then RS.Close
 '   RS.Open "SELECT * FROM USERMAST WHERE  COMP = '" & compPth & "' AND UID = '" & cUName & "' AND EXTRA2 = '" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
  '  If RS.EOF Then
 '      MsgBox "User Doesn't have  Authentication For This Unit", vbOKOnly
 '      Exit Sub
 '   End If
'End If

    If LSTUNIT.ListItems.COUNT = 0 Then
        MsgBox "Please Create Division !!", vbInformation
        UNTMST.Show 1
    Else
        Me.Hide
        Call UptoDateDatabase
        UntNm = LSTUNIT.SelectedItem.Text
        frm_Main.lblUnitName.Caption = UntNm
        UNCD = LSTUNIT.SelectedItem.SubItems(1)
        M_COMPBILL = LSTUNIT.SelectedItem.SubItems(2)
        '----------------------------------------------
        'frm_Main.mnuUtilOp(13).Visible = False
        'frm_Main.mnuUtilOp(14).Visible = False
        
        
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT ISNULL(EXTRA5,'N') AS SHOW2NO FROM UNTMST WHERE COMP='" & compPth & "' AND CODE='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
        If RS.EOF = False Then
            If RS!SHOW2NO = "Y" Then
                'frm_Main.mnuUtilOp(13).Visible = True
                'frm_Main.mnuUtilOp(14).Visible = True
            Else
                'frm_Main.mnuUtilOp(13).Visible = False
                'frm_Main.mnuUtilOp(14).Visible = False
            End If
        End If
        '----------------------------------------------
        Dim UNTCFG As New ADODB.Recordset
        Set UNTCFG = New ADODB.Recordset
        If UNTCFG.State = 1 Then UNTCFG.Close
        UNTCFG.Open "SELECT * FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
        M_CUNT = "N"
        UNT_ISPAPPER = "N"
        If Not UNTCFG.EOF Then
           If UNTCFG!CUNT & "" = "" Then M_CUNT = "N" Else M_CUNT = UNTCFG!CUNT
              UNT_TRANSFER_REQ = Trim(UNTCFG!TRANSFER_REQ)
              UNT_ORDER_REQ = Trim(UNTCFG!ORDER_REQ)
              UNT_RATEMST_REQ = Trim(UNTCFG!RATEMST_REQ)
              UNT_MMS_INSTALL = Trim(UNTCFG!MMS_INSTALL)
              UNT_DIVSERIES_REQ = Trim(UNTCFG!DIVSERIES)
              UNT_EXPPKG_REQ = Trim(UNTCFG!EXPPKG_REQ)
              BOXREQ = Trim(UNTCFG!BOXREQ)
              UNT_ISPAPPER = Trim(UNTCFG!ISPAPPER)
              UNT_LRONCHLN = Trim(UNTCFG!LRONCHLN)
              
              IsOnlineChallanPrintReq = IIf(Trim(UNTCFG!ONLINECHLN & "") = "Y", True, False)
              IsOnlineBillPrintReq = IIf(Trim(UNTCFG!ONLINEBILL & "") = "Y", True, False)
                       
              If UNTCFG!EXCUNIT & "" = "Y" Then
                  frm_Main.mnuMasterGrp1(5).Visible = True
                  frm_Main.mnuExcise(0).Visible = True
                  If UCase(cUName) = "ADMIN" Then
                     frm_Main.mnuMiscRepoOp1(14).Visible = True
                  End If
              Else
                  frm_Main.mnuMasterGrp1(5).Visible = False
                  frm_Main.mnuExcise(0).Visible = False
                  frm_Main.mnuMiscRepoOp1(14).Visible = False
              End If
        End If
                
        'CHALLAN TRANSFER
        If UNT_TRANSFER_REQ = "N" Then
           frm_Main.mnuTransOp(17).Visible = False
        End If
        '----------------------------------------
        
        'EXPORT PACKING
        
        If UNT_EXPPKG_REQ = "N" Then
           frm_Main.mnuCarton(2).Visible = False
           frm_Main.mnuCarton(3).Visible = False
           frm_Main.mnuCarton(4).Visible = False
           frm_Main.mnuCarton(5).Visible = False
           frm_Main.mnuPacking(0).Visible = False
        End If
        '----------------------------------------
        
        'MAIN FORM
        Dim i As Long
        Dim VIS_FLAG As Boolean: VIS_FLAG = True
        If UNT_MMS_INSTALL = "Y" Then
           frm_Main.mnuMasterItmOp(5).Visible = False
           VIS_FLAG = False
        End If
        
        If UNT_DIVSERIES_REQ = "Y" Then
           frm_Main.MNUsETUPOP(6).Visible = True
        End If
        
        If BOXREQ = "Y" Then
        frm_Main.mnuOrderBooking(17).Visible = True
        frm_Main.mnuStoreReport(6).Visible = True
        Else
        frm_Main.mnuOrderBooking(17).Visible = False
        frm_Main.mnuStoreReport(6).Visible = False
        End If
        
           
        For i = 7 To 16
          frm_Main.mnuOrderBooking(i).Visible = VIS_FLAG
        Next i
        '======================================================
        
        Unload Me
    End If
    
    Exit Sub
    
errOkClick:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub Form_Activate()
Dim Item As ListItem
    Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    Set rsTemp = New Recordset
        
    rsTemp.Open "Select CODE,NAME,BLNO From UNTMST WHERE COMP='" & compPth & "' AND " & _
                "CODE IN (SELECT DISTINCT UNIT FROM SERIALMASTER WHERE FYCD='" & FYCD & "')  ", CN
    
    LSTUNIT.ListItems.Clear
    
    
    Do While Not rsTemp.EOF
        Set Item = LSTUNIT.ListItems.ADD
        Item = rsTemp!NAME
        Item.SubItems(1) = rsTemp!CODE
        Item.SubItems(2) = rsTemp!BLNO
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    lblCompany.Caption = StrConv(compNm, vbProperCase)
    
    If LSTUNIT.ListItems.COUNT = 0 Then
        IsUnitFound = False
        MsgBox "Setup Is Proceeding to Create Unit !!", vbInformation
        Me.Visible = False
        If Not IsUnitFound Then UNCD = Empty: UnitFound = False: UNTMST.Show: Unload Me: Exit Sub
    Else
        UnitFound = True
        LSTUNIT.SetFocus
        cmdOk.Default = True
    End If
    
End Sub

Private Sub UptoDateDatabase()
On Error Resume Next
Dim SQL As String

CN.Execute "ALTER TABLE BILLMAIN ADD RMCENVAT DECIMAL(18,3) NOT NULL DEFAULT (0)"
CN.Execute "ALTER TABLE BILLMAIN ADD RMEDUCESS DECIMAL(18,3) NOT NULL DEFAULT (0)"
CN.Execute "ALTER TABLE BILLMAIN ADD RMHEDCESS DECIMAL(18,3) NOT NULL DEFAULT (0)"

CN.Execute "ALTER TABLE BOXREGISTER ADD BICOD CHAR(10) NULL"
CN.Execute "ALTER TABLE GRPACKING ADD FRESH DECIMAL(18,3) NOT NULL DEFAULT 0"
CN.Execute "ALTER TABLE GRPACKING ADD WASTAGE DECIMAL(18,3) NOT NULL DEFAULT 0"
CN.Execute "ALTER TABLE GRPACKING ADD RMK CHAR(250) NULL"

CN.Execute "DROP VIEW VWITEM"
CN.Execute "create view VWITEM as select itmmst.code as code,itmmst.name as  name from itmmst inner join igmmst on itmmst.igcd=igmmst.code WHERE IGMMST.MERGE = 'Y'"

CN.Execute "DROP VIEW JOBTRACK_RGP"
CN.Execute "CREATE VIEW JOBTRACK_RGP AS " & _
           "SELECT JOBOUT.COMP,JOBOUT.UNIT,VTYP,DBCD,WONO,WODT,VBNO,'' AS RECNO,SRCH,ACCMST.NAME AS PARTY," & _
           "ITMMST.NAME AS ITEM,JOBOUT.RATE,DATE,'' AS LOTNO,CLRSTATUS,SUM(JOBOUT.QNTY) AS QNTY," & _
           "SUM(JOBOUT.PCES) AS PCS,'' AS CHLN FROM JOBOUT INNER JOIN ACCMST ON JOBOUT.PCOD=ACCMST.CODE " & _
           "INNER JOIN ITMMST ON JOBOUT.ICOD=ITMMST.CODE " & _
           "WHERE RECSTAT<>'D' AND VTYP ='RGP' GROUP BY JOBOUT.COMP,JOBOUT.UNIT,VTYP,DBCD,WONO," & _
           "WODT,VBNO,SRCH,ACCMST.NAME,ITMMST.NAME,JOBOUT.RATE,DATE,CLRSTATUS " & _
           "Union " & _
           "SELECT JOBOUT.COMP,JOBOUT.UNIT,VTYP,DBCD,WONO,WODT,RECNO AS VBNO,VBNO AS RECNO,SRCH,ACCMST.NAME AS PARTY, " & _
           "ITMMST.NAME AS ITEM,JOBOUT.RATE,DATE,LTNO AS LOTNO,CLRSTATUS, " & _
           "SUM(JOBOUT.QNTY) AS QNTY,SUM(JOBOUT.PCES) AS PCS,CHLN FROM JOBOUT " & _
           "INNER JOIN ACCMST ON JOBOUT.PCOD=ACCMST.CODE " & _
           "INNER JOIN ITMMST ON JOBOUT.ICOD=ITMMST.CODE " & _
           "Where RECSTAT<>'D' AND VTYP='IVR' AND DBCD='000003' " & _
           "GROUP BY JOBOUT.COMP,JOBOUT.UNIT,VTYP,DBCD,WONO,WODT,RECNO,VBNO,SRCH,ACCMST.NAME,ITMMST.NAME," & _
           "SRCH,JOBOUT.RATE,DATE,LTNO,CLRSTATUS,CHLN"

If M_COMPBILL = "LKN" Then Exit Sub

    CN.Execute "ALTER TABLE UNTMST ADD EXMNO CHAR(100) NULL"
    CN.Execute "ALTER TABLE UNTCFG ADD ISPAPPER [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTMST ADD TAXDEPT CHAR(150) NULL "
    CN.Execute "ALTER TABLE UNTCFG ADD LRONCHLN [char] (1) NOT NULL DEFAULT 'N'"
    
    CN.Execute "ALTER TABLE UNTMST ADD ADD_ACOM CHAR(250) NOT NULL DEFAULT ''"
    CN.Execute "ALTER TABLE UNTMST ADD ADD_SUP CHAR(250) NOT NULL DEFAULT ''"
    CN.Execute "ALTER TABLE UNTMST ADD ADD_DCOM CHAR(250) NOT NULL DEFAULT ''"
    CN.Execute "ALTER TABLE UNTMST ADD RMK1 CHAR(250) NOT NULL DEFAULT ''"
    CN.Execute "ALTER TABLE UNTMST ADD RMK2 CHAR(250) NOT NULL DEFAULT ''"
    CN.Execute "ALTER TABLE UNTMST ADD RMK3 CHAR(250) NOT NULL DEFAULT ''"
    
    CN.Execute "ALTER TABLE UNTCFG ADD SCANIMP VARCHAR(250) NOT NULL DEFAULT 'C:\'"
    CN.Execute "ALTER TABLE UNTCFG ADD SCANEXP VARCHAR(250) NOT NULL DEFAULT 'C:\'"
    
    CN.Execute "ALTER TABLE UNTCFG ADD ONLINECHLN CHAR(1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD ONLINEBILL CHAR(1) NOT NULL DEFAULT 'N'"
      
    CN.Execute "ALTER TABLE UNTCFG ADD WEXCO CHAR(50) NULL"
    CN.Execute "ALTER TABLE UNTCFG ADD WCHAP CHAR(50) NULL"
    CN.Execute "ALTER TABLE UNTCFG ADD ITEMRO CHAR(1) NOT NULL DEFAULT 'N'"
       
      CN.Execute "ALTER TABLE UNTCFG ADD POSTAC CHAR(1) NOT NULL DEFAULT 'N'"
      CN.Execute "ALTER TABLE UNTCFG ADD POSTACEX CHAR(1) NOT NULL DEFAULT 'N'"
      
      CN.Execute "ALTER TABLE UNTCFG ADD WEXCO CHAR(50) NULL"
      CN.Execute "ALTER TABLE UNTCFG ADD WCHAP CHAR(50) NULL"
      CN.Execute "ALTER TABLE UNTCFG ADD POSTAC CHAR(1) NOT NULL DEFAULT 'N'"
      CN.Execute "ALTER TABLE UNTCFG ADD EXCUNIT CHAR(1) NOT NULL DEFAULT 'N'"
      CN.Execute "ALTER TABLE UNTCFG ADD EXCCPERC DECIMAL(18,3) NOT NULL DEFAULT (0)"
    
    CN.Execute "ALTER TABLE COMPMAST ADD [SEGMENTREQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [JOBCARDREQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [MMS_INSTALL] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [SEGMENT_REQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [RG23D] [char] (10) NOT NULL DEFAULT '00000'"
    CN.Execute "ALTER TABLE UNTCFG ADD [YRNDYGREQ] [char] (1) NOT NULL  DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [TRANSFER_CHLN_TYP] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [FIFOREQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA11] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA12] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA13] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA14] [char] (10) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA15] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA16] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA17] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA18] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA19] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA20] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA21] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA22] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA23] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA24] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA25] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA26] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA27] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA28] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA29] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA30] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA31] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA32] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA33] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA34] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXTRA35] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXCISEAT] [char] (1) NOT NULL DEFAULT '1'"
    
    CN.Execute "ALTER TABLE UNTCFG ADD [BOMREQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [DRNOTERPR] [char] (10) NOT NULL  DEFAULT '00000'"
    CN.Execute "ALTER TABLE UNTCFG ADD [CRNOTERSL] [char] (10) NOT NULL  DEFAULT '00000'"
    CN.Execute "ALTER TABLE UNTCFG ADD [COMMBILL] [char] (10) NOT NULL  DEFAULT '00000'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXPBILL] [char] (10) NOT NULL  DEFAULT '00000'"
    CN.Execute "ALTER TABLE UNTCFG ADD [AMENDMENTNo] [char] (10) NOT NULL DEFAULT '00000'"
    CN.Execute "ALTER TABLE UNTCFG ADD [SEGMENT_REQ1] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [TRASALORD_REQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [TRAPURORD_REQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [EXPPKG_REQ] [char] (1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD [POPERC] [decimal](18, 2) NOT NULL DEFAULT 0"
             
    CN.Execute "ALTER TABLE UNTCFG ADD [POPERC] [decimal](18, 2) NOT NULL DEFAULT 0"
    CN.Execute "ALTER TABLE UNTCFG ADD EXPPKG_REQ CHAR(1) NOT NULL DEFAULT 'N'"
                 
    CN.Execute "ALTER TABLE UNTCFG ADD MMS_INSTALL CHAR(1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD REQ_REQ CHAR(1) NOT NULL DEFAULT 'N'"
    CN.Execute "ALTER TABLE UNTCFG ADD GATE_ENTRY_REQ CHAR(1) NOT NULL DEFAULT 'N'"
    
    CN.Execute "ALTER TABLE UNTMST ADD [DFAD3] [varchar] (50) NULL"
    CN.Execute "ALTER TABLE UNTMST ADD [URL] [varchar] (50) NULL"
    CN.Execute "ALTER TABLE UNTMST ADD [PANO] [varchar] (50) NULL"
    CN.Execute "ALTER TABLE UNTMST ADD [TANO] [varchar] (50) NULL"
    CN.Execute "ALTER TABLE UNTMST ADD [STNO] [varchar] (50) NULL"
        
    CN.Execute "ALTER TABLE COMPMAST ADD [COMP_FNAM] [varchar] (70) NULL "
  
End Sub

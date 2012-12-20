Attribute VB_Name = "modCustomReport"
Sub CustomReportMenu()
On Error GoTo HandelerrMenuLoading
Dim rsTemp As Recordset
Dim LastFilter As String
''Dim BLN_SALESMENUFOUND As Boolean
''Dim BLN_PURMENUFOUND As Boolean
''
''Dim BLN_IVIMENUFOUND As Boolean
''Dim BLN_IVRMENUFOUND As Boolean
''
''Dim BLN_RPRMENUFOUND As Boolean
''Dim BLN_RSLMENUFOUND As Boolean
''
''Dim BLN_OSSMENUFOUND As Boolean
Dim BLN_MENUFOUND As Boolean


    Set rsTemp = New Recordset
    
    
    rsTemp.Open "Select * From REPCNF Order By RPCD", CN, adOpenDynamic
    BLN_MENUFOUND = False
    Do While Not rsTemp.EOF
        'If Left(rsTemp!RPCD, 3) = "PUR" And frm_Main.CSMREP(1).Visible = False Then frm_Main.CSMREP(1).Visible = True: frm_Main.mnuCustomReports.Visible = True: BLN_MENUFOUND = True
        'If Left(rsTemp!RPCD, 3) = "RSL" And frm_Main.CSMREP(2).Visible = False Then frm_Main.CSMREP(2).Visible = True: frm_Main.mnuCustomReports.Visible = True: BLN_MENUFOUND = True
        'If Left(rsTemp!RPCD, 3) = "RPR" And frm_Main.CSMREP(3).Visible = False Then frm_Main.CSMREP(3).Visible = True: frm_Main.mnuCustomReports.Visible = True: BLN_MENUFOUND = True
        'If Left(rsTemp!RPCD, 3) = "PSR" And frm_Main.CSMREP(4).Visible = False Then frm_Main.CSMREP(4).Visible = True: frm_Main.mnuCustomReports.Visible = True: BLN_MENUFOUND = True
        'If Left(rsTemp!RPCD, 3) = "IVR" And frm_Main.CSMREP(5).Visible = False Then frm_Main.CSMREP(5).Visible = True: frm_Main.mnuCustomReports.Visible = True: BLN_MENUFOUND = True
        If Left(rsTemp!RPCD, 3) = "SAL" Then
            frm_Main.CSMREP(0).Visible = True: frm_Main.mnuCustomReports.Visible = True
        Else
            If BLN_MENUFOUND = True Then
                frm_Main.CSMREP(0).Visible = False
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    
    Exit Sub
    
HandelerrMenuLoading:
    
    MsgBox Err.Number & Err.Description
End Sub


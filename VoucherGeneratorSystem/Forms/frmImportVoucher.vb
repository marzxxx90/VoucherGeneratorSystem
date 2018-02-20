Imports Microsoft.Office.Interop

Public Class frmImportVoucher

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        If txtPath.Text = "" Then Exit Sub
        If cboMins.Text = "" Then Exit Sub

        'Load Excel
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oWB = oXL.Workbooks.Open(txtPath.Text)
        oSheet = oWB.Worksheets(1)

        Dim MaxColumn As Integer = oSheet.Cells(1, oSheet.Columns.Count).End(Excel.XlDirection.xlToLeft).column
        Dim MaxEntries As Integer = oSheet.Cells(oSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).row

        Dim checkHeaders(MaxColumn) As String
        For cnt As Integer = 0 To MaxColumn - 1
            checkHeaders(cnt) = oSheet.Cells(1, cnt + 1).value
        Next : checkHeaders(MaxColumn) = oWB.Worksheets(1).name

        'If Not TemplateIntegrityCheck(checkHeaders) Then
        '    AddTimelyLogs("IMPORT MASTER DATA", "Template was tampered", , False, "IMD Template has been modify", )
        '    MsgBox("Template was tampered", MsgBoxStyle.Critical, "Please Contact Warehouse")
        '    GoTo unloadObj
        'End If

        Disable(1)
        pbStatus.Maximum = MaxEntries
        For cnt = 8 To MaxEntries
            pbStatus.Value = cnt
            Dim tmpCode As New VoucherClass
            With tmpCode
                If .isVoucherExist(oSheet.Cells(cnt, 1).Value, cboMins.Text) = True Then Continue For
                ._vCode = oSheet.Cells(cnt, 1).Value
                ._status = 1
                ._mins_Time = cboMins.Text
                .SaveVoucher()
            End With

            Application.DoEvents()
            Console.WriteLine("Voucher Code " & oSheet.Cells(cnt, 1).Value)
        Next
        Disable(0)

unloadObj:
        'Memory Unload
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing

        MsgBox("Voucher Code Imported!", MsgBoxStyle.Information, "Voucher Generator System")
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ofdImport.ShowDialog()
        txtPath.Text = ofdImport.FileName
    End Sub

    Private Sub frmImportVoucher_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mysql As String = "Select DISTINCT(Mins_Time) From tblVoucher"
        Dim ds As DataSet = LoadSQL(mysql, "tblVoucher")

        cboMins.Items.Clear()
        For Each dr In ds.Tables(0).Rows
            cboMins.Items.Add(dr.item("Mins_Time"))
        Next
    End Sub

    Private Sub cboMins_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboMins.KeyPress
        DigitOnly(e)
    End Sub

    Private Sub Disable(ByVal st As Boolean)
        btnImport.Enabled = Not st
        cboMins.Enabled = Not st
    End Sub
End Class
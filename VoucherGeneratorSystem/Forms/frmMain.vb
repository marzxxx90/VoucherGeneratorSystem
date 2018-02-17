Imports Microsoft.Reporting.WinForms
Imports System.Drawing.Printing
Imports System.Drawing.Imaging
Imports System.IO

Public Class frmMain

    Private Sub ImportVoucherToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportVoucherToolStripMenuItem.Click
        frmImportVoucher.Show()
    End Sub

    Private Sub ClearVoucherToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearVoucherToolStripMenuItem.Click
        Dim ans As DialogResult = _
           MsgBox("Are you sure you want to clear?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "Voucher Generator System")
        If ans = Windows.Forms.DialogResult.No Then Exit Sub

        Dim mysql As String = "Delete From tblVoucher"
        Dim CLearVoucher As String = "SET GENERATOR TBLVOUCHER_ID_GEN TO 0;"

        SQLCommand(mysql)
        SQLCommand(CLearVoucher)

        MsgBox("Voucher Cleared!", MsgBoxStyle.Information, "Voucher Generator System")
    End Sub

    Private Sub tCurrent_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tCurrent.Tick
        lblTime.Text = Date.Now
    End Sub

    Private Sub btnGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerate.Click
        PrintVoucher()
    End Sub

    Private Sub PrintVoucher()
        Dim ans As DialogResult = _
            MsgBox("Do you want to Generate Voucher?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "Print")
        If ans = Windows.Forms.DialogResult.No Then Exit Sub

        Dim autoPrintPT As Reporting
        'On Error Resume Next

        Dim printerName As String = GetOption("VoucherPrinter")
        If Not canPrint(printerName) Then Exit Sub

        Dim report As LocalReport = New LocalReport
        autoPrintPT = New Reporting

        Dim mySql As String, dsName As String = "dsVoucher"
        mySql = "Select * From tblVoucher Where Status = 1 Order By ID Asc Rows 1"

        Dim ds As DataSet = LoadSQL(mySql, dsName)

        report.ReportPath = "Reports\rpt_VoucherLayout2.rdlc"
        report.DataSources.Add(New ReportDataSource(dsName, ds.Tables(dsName)))

        Dim addParameters As New Dictionary(Of String, String)
        addParameters.Add("VoucherCode", ds.Tables(0).Rows(0).Item("VCode"))

        Dim paperSize As New Dictionary(Of String, Double)
        paperSize.Add("width", 1.5)
        paperSize.Add("height", 1.8) 'Reprint only

        If Not addParameters Is Nothing Then
            For Each nPara In addParameters
                Dim tmpPara As New ReportParameter
                tmpPara.Name = nPara.Key
                tmpPara.Values.Add(nPara.Value)
                report.SetParameters(New ReportParameter() {tmpPara})
                Console.WriteLine(String.Format("{0}: {1}", nPara.Key, nPara.Value))
            Next
        End If
        If DEV_MODE Then
            frmReport.ReportInit(mySql, dsName, report.ReportPath, addParameters, False)
            frmReport.Show()
        Else
            autoPrintPT.Export(report, paperSize, True)
            autoPrintPT.m_currentPageIndex = 0
            autoPrintPT.Print(printerName)
        End If

        Dim tmpVoucher As VoucherClass
        tmpVoucher = New VoucherClass
        tmpVoucher.UpdateVoucherStatus(ds.Tables(0).Rows(0).Item("ID"))
        Me.Focus()
    End Sub

    Private Function canPrint(ByVal printerName As String) As Boolean
        Try
            Dim printDocument As Drawing.Printing.PrintDocument = New Drawing.Printing.PrintDocument
            printDocument.PrinterSettings.PrinterName = printerName
            Return printDocument.PrinterSettings.IsValid
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub SetPrinterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SetPrinterToolStripMenuItem.Click
        frmSetPrinter.Show()
    End Sub

    Private Sub AboutUsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutUsToolStripMenuItem.Click
        frmAboutUs.Show()
    End Sub

    Private Sub CountOfAvailableVoucherToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CountOfAvailableVoucherToolStripMenuItem.Click
        Dim mysql As String = "Select Count(id)as TotalCount From tblVoucher Where Status = 1"
        Dim ds As DataSet = LoadSQL(mysql, "tblVoucher")

        MsgBox(ds.Tables(0).Rows(0).Item("TotalCount") & " Only Available Voucher", MsgBoxStyle.Information, "Voucher Generator System")
    End Sub
End Class
Public Class frmSetPrinter

    Private Sub btnSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet.Click
        UpdateOptions("VoucherPrinter", cboPrinter.Text)
        MsgBox("Printer Set", MsgBoxStyle.Information, "Voucher Generator System")
        Me.Close()
    End Sub

    Private Sub frmSetPrinter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboPrinter.Items.Clear()
        Dim tmpPrinterName As String
        For Each tmpPrinterName In Printing.PrinterSettings.InstalledPrinters
            cboPrinter.Items.Add(tmpPrinterName)
        Next
        cboPrinter.Text = GetOption("VoucherPrinter")
    End Sub
End Class
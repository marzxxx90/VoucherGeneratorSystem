Imports Microsoft.Reporting.WinForms
Imports System.IO

Public Class frmReport

    Dim subReportPassing As Dictionary(Of String, String)

    Friend Sub ReportInit(ByVal mySql As String, ByVal dsName As String, ByVal rptUrl As String, _
                          Optional ByVal addPara As Dictionary(Of String, String) = Nothing, Optional ByVal hasUser As Boolean = True)
        Try
            Dim ds As DataSet = LoadSQL(mySql, dsName)
            If ds Is Nothing Then Exit Sub

            Console.WriteLine("SQL: " & mySql)
            Console.WriteLine("Max: " & ds.Tables(dsName).Rows.Count)
            Console.WriteLine("Report is Existing? " & System.IO.File.Exists(Application.StartupPath & "\" & rptUrl))
            With rv_display
                .ProcessingMode = ProcessingMode.Local
                .LocalReport.ReportPath = rptUrl
                .LocalReport.DataSources.Clear()

                .LocalReport.DataSources.Add(New ReportDataSource(dsName, ds.Tables(dsName)))

                'If hasUser Then
                ' Dim myPara As New ReportParameter
                ' myPara.Name = "txtUsername"
                ' If POSuser.UserName Is Nothing Then POSuser.UserName = "Sample Eskie"
                ' myPara.Values.Add(POSuser.UserName)
                ' .LocalReport.SetParameters(New ReportParameter() {myPara})
                ' End If

                If Not addPara Is Nothing Then
                    For Each nPara In addPara
                        Dim tmpPara As New ReportParameter
                        tmpPara.Name = nPara.Key
                        tmpPara.Values.Add(nPara.Value)
                        .LocalReport.SetParameters(New ReportParameter() {tmpPara})
                    Next
                End If

                .RefreshReport()
            End With
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "REPORT GENERATE ERROR")
            Log_Report("REPORT - " & ex.ToString)
        End Try
    End Sub

    Friend Sub ExportReport(ByVal mysql As String, ByVal dsName As String, ByVal rptUrl As String, ByVal SavePath As String, _
                            Optional ByVal addPara As Dictionary(Of String, String) = Nothing, Optional ByVal hasUser As Boolean = True, _
                            Optional ByVal ExtractTo As String = "PDF")
        Try
            Dim ds As DataSet = LoadSQL(mysql, dsName)

            Dim warnings As Warning() = Nothing
            Dim streamids As String() = Nothing
            Dim mimeType As String = Nothing
            Dim encoding As String = Nothing
            Dim extension As String = Nothing
            Dim deviceInfo As String = Nothing
            Dim bytes As Byte()

            Console.WriteLine("Process " & rv_display.ShowProgress)
            rv_display.Reset()
            With rv_display

                .ProcessingMode = ProcessingMode.Local
                .LocalReport.ReportPath = rptUrl
                .LocalReport.DataSources.Clear()

                .LocalReport.DataSources.Add(New ReportDataSource(dsName, ds.Tables(dsName)))

                'If hasUser Then
                '    Dim myPara As New ReportParameter
                '    myPara.Name = "txtUsername"
                '    If POSuser.UserName Is Nothing Then POSuser.UserName = "HOO JUN MAA"
                '    myPara.Values.Add(POSuser.UserName)
                '    .LocalReport.SetParameters(New ReportParameter() {myPara})
                'End If

                If Not addPara Is Nothing Then
                    For Each nPara In addPara
                        Dim tmpPara As New ReportParameter
                        tmpPara.Name = nPara.Key
                        tmpPara.Values.Add(nPara.Value)
                        .LocalReport.SetParameters(New ReportParameter() {tmpPara})
                    Next
                End If
                .RefreshReport()
            End With

            bytes = rv_display.LocalReport.Render(ExtractTo, "", mimeType, encoding, extension, streamids, warnings)

            Using Stream As New FileStream(SavePath, FileMode.Create)
                Stream.Write(bytes, 0, bytes.Length)
            End Using

            'Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Friend Sub ReportInit2(ByVal tmp As LocalReport, Optional ByVal addPara As Dictionary(Of String, String) = Nothing)
        Try

            With rv_display
                .ProcessingMode = ProcessingMode.Local
                .LocalReport.ReportPath = tmp.ReportPath
                .LocalReport.DataSources.Clear()

                .LocalReport.DataSources.Add(New ReportDataSource("dsVoucher", tmp.DataSources))

                If Not addPara Is Nothing Then
                    For Each nPara In addPara
                        Dim tmpPara As New ReportParameter
                        tmpPara.Name = nPara.Key
                        tmpPara.Values.Add(nPara.Value)
                        .LocalReport.SetParameters(New ReportParameter() {tmpPara})
                    Next
                End If
                .RefreshReport()
            End With

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "REPORT GENERATE ERROR")
            Log_Report("REPORT - " & ex.ToString)
        End Try
    End Sub
End Class
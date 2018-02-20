Public Class VoucherClass

#Region "Properties"
    Private ID As Integer
    Public Property _id() As Integer
        Get
            Return ID
        End Get
        Set(ByVal value As Integer)
            ID = value
        End Set
    End Property

    Private VCode As String
    Public Property _vCode() As String
        Get
            Return VCode
        End Get
        Set(ByVal value As String)
            VCode = value
        End Set
    End Property

    Private Status As Integer
    Public Property _status() As Integer
        Get
            Return Status
        End Get
        Set(ByVal value As Integer)
            Status = value
        End Set
    End Property

    Private Mins_Time As Integer
    Public Property _mins_Time() As Integer
        Get
            Return Mins_Time
        End Get
        Set(ByVal value As Integer)
            Mins_Time = value
        End Set
    End Property

#End Region

#Region "Procedures"
    Friend Sub SaveVoucher()
        Dim mysql As String = "Select * From tblVoucher Rows 0"
        Dim ds As DataSet = LoadSQL(mysql, "tblVoucher")

        Dim dsNewRow As DataRow
        dsNewRow = ds.Tables(0).NewRow
        With dsNewRow
            .Item("VCode") = _vCode
            .Item("Status") = _status
            .Item("Mins_Time") = _mins_Time
        End With
        ds.Tables(0).Rows.Add(dsNewRow)
        SaveEntry(ds)
    End Sub

    Private Sub LoadByRows(ByVal dr As DataRow)
        With dr
            _id = .Item("ID")
            _vCode = .Item("VCode")
            _status = .Item("Status")
            _mins_Time = .Item("Mins_Time")
        End With
    End Sub

    Friend Function isVoucherExist(ByVal tmpCode As String, ByVal tmpTime As Integer)
        Dim mysql As String = "Select * From tblVoucher Where VCode = '" & tmpCode & "' And Mins_Time = '" & tmpTime & "'"
        Dim ds As DataSet = LoadSQL(mysql, "tblVoucher")

        If ds.Tables(0).Rows.Count >= 1 Then Return True

        Return False
    End Function

    Friend Sub UpdateVoucherStatus(ByVal idx As Integer)
        Dim mysql As String = "Select * From tblVoucher Where id = " & idx
        Dim ds As DataSet = LoadSQL(mysql, "tblVoucher")

        ds.Tables(0).Rows(0).Item("Status") = 0

        SaveEntry(ds, False)
    End Sub
#End Region

End Class

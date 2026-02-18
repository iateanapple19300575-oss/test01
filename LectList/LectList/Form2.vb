Public Class Form2
    Private _list As List(Of DailyWorkRow)

    Private Sub LoadData()

        'Dim dt As DataTable = LoadDailyWorkViewTable2()

        _list = New List(Of DailyWorkRow)

        'For Each r As DataRow In dt.Rows
        '    _list.Add(New DailyWorkRow With {
        '        .StaffId = CStr(r("StaffId")),
        '        .WorkDate = CDate(r("WorkDate")),
        '        .MainStart = If(IsDBNull(r("MainStart")), Nothing, CDate(r("MainStart"))),
        '        .MainEnd = If(IsDBNull(r("MainEnd")), Nothing, CDate(r("MainEnd"))),
        '        .ErrorMessage = CStr(r("ErrorMessage"))
        '    })
        'Next

    End Sub

    Private Sub SetupGrid()

        With DataGridView1
            .VirtualMode = True
            .RowCount = _list.Count
            .EnableDoubleBuffer() ' ← これも必須
        End With

    End Sub

    Private Sub DataGridView1_CellValueNeeded(
    sender As Object, e As DataGridViewCellValueEventArgs
) Handles DataGridView1.CellValueNeeded

        Dim row = _list(e.RowIndex)

        Select Case DataGridView1.Columns(e.ColumnIndex).Name

            Case "StaffId"
                e.Value = row.StaffId

            Case "WorkDate"
                e.Value = row.WorkDate.ToString("yyyy/MM/dd")

            Case "MainStart"
                If row.MainStart.HasValue Then
                    e.Value = row.MainStart.Value.ToString("HH:mm")
                End If

            Case "MainEnd"
                If row.MainEnd.HasValue Then
                    e.Value = row.MainEnd.Value.ToString("HH:mm")
                End If

            Case "ErrorMessage"
                e.Value = row.ErrorMessage

        End Select

    End Sub

    Private Sub DataGridView1_RowPrePaint(
    sender As Object, e As DataGridViewRowPrePaintEventArgs
) Handles DataGridView1.RowPrePaint

        Dim row = _list(e.RowIndex)

        If row.ErrorMessage <> "" Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.MistyRose
        End If

    End Sub


End Class
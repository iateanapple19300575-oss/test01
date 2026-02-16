Public Class Logic

    Public Sub aaaaa()
        Dim staffList = staffWorkDays _
            .Select(Function(x) x.StaffId) _
            .Distinct() _
            .ToList()

        Dim col As DataGridViewComboBoxColumn = CType(DataGridView1.Columns("StaffColumn"), DataGridViewComboBoxColumn)
        col.DataSource = staffList


        Dim items = staffWorkDays _
            .Select(Function(x) New With {
                .Key = x.StaffId & " " & x.WorkDate.ToString("yyyy/MM/dd"),
                .Value = x
            }) _
            .ToList()

        Dim col As DataGridViewComboBoxColumn = CType(DataGridView1.Columns("WorkDayColumn"), DataGridViewComboBoxColumn)
        col.DisplayMember = "Key"
        col.ValueMember = "Value"
        col.DataSource = items


        Dim items As New List(Of Object)

        For Each swd In staffWorkDays
            For Each r In swd.TimeRanges
                items.Add(New With {
                    .Key = swd.StaffId & " " &
                           swd.WorkDate.ToString("yyyy/MM/dd") & " " &
                           r.StartTime.ToString("HH:mm") & "-" &
                           r.EndTime.ToString("HH:mm"),
                    .Value = r
                })
            Next
        Next

        Dim col As DataGridViewComboBoxColumn = CType(DataGridView1.Columns("WorkRangeColumn"), DataGridViewComboBoxColumn)
        col.DisplayMember = "Key"
        col.ValueMember = "Value"
        col.DataSource = items

    End Sub


End Class

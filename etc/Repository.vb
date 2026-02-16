Imports System.Data.SqlClient

Public Class Repository
    Public Function LoadStaffWorkList() As List(Of StaffWorkDay)

        Dim result As New Dictionary(Of String, StaffWorkDay)()

        Dim sql As String = "
        SELECT StaffId, WorkDate, StartTime, EndTime
        FROM StaffWork
        ORDER BY StaffId, WorkDate, StartTime
    "

        Using cn As New SqlConnection("接続文字列")
            cn.Open()

            Using cmd As New SqlCommand(sql, cn)
                Using rd As SqlDataReader = cmd.ExecuteReader()

                    While rd.Read()

                        Dim staffId As String = rd("StaffId").ToString()
                        Dim workDate As Date = CDate(rd("WorkDate"))

                        Dim key As String = staffId & "|" & workDate.ToString("yyyyMMdd")

                        If Not result.ContainsKey(key) Then
                            Dim swd As New StaffWorkDay()
                            swd.StaffId = staffId
                            swd.WorkDate = workDate
                            result.Add(key, swd)
                        End If

                        ' 時間帯作成
                        Dim range As New WorkTimeRange()
                        range.StartTime = CDate(rd("StartTime"))
                        range.EndTime = CDate(rd("EndTime"))

                        ' ★ 日跨ぎ補正 ★
                        range.FixCrossDay()

                        result(key).TimeRanges.Add(range)

                    End While

                End Using
            End Using
        End Using

        Return result.Values.ToList()

    End Function
End Class

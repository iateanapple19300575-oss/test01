Public Class StaffWorkDay

    Public Property StaffId As String
    Public Property WorkDate As Date
    Public Property TimeRanges As List(Of WorkTimeRange)

    Public Sub New()
        TimeRanges = New List(Of WorkTimeRange)()
    End Sub

    ' 1日の総勤務時間
    Public ReadOnly Property TotalDuration As TimeSpan
        Get
            Dim sum As TimeSpan = TimeSpan.Zero
            For Each r In TimeRanges
                sum += r.Duration
            Next
            Return sum
        End Get
    End Property

    ' 深夜帯の総計
    Public ReadOnly Property TotalNight As TimeSpan
        Get
            Dim sum As TimeSpan = TimeSpan.Zero
            For Each r In TimeRanges
                sum += r.NightDuration
            Next
            Return sum
        End Get
    End Property

    ' 重複チェック（業務システムで必須）
    Public Function HasOverlap() As Boolean
        For i As Integer = 0 To TimeRanges.Count - 2
            For j As Integer = i + 1 To TimeRanges.Count - 1
                If TimeRanges(i).IsOverlap(TimeRanges(j)) Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function

    '総計
    Public ReadOnly Property TotalNightExtra As TimeSpan
        Get
            Dim sum As TimeSpan = TimeSpan.Zero
            For Each r In TimeRanges
                sum += r.NightExtra
            Next
            Return sum
        End Get
    End Property
End Class
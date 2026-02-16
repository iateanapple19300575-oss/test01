Public Class WorkTimeRange

    Public Property StartTime As DateTime
    Public Property EndTime As DateTime

    ' 補正後の実時間（複数日対応）
    Public ReadOnly Property Duration As TimeSpan
        Get
            Return EndTime - StartTime
        End Get
    End Property

    ' 深夜帯（00:00–05:00）の時間
    Public ReadOnly Property NightDuration As TimeSpan
        Get
            Return CalcOverlap(StartTime, EndTime,
                               StartTime.Date.AddHours(0),
                               StartTime.Date.AddHours(5))
        End Get
    End Property

    ' 早朝帯（05:00–09:00）なども必要なら追加可能

    Public ReadOnly Property EarlyDuration As TimeSpan
        Get
            Return CalcOverlap(StartTime, EndTime,
                               StartTime.Date.AddHours(5),
                               StartTime.Date.AddHours(9))
        End Get
    End Property

    ' 通常帯（全体 − 深夜 − 早朝
    Public ReadOnly Property NormalDuration As TimeSpan
        Get
            Return Duration - NightDuration - EarlyDuration
        End Get
    End Property

    ' ★ 日跨ぎ補正（翌日・翌々日も対応）
    Public Sub FixCrossDay()
        If EndTime < StartTime Then
            EndTime = EndTime.AddDays(1)
        End If
    End Sub

    ' ★ 深夜割増（25%）
    Public ReadOnly Property NightExtra As TimeSpan
        Get
            Dim extraMinutes As Double = NightDuration.TotalMinutes * 0.25
            Return TimeSpan.FromMinutes(extraMinutes)
        End Get
    End Property

    ' 重複判定
    Public Function IsOverlap(other As WorkTimeRange) As Boolean
        Return (Me.StartTime < other.EndTime) AndAlso
               (other.StartTime < Me.EndTime)
    End Function

    ' 任意の時間帯との重なり計算
    Private Function CalcOverlap(s1 As DateTime, e1 As DateTime,
                                 s2 As DateTime, e2 As DateTime) As TimeSpan

        Dim startMax As DateTime = If(s1 > s2, s1, s2)
        Dim endMin As DateTime = If(e1 < e2, e1, e2)

        If endMin <= startMax Then
            Return TimeSpan.Zero
        End If

        Return endMin - startMax
    End Function

End Class
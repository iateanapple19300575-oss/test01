Public Class WorkDayInfo
    Public Property StaffId As String
    Public Property WorkDate As Date

    Public Property MainStart As Date?
    Public Property MainEnd As Date?

    Public Property AWorkNo As New List(Of String)
    Public Property AStart As New List(Of Date)
    Public Property AEnd As New List(Of Date)

    Public Property BWorkNo As New List(Of String)
    Public Property BStart As New List(Of Date)
    Public Property BEnd As New List(Of Date)

    Public Property CWorkNo As New List(Of String)
    Public Property CStart As New List(Of Date)
    Public Property CEnd As New List(Of Date)

    ' 重複件数
    Public Property CntAA As Integer
    Public Property CntBB As Integer
    Public Property CntCC As Integer
    Public Property CntAB As Integer
    Public Property CntAC As Integer
    Public Property CntBC As Integer

    ' 主テーブル整合性
    Public Property ABeforeStart As Integer?
    Public Property AAfterEnd As Integer?
    Public Property BBeforeStart As Integer?
    Public Property BAfterEnd As Integer?
    Public Property CBeforeStart As Integer?
    Public Property CAfterEnd As Integer?
End Class
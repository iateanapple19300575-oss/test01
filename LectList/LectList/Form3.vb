
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Runtime.CompilerServices

Public Class Form3

    '===============================
    ' 構造体：1 行分の表示データ
    '===============================
    Public Structure DailyWorkRow
        Public StaffId As String
        Public WorkDate As Date
        Public MainStart As Date
        Public MainEnd As Date
        Public ErrorMessage As String
    End Structure

    '===============================
    ' フィールド
    '===============================
    Private _connStr As String = "Data Source = DESKTOP-L98IE79;Initial Catalog = DeveloperDB;Integrated Security = SSPI"
    Private _rows() As DailyWorkRow      ' 全データ
    Private _viewIndex() As Integer      ' 表示用インデックス

    '===============================
    ' フォームロード
    '===============================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' DataGridView 初期設定
        SetupGridColumns()
        DataGridView1.EnableDoubleBuffer()

        ' データ読み込み
        LoadData()

        ' VirtualMode 設定
        SetupVirtualMode()

        ' 行ヘッダ幅調整
        AdjustRowHeaderWidth()

    End Sub

    '===============================
    ' DataGridView 列定義
    '===============================
    Private Sub SetupGridColumns()

        With DataGridView1
            .Columns.Clear()
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .ReadOnly = True
            .RowHeadersVisible = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Dim col As DataGridViewTextBoxColumn

            col = New DataGridViewTextBoxColumn()
            col.Name = "StaffId"
            col.HeaderText = "担当者"
            col.Width = 80
            .Columns.Add(col)

            col = New DataGridViewTextBoxColumn()
            col.Name = "WorkDate"
            col.HeaderText = "日付"
            col.Width = 90
            .Columns.Add(col)

            col = New DataGridViewTextBoxColumn()
            col.Name = "MainStart"
            col.HeaderText = "出勤"
            col.Width = 60
            .Columns.Add(col)

            col = New DataGridViewTextBoxColumn()
            col.Name = "MainEnd"
            col.HeaderText = "退勤"
            col.Width = 60
            .Columns.Add(col)

            col = New DataGridViewTextBoxColumn()
            col.Name = "ErrorMessage"
            col.HeaderText = "エラー"
            col.Width = 300
            .Columns.Add(col)
        End With

    End Sub

    '===============================
    ' データ読み込み（DailyWorkViewTable → 構造体配列）
    '===============================
    Private Sub LoadData()

        Dim dt As New DataTable()

        Dim sql As String = "
SELECT StaffId, WorkDate, MainStart, MainEnd, ErrorMessage
FROM DailyWorkViewTable
ORDER BY StaffId, WorkDate;
"

        Using cn As New SqlConnection(_connStr)
            Using ad As New SqlDataAdapter(sql, cn)
                ad.Fill(dt)
            End Using
        End Using

        ReDim _rows(dt.Rows.Count - 1)

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim r = dt.Rows(i)

            _rows(i).StaffId = CStr(r("StaffId"))
            _rows(i).WorkDate = CDate(r("WorkDate"))
            _rows(i).MainStart = If(IsDBNull(r("MainStart")), Date.MinValue, CDate(r("MainStart")))
            _rows(i).MainEnd = If(IsDBNull(r("MainEnd")), Date.MinValue, CDate(r("MainEnd")))
            _rows(i).ErrorMessage = CStr(r("ErrorMessage"))
        Next

        ' 初期表示は全件
        _viewIndex = Enumerable.Range(0, _rows.Length).ToArray()

    End Sub

    '===============================
    ' VirtualMode 設定
    '===============================
    Private Sub SetupVirtualMode()

        With DataGridView1
            .VirtualMode = True
            .RowCount = _viewIndex.Length
        End With

    End Sub

    '===============================
    ' CellValueNeeded：表示値を返す
    '===============================
    Private Sub DataGridView1_CellValueNeeded(
        sender As Object, e As DataGridViewCellValueEventArgs
    ) Handles DataGridView1.CellValueNeeded

        If _viewIndex Is Nothing OrElse _viewIndex.Length = 0 Then
            Return
        End If

        If e.RowIndex < 0 OrElse e.RowIndex >= _viewIndex.Length Then
            Return
        End If

        Dim realIndex As Integer = _viewIndex(e.RowIndex)
        Dim row = _rows(realIndex)

        Select Case DataGridView1.Columns(e.ColumnIndex).Name

            Case "StaffId"
                e.Value = row.StaffId

            Case "WorkDate"
                e.Value = row.WorkDate.ToString("yyyy/MM/dd")

            Case "MainStart"
                If row.MainStart <> Date.MinValue Then
                    e.Value = row.MainStart.ToString("HH:mm")
                End If

            Case "MainEnd"
                If row.MainEnd <> Date.MinValue Then
                    e.Value = row.MainEnd.ToString("HH:mm")
                End If

            Case "ErrorMessage"
                e.Value = row.ErrorMessage

        End Select

    End Sub

    '===============================
    ' RowPrePaint：エラー行の色付け
    '===============================
    Private Sub DataGridView1_RowPrePaint(
        sender As Object, e As DataGridViewRowPrePaintEventArgs
    ) Handles DataGridView1.RowPrePaint

        If _viewIndex Is Nothing OrElse _viewIndex.Length = 0 Then Return
        If e.RowIndex < 0 OrElse e.RowIndex >= _viewIndex.Length Then Return

        Dim realIndex As Integer = _viewIndex(e.RowIndex)
        If _rows(realIndex).ErrorMessage <> "" Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.MistyRose
        End If

    End Sub

    '===============================
    ' RowPostPaint：行番号表示（高速版）
    '===============================
    Private Sub DataGridView1_RowPostPaint(
        ByVal sender As Object,
        ByVal e As DataGridViewRowPostPaintEventArgs
    ) Handles DataGridView1.RowPostPaint

        Dim grid As DataGridView = DirectCast(sender, DataGridView)

        Dim rowNumber As String = (e.RowIndex + 1).ToString()

        Dim centerFormat As New StringFormat()
        centerFormat.Alignment = StringAlignment.Center
        centerFormat.LineAlignment = StringAlignment.Center

        Dim headerBounds As New Rectangle(
            e.RowBounds.Left,
            e.RowBounds.Top,
            grid.RowHeadersWidth,
            e.RowBounds.Height
        )

        e.Graphics.DrawString(
            rowNumber,
            grid.Font,
            SystemBrushes.ControlText,
            headerBounds,
            centerFormat
        )

    End Sub

    '===============================
    ' 行ヘッダ幅調整
    '===============================
    Private Sub AdjustRowHeaderWidth()

        Dim maxRow As Integer = DataGridView1.RowCount
        Dim digits As Integer = Math.Max(1, maxRow.ToString().Length)

        DataGridView1.RowHeadersWidth = 30 + (digits * 6)

    End Sub

    '===============================
    ' 担当者コードフィルタ（例）
    ' TextBox txtStaffFilter + Button btnFilter を想定
    '===============================
    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click

        Dim keyword As String = txtStaffFilter.Text.Trim()

        If keyword = "" Then
            ' フィルタ解除
            _viewIndex = Enumerable.Range(0, _rows.Length).ToArray()
        Else
            Dim list As New List(Of Integer)

            For i As Integer = 0 To _rows.Length - 1
                If _rows(i).StaffId.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    list.Add(i)
                End If
            Next

            _viewIndex = list.ToArray()
        End If

        DataGridView1.RowCount = _viewIndex.Length
        AdjustRowHeaderWidth()
        DataGridView1.Refresh()

    End Sub

End Class

''===============================
'' DataGridView 拡張メソッド：ダブルバッファリング
''===============================
'Public Module DataGridViewExtensions

'    <Extension()>
'    Public Sub EnableDoubleBuffer(grid As DataGridView)

'        Dim dgvType As Type = grid.GetType()
'        Dim prop As PropertyInfo = dgvType.GetProperty(
'            "DoubleBuffered",
'            BindingFlags.Instance Or BindingFlags.NonPublic)

'        If prop IsNot Nothing Then
'            prop.SetValue(grid, True, Nothing)
'        End If

'    End Sub

'End Module
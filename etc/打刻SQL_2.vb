Private Function BuildDailyWorkSql() As String

    Dim sb As New System.Text.StringBuilder()

    sb.AppendLine("WITH StaffDates AS (")
    sb.AppendLine("    SELECT StaffId, WorkDate FROM MainTable WHERE WorkDate BETWEEN @StartDate AND @EndDate")
    sb.AppendLine("    UNION")
    sb.AppendLine("    SELECT StaffId, WorkDate FROM TableA WHERE WorkDate BETWEEN @StartDate AND @EndDate")
    sb.AppendLine("    UNION")
    sb.AppendLine("    SELECT StaffId, WorkDate FROM TableB WHERE WorkDate BETWEEN @StartDate AND @EndDate")
    sb.AppendLine("    UNION")
    sb.AppendLine("    SELECT StaffId, WorkDate FROM TableC WHERE WorkDate BETWEEN @StartDate AND @EndDate")
    sb.AppendLine("),")

    sb.AppendLine("ARows AS (")
    sb.AppendLine("    SELECT *, ROW_NUMBER() OVER (PARTITION BY StaffId, WorkDate ORDER BY StartTime) AS RN")
    sb.AppendLine("    FROM TableA")
    sb.AppendLine("),")

    sb.AppendLine("BRows AS (")
    sb.AppendLine("    SELECT *, ROW_NUMBER() OVER (PARTITION BY StaffId, WorkDate ORDER BY StartTime) AS RN")
    sb.AppendLine("    FROM TableB")
    sb.AppendLine("),")

    sb.AppendLine("CRows AS (")
    sb.AppendLine("    SELECT *, ROW_NUMBER() OVER (PARTITION BY StaffId, WorkDate ORDER BY StartTime) AS RN")
    sb.AppendLine("    FROM TableC")
    sb.AppendLine("),")

    sb.AppendLine("AMinMax AS (")
    sb.AppendLine("    SELECT StaffId, WorkDate, MIN(StartTime) AS AMinStart, MAX(EndTime) AS AMaxEnd")
    sb.AppendLine("    FROM TableA GROUP BY StaffId, WorkDate")
    sb.AppendLine("),")

    sb.AppendLine("BMinMax AS (")
    sb.AppendLine("    SELECT StaffId, WorkDate, MIN(StartTime) AS BMinStart, MAX(EndTime) AS BMaxEnd")
    sb.AppendLine("    FROM TableB GROUP BY StaffId, WorkDate")
    sb.AppendLine("),")

    sb.AppendLine("CMinMax AS (")
    sb.AppendLine("    SELECT StaffId, WorkDate, MIN(StartTime) AS CMinStart, MAX(EndTime) AS CMaxEnd")
    sb.AppendLine("    FROM TableC GROUP BY StaffId, WorkDate")
    sb.AppendLine("),")

    '--- AExpanded（3件） ---
    sb.AppendLine("AExpanded AS (")
    sb.AppendLine("    SELECT StaffId, WorkDate,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN WorkNo END) AS AWorkNo1,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN StartTime END) AS AStart1,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN EndTime END) AS AEnd1,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN WorkNo END) AS AWorkNo2,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN StartTime END) AS AStart2,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN EndTime END) AS AEnd2,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN WorkNo END) AS AWorkNo3,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN StartTime END) AS AStart3,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN EndTime END) AS AEnd3")
    sb.AppendLine("    FROM ARows GROUP BY StaffId, WorkDate")
    sb.AppendLine("),")

    '--- BExpanded（3件） ---
    sb.AppendLine("BExpanded AS (")
    sb.AppendLine("    SELECT StaffId, WorkDate,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN WorkNo END) AS BWorkNo1,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN StartTime END) AS BStart1,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN EndTime END) AS BEnd1,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN WorkNo END) AS BWorkNo2,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN StartTime END) AS BStart2,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN EndTime END) AS BEnd2,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN WorkNo END) AS BWorkNo3,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN StartTime END) AS BStart3,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN EndTime END) AS BEnd3")
    sb.AppendLine("    FROM BRows GROUP BY StaffId, WorkDate")
    sb.AppendLine("),")

    '--- CExpanded（3件） ---
    sb.AppendLine("CExpanded AS (")
    sb.AppendLine("    SELECT StaffId, WorkDate,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN WorkNo END) AS CWorkNo1,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN StartTime END) AS CStart1,")
    sb.AppendLine("        MAX(CASE WHEN RN=1 THEN EndTime END) AS CEnd1,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN WorkNo END) AS CWorkNo2,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN StartTime END) AS CStart2,")
    sb.AppendLine("        MAX(CASE WHEN RN=2 THEN EndTime END) AS CEnd2,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN WorkNo END) AS CWorkNo3,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN StartTime END) AS CStart3,")
    sb.AppendLine("        MAX(CASE WHEN RN=3 THEN EndTime END) AS CEnd3")
    sb.AppendLine("    FROM CRows GROUP BY StaffId, WorkDate")
    sb.AppendLine("),")

    '--- 重複チェック ---
    sb.AppendLine("OverlapDetails AS (")
    sb.AppendLine("    SELECT s.StaffId, s.WorkDate,")
    sb.AppendLine("        (SELECT COUNT(*) FROM TableA a1 JOIN TableA a2")
    sb.AppendLine("         ON a1.StaffId=s.StaffId AND a2.StaffId=s.StaffId")
    sb.AppendLine("        AND a1.WorkDate=s.WorkDate AND a2.WorkDate=s.WorkDate")
    sb.AppendLine("        AND a1.Id<a2.Id AND a1.StartTime<a2.EndTime AND a1.EndTime>a2.StartTime) AS CntAA,")
    sb.AppendLine("        (SELECT COUNT(*) FROM TableB b1 JOIN TableB b2")
    sb.AppendLine("         ON b1.StaffId=s.StaffId AND b2.StaffId=s.StaffId")
    sb.AppendLine("        AND b1.WorkDate=s.WorkDate AND b2.WorkDate=s.WorkDate")
    sb.AppendLine("        AND b1.Id<b2.Id AND b1.StartTime<b2.EndTime AND b1.EndTime>b2.StartTime) AS CntBB,")
    sb.AppendLine("        (SELECT COUNT(*) FROM TableC c1 JOIN TableC c2")
    sb.AppendLine("         ON c1.StaffId=s.StaffId AND c2.StaffId=s.StaffId")
    sb.AppendLine("        AND c1.WorkDate=s.WorkDate AND c2.WorkDate=s.WorkDate")
    sb.AppendLine("        AND c1.Id<c2.Id AND c1.StartTime<c2.EndTime AND c1.EndTime>c2.StartTime) AS CntCC,")
    sb.AppendLine("        (SELECT COUNT(*) FROM TableA a JOIN TableB b")
    sb.AppendLine("         ON a.StaffId=s.StaffId AND b.StaffId=s.StaffId")
    sb.AppendLine("        AND a.WorkDate=s.WorkDate AND b.WorkDate=s.WorkDate")
    sb.AppendLine("        AND a.StartTime<b.EndTime AND a.EndTime>b.StartTime) AS CntAB,")
    sb.AppendLine("        (SELECT COUNT(*) FROM TableA a JOIN TableC c")
    sb.AppendLine("         ON a.StaffId=s.StaffId AND c.StaffId=s.StaffId")
    sb.AppendLine("        AND a.WorkDate=s.WorkDate AND c.WorkDate=s.WorkDate")
    sb.AppendLine("        AND a.StartTime<c.EndTime AND a.EndTime>c.StartTime) AS CntAC,")
    sb.AppendLine("        (SELECT COUNT(*) FROM TableB b JOIN TableC c")
    sb.AppendLine("         ON b.StaffId=s.StaffId AND c.StaffId=s.StaffId")
    sb.AppendLine("        AND b.WorkDate=s.WorkDate AND c.WorkDate=s.WorkDate")
    sb.AppendLine("        AND b.StartTime<c.EndTime AND b.EndTime>c.StartTime) AS CntBC")
    sb.AppendLine("    FROM StaffDates s")
    sb.AppendLine("),")

    '--- MainCheck ---
    sb.AppendLine("MainCheck AS (")
    sb.AppendLine("    SELECT s.StaffId, s.WorkDate,")
    sb.AppendLine("        CASE WHEN m.StaffId IS NULL THEN @MainMissingMode")
    sb.AppendLine("             WHEN EXISTS (SELECT 1 FROM TableA a WHERE a.StaffId=s.StaffId AND a.WorkDate=s.WorkDate AND a.StartTime<m.StartTime)")
    sb.AppendLine("             THEN 1 ELSE 0 END AS ABeforeStart,")
    sb.AppendLine("        CASE WHEN m.StaffId IS NULL THEN @MainMissingMode")
    sb.AppendLine("             WHEN EXISTS (SELECT 1 FROM TableA a WHERE a.StaffId=s.StaffId AND a.WorkDate=s.WorkDate AND a.EndTime>m.EndTime)")
    sb.AppendLine("             THEN 1 ELSE 0 END AS AAfterEnd,")
    sb.AppendLine("        CASE WHEN m.StaffId IS NULL THEN @MainMissingMode")
    sb.AppendLine("             WHEN EXISTS (SELECT 1 FROM TableB b WHERE b.StaffId=s.StaffId AND b.WorkDate=s.WorkDate AND b.StartTime<m.StartTime)")
    sb.AppendLine("             THEN 1 ELSE 0 END AS BBeforeStart,")
    sb.AppendLine("        CASE WHEN m.StaffId IS NULL THEN @MainMissingMode")
    sb.AppendLine("             WHEN EXISTS (SELECT 1 FROM TableB b WHERE b.StaffId=s.StaffId AND b.WorkDate=s.WorkDate AND b.EndTime>m.EndTime)")
    sb.AppendLine("             THEN 1 ELSE 0 END AS BAfterEnd,")
    sb.AppendLine("        CASE WHEN m.StaffId IS NULL THEN @MainMissingMode")
    sb.AppendLine("             WHEN EXISTS (SELECT 1 FROM TableC c WHERE c.StaffId=s.StaffId AND c.WorkDate=s.WorkDate AND c.StartTime<m.StartTime)")
    sb.AppendLine("             THEN 1 ELSE 0 END AS CBeforeStart,")
    sb.AppendLine("        CASE WHEN m.StaffId IS NULL THEN @MainMissingMode")
    sb.AppendLine("             WHEN EXISTS (SELECT 1 FROM TableC c WHERE c.StaffId=s.StaffId AND c.WorkDate=s.WorkDate AND c.EndTime>m.EndTime)")
    sb.AppendLine("             THEN 1 ELSE 0 END AS CAfterEnd")
    sb.AppendLine("    FROM StaffDates s")
    sb.AppendLine("    LEFT JOIN MainTable m ON m.StaffId=s.StaffId AND m.WorkDate=s.WorkDate")
    sb.AppendLine(")")

    '--- 最終 SELECT ---
    sb.AppendLine("SELECT")
    sb.AppendLine("    s.StaffId, s.WorkDate,")
    sb.AppendLine("    m.StartTime AS MainStart, m.EndTime AS MainEnd,")
    sb.AppendLine("    amin.AMinStart, amin.AMaxEnd,")
    sb.AppendLine("    bmin.BMinStart, bmin.BMaxEnd,")
    sb.AppendLine("    cmin.CMinStart, cmin.CMaxEnd,")
    sb.AppendLine("    a.AWorkNo1, a.AStart1, a.AEnd1,")
    sb.AppendLine("    a.AWorkNo2, a.AStart2, a.AEnd2,")
    sb.AppendLine("    a.AWorkNo3, a.AStart3, a.AEnd3,")
    sb.AppendLine("    b.BWorkNo1, b.BStart1, b.BEnd1,")
    sb.AppendLine("    b.BWorkNo2, b.BStart2, b.BEnd2,")
    sb.AppendLine("    b.BWorkNo3, b.BStart3, b.BEnd3,")
    sb.AppendLine("    c.CWorkNo1, c.CStart1, c.CEnd1,")
    sb.AppendLine("    c.CWorkNo2, c.CStart2, c.CEnd2,")
    sb.AppendLine("    c.CWorkNo3, c.CStart3, c.CEnd3,")
    sb.AppendLine("    o.CntAA, o.CntBB, o.CntCC, o.CntAB, o.CntAC, o.CntBC,")
    sb.AppendLine("    mc.ABeforeStart, mc.AAfterEnd, mc.BBeforeStart, mc.BAfterEnd, mc.CBeforeStart, mc.CAfterEnd")
    sb.AppendLine("FROM StaffDates s")
    sb.AppendLine("LEFT JOIN MainTable m ON m.StaffId=s.StaffId AND m.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN AMinMax amin ON amin.StaffId=s.StaffId AND amin.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN BMinMax bmin ON bmin.StaffId=s.StaffId AND bmin.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN CMinMax cmin ON cmin.StaffId=s.StaffId AND cmin.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN AExpanded a ON a.StaffId=s.StaffId AND a.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN BExpanded b ON b.StaffId=s.StaffId AND b.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN CExpanded c ON c.StaffId=s.StaffId AND c.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN OverlapDetails o ON o.StaffId=s.StaffId AND o.WorkDate=s.WorkDate")
    sb.AppendLine("LEFT JOIN MainCheck mc ON mc.StaffId=s.StaffId AND mc.WorkDate=s.WorkDate")
    sb.AppendLine("WHERE m.StaffId IS NOT NULL OR a.AWorkNo1 IS NOT NULL OR b.BWorkNo1 IS NOT NULL OR c.CWorkNo1 IS NOT NULL")
    sb.AppendLine("ORDER BY s.StaffId, s.WorkDate;")

    Return sb.ToString()

End Function


Public Function LoadWorkDayTable(startDate As Date, endDate As Date) As DataTable

    Dim dt As New DataTable()
    Dim sql As String = BuildDailyWorkSql()

    Using cn As New SqlConnection("Your Connection String")
        Using cmd As New SqlCommand(sql, cn)

            cmd.Parameters.AddWithValue("@StartDate", startDate)
            cmd.Parameters.AddWithValue("@EndDate", endDate)
            cmd.Parameters.AddWithValue("@MainMissingMode", 1) ' 0/1/2 切替可能

            Using ad As New SqlDataAdapter(cmd)
                ad.Fill(dt)
            End Using

        End Using
    End Using

    Return dt

End Function




Private Sub LoadGrid()

    Dim dt As DataTable = LoadWorkDayTable(#2026/02/01#, #2026/02/28#)

    DataGridView1.AutoGenerateColumns = True
    DataGridView1.DataSource = dt

End Sub



Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    LoadGrid()
    FormatGrid()

End Sub


Private Sub FormatGrid()

    With DataGridView1

        .Columns("WorkDate").DefaultCellStyle.Format = "yyyy/MM/dd"

        .Columns("MainStart").DefaultCellStyle.Format = "HH:mm"
        .Columns("MainEnd").DefaultCellStyle.Format = "HH:mm"

        ' A/B/C の Start/End も同様に
        For Each colName As String In New String() {
            "AStart1","AEnd1","AStart2","AEnd2","AStart3","AEnd3",
            "BStart1","BEnd1","BStart2","BEnd2","BStart3","BEnd3",
            "CStart1","CEnd1","CStart2","CEnd2","CStart3","CEnd3"
        }
            If .Columns.Contains(colName) Then
                .Columns(colName).DefaultCellStyle.Format = "HH:mm"
            End If
        Next

    End With

End Sub







???????????????????????????????????
Imports System.Data
Imports System.Data.SqlClient

Public Function LoadWorkDayTable(startDate As Date, endDate As Date) As DataTable

    Dim dt As New DataTable()

    ' ★ 完全版 SQL をリソース or 文字列で保持
    Dim sql As String = My.Resources.Sql_DailyWorkView   ' ← 推奨

    Using cn As New SqlConnection("Your Connection String")
        Using cmd As New SqlCommand(sql, cn)

            cmd.Parameters.AddWithValue("@StartDate", startDate)
            cmd.Parameters.AddWithValue("@EndDate", endDate)

            Using ad As New SqlDataAdapter(cmd)
                ad.Fill(dt)
            End Using

        End Using
    End Using

    Return dt

End Function

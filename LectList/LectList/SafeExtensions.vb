Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices

'===========================================================
'  SafeExtensions.vb
'  FW3.5 用 安全変換ユーティリティ + DataGridView バインド補助
'===========================================================

Public Module SafeExtensions

    '---------------------------------------------
    '  DataRow 安全取得
    '---------------------------------------------
    <Extension()>
    Public Function GetStringSafe(row As DataRow, column As String) As String
        If row.IsNull(column) Then Return ""
        Return CStr(row(column))
    End Function

    <Extension()>
    Public Function GetIntSafe(row As DataRow, column As String) As Integer
        If row.IsNull(column) Then Return 0
        Return CInt(row(column))
    End Function

    <Extension()>
    Public Function GetDecimalSafe(row As DataRow, column As String) As Decimal
        If row.IsNull(column) Then Return 0D
        Return CDec(row(column))
    End Function

    <Extension()>
    Public Function GetDateSafe(row As DataRow, column As String) As Date?
        If row.IsNull(column) Then Return Nothing
        Return CDate(row(column))
    End Function

    <Extension()>
    Public Function GetBoolSafe(row As DataRow, column As String) As Boolean
        If row.IsNull(column) Then Return False
        Return CBool(row(column))
    End Function


    '---------------------------------------------
    '  SqlDataReader 安全取得
    '---------------------------------------------
    <Extension()>
    Public Function GetStringSafe(reader As SqlDataReader, column As String) As String
        Dim idx As Integer = reader.GetOrdinal(column)
        If reader.IsDBNull(idx) Then Return ""
        Return reader.GetString(idx)
    End Function

    <Extension()>
    Public Function GetIntSafe(reader As SqlDataReader, column As String) As Integer
        Dim idx As Integer = reader.GetOrdinal(column)
        If reader.IsDBNull(idx) Then Return 0
        Return reader.GetInt32(idx)
    End Function

    <Extension()>
    Public Function GetDecimalSafe(reader As SqlDataReader, column As String) As Decimal
        Dim idx As Integer = reader.GetOrdinal(column)
        If reader.IsDBNull(idx) Then Return 0D
        Return reader.GetDecimal(idx)
    End Function

    <Extension()>
    Public Function GetDateSafe(reader As SqlDataReader, column As String) As Date?
        Dim idx As Integer = reader.GetOrdinal(column)
        If reader.IsDBNull(idx) Then Return Nothing
        Return reader.GetDateTime(idx)
    End Function

    <Extension()>
    Public Function GetBoolSafe(reader As SqlDataReader, column As String) As Boolean
        Dim idx As Integer = reader.GetOrdinal(column)
        If reader.IsDBNull(idx) Then Return False
        Return reader.GetBoolean(idx)
    End Function


    '---------------------------------------------
    '  Object → 安全変換（DataGridView Cell.Value 用）
    '---------------------------------------------
    <Extension()>
    Public Function ToStringSafe(obj As Object) As String
        If obj Is Nothing OrElse obj Is DBNull.Value Then Return ""
        Return CStr(obj)
    End Function

    <Extension()>
    Public Function ToIntSafe(obj As Object) As Integer
        If obj Is Nothing OrElse obj Is DBNull.Value Then Return 0
        Return Convert.ToInt32(obj)
    End Function

    <Extension()>
    Public Function ToDateSafe(obj As Object) As Date?
        If obj Is Nothing OrElse obj Is DBNull.Value Then Return Nothing
        Return Convert.ToDateTime(obj)
    End Function

    <Extension()>
    Public Function ToDecimalSafe(obj As Object) As Decimal
        If obj Is Nothing OrElse obj Is DBNull.Value Then Return 0D
        Return Convert.ToDecimal(obj)
    End Function


    '---------------------------------------------
    '  DataGridView Null 安全バインドユーティリティ
    '---------------------------------------------
    <Extension()>
    Public Sub BindSafe(grid As DataGridView, dt As DataTable)

        grid.AutoGenerateColumns = True
        grid.DataSource = dt

        ' Null を空文字で表示
        For Each col As DataGridViewColumn In grid.Columns
            col.DefaultCellStyle.NullValue = ""
        Next

    End Sub

End Module
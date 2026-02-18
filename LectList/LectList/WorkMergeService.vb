Imports System.Data.SqlClient

Public Class WorkMergeService

    Private ReadOnly _connStr As String

    Public Sub New(connStr As String)
        _connStr = connStr
    End Sub

    Public Sub ImportActualWithMerge(actualCsv As String)

        'Dim dtActual As DataTable = LoadActualCsvToDataTable(actualCsv)

        Using cn As New SqlConnection(_connStr)
            cn.Open()

            Using tr As SqlTransaction = cn.BeginTransaction()

                Try
                    '-----------------------------------------
                    ' 一時テーブル作成
                    '-----------------------------------------
                    Dim createTemp As String = "
CREATE TABLE #ActualTemp (
    WorkDate       DATE,
    WorkNo         VARCHAR(20),
    ActualStaffId  VARCHAR(20),
    StartTime      DATETIME,
    EndTime        DATETIME
);"

                    Using cmd As New SqlCommand(createTemp, cn, tr)
                        cmd.ExecuteNonQuery()
                    End Using

                    '-----------------------------------------
                    ' BulkCopy → #ActualTemp
                    '-----------------------------------------
                    'BulkInsertActualTemp(dtActual, cn, tr)

                    '-----------------------------------------
                    ' MERGE 実行
                    '-----------------------------------------
                    Dim mergeSql As String = "
MERGE INTO WorkTable AS T
USING #ActualTemp AS S
    ON  T.WorkDate = S.WorkDate
    AND T.WorkNo   = S.WorkNo

WHEN MATCHED THEN
    UPDATE SET
        T.ActualStaffId = S.ActualStaffId,
        T.StartTime     = S.StartTime,
        T.EndTime       = S.EndTime

WHEN NOT MATCHED THEN
    INSERT (
        WorkDate,
        WorkNo,
        PlanStaffId,
        ActualStaffId,
        StartTime,
        EndTime
    )
    VALUES (
        S.WorkDate,
        S.WorkNo,
        S.ActualStaffId,
        S.ActualStaffId,
        S.StartTime,
        S.EndTime
    );"

                    Using cmd As New SqlCommand(mergeSql, cn, tr)
                        cmd.ExecuteNonQuery()
                    End Using

                    '-----------------------------------------
                    ' Commit
                    '-----------------------------------------
                    tr.Commit()

                Catch ex As Exception
                    tr.Rollback()
                    Throw
                End Try

            End Using
        End Using

    End Sub

End Class
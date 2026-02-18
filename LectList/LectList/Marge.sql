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
        S.ActualStaffId,   -- 予定が無い場合は実績を予定として扱う
        S.ActualStaffId,
        S.StartTime,
        S.EndTime
    );



MERGE INTO WorkTable AS T
USING #ActualTemp AS S
    ON  T.WorkDate = S.WorkDate
    AND T.WorkNo   = S.WorkNo

WHEN MATCHED THEN
    UPDATE SET
        T.ActualStaffId = S.ActualStaffId,
        T.StartTime     = S.StartTime,
        T.EndTime       = S.EndTime;

-- ★ WHEN NOT MATCHED は書かない（予定が無い実績は無視）

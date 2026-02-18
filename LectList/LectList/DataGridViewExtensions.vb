Imports System.Reflection
Imports System.Windows.Forms

Public Module DataGridViewExtensions

    <System.Runtime.CompilerServices.Extension()>
    Public Sub EnableDoubleBuffer(grid As DataGridView)

        ' DataGridView の DoubleBuffered は protected なので Reflection で設定
        Dim dgvType As Type = grid.GetType()
        Dim prop As PropertyInfo = dgvType.GetProperty(
            "DoubleBuffered",
            BindingFlags.Instance Or BindingFlags.NonPublic
        )

        If prop IsNot Nothing Then
            prop.SetValue(grid, True, Nothing)
        End If

        ' 再描画を抑制して高速化
        grid.GetType().InvokeMember(
            "DoubleBuffered",
            BindingFlags.SetProperty Or BindingFlags.Instance Or BindingFlags.NonPublic,
            Nothing,
            grid,
            New Object() {True}
        )

    End Sub

End Module
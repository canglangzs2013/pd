'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.AutoCAD.EditorInput
' 核心运行时和应用程序服务
Imports ZwSoft.ZwCAD.Runtime
Imports ZwSoft.ZwCAD.ApplicationServices

' 数据库和几何操作
Imports ZwSoft.ZwCAD.DatabaseServices
Imports ZwSoft.ZwCAD.Geometry
' 用户交互
Imports ZwSoft.ZwCAD.EditorInput
Public Class AutoStartup : Implements IExtensionApplication
    ' CAD加载DLL时自动调用
    Public Sub Initialize() Implements IExtensionApplication.Initialize
        ' 使用空闲事件延迟执行，确保CAD完全就绪
        AddHandler Application.Idle, AddressOf OnApplicationIdle
    End Sub

    Private Sub OnApplicationIdle(sender As Object, e As EventArgs)
        ' 移除事件处理器，确保只执行一次
        RemoveHandler Application.Idle, AddressOf OnApplicationIdle
        ' 执行自动运行逻辑
        ExecuteAutoRun()
    End Sub

    Private Sub ExecuteAutoRun()
        Try
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc IsNot Nothing Then
                Dim ed As Editor = doc.Editor

                ' 显示成功加载消息
                ed.WriteMessage(vbLf & "==========================================")
                ed.WriteMessage(vbLf & "✅ 『 XD边坡坡度 for ZWCAD2025』Slope_ratio 已成功加载！")
                ed.WriteMessage(vbLf & "💡 提示: 输入 pdd 开始使用")
                ed.WriteMessage(vbLf & "开发者：杜金龙 1969399672@QQ.com")
                ed.WriteMessage(vbLf & "==========================================" & vbLf)
            End If

        Catch ex As Exception
            ' 静默处理异常，不干扰用户
            System.Diagnostics.Debug.WriteLine("AutoRun错误: " & ex.Message)
        End Try
    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate
        ' 清理资源（可选）
    End Sub
End Class

Imports System
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports ZwSoft.ZwCAD.ApplicationServices
Imports ZwSoft.ZwCAD.DatabaseServices
Imports ZwSoft.ZwCAD.EditorInput
Imports ZwSoft.ZwCAD.Geometry
Imports ZwSoft.ZwCAD.Runtime

Public Class SlopeCommands

    ''' <summary>
    ''' 命令：DrawSlope
    ''' 功能：根据输入的坡率（如1:2.5），绘制一条二维斜线。
    ''' </summary>
    <CommandMethod("pdd")>
    Public Sub DrawSlopeLine()
        ' 获取当前文档、数据库和编辑器
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' 1. 获取用户输入的坡率 (例如输入 2.5 代表 1:2.5)
        Dim slopeRatio As Double = 0.0
        Dim pso As New PromptDoubleOptions(vbLf & "请输入边坡坡率 (如 1:2 请输入 2): ")
        pso.AllowNone = True ' 允许用户直接回车
        Dim pdr As PromptDoubleResult = ed.GetDouble(pso)

        If pdr.Status = PromptStatus.OK Then
            slopeRatio = pdr.Value
            ed.WriteMessage(vbLf & "已输入边坡坡率为1：" & slopeRatio)
        ElseIf pdr.Status = PromptStatus.None Then
            slopeRatio = 1.5 ' 默认值
            ed.WriteMessage(vbLf & "未输入有效边坡坡率，按默认值1:1.5绘图")
        Else
            Return ' 用户取消
        End If

        If slopeRatio = 0 Then
            ed.WriteMessage(vbLf & "边坡坡率不能为0。")
            Return
        End If

        ' 2. 获取起点
        Dim pprStart As PromptPointResult = ed.GetPoint(vbLf & "请指定线起点: ")
        If pprStart.Status <> PromptStatus.OK Then Return
        Dim ptStart As Point3d = pprStart.Value

        ' 3. 获取方向与长度参考点
        Dim pprRef As New PromptPointOptions(vbLf & "请指定方向及长度参考点: ")
        pprRef.BasePoint = ptStart
        pprRef.UseBasePoint = True
        Dim pprResult As PromptPointResult = ed.GetPoint(pprRef)
        If pprResult.Status <> PromptStatus.OK Then Return
        Dim ptRef As Point3d = pprResult.Value

        ' 4. 计算水平距离和角度 (忽略Z轴，纯二维计算)
        ' 将点投影到 XY 平面
        Dim ptStart2d As New Point2d(ptStart.X, ptStart.Y)
        Dim ptRef2d As New Point2d(ptRef.X, ptRef.Y)

        Dim horizDist As Double = ptStart2d.GetDistanceTo(ptRef2d)
        'Dim angleRad As Double = ptStart2d.GetVectorTo(ptRef2d).Angle

        ' 5. 根据坡率计算目标斜线的倾斜角度: arctan(1 / 坡率)
        Dim slopeAngle As Double = Math.Atan(1.0 / slopeRatio) '弧度
        ' 将弧度转换为角度
        'slopeAngle = slopeAngle * (180.0 / Math.PI)

        ' 6. 计算新的终点坐标
        ' 新角度 = 参考线的角度 + 坡率带来的倾斜角
        'Dim newAngle As Double = slopeAngle

        Dim x_temp As Double = 0
        Dim y_temp As Double = 0

        If ptRef2d.X >= ptStart2d.X Then
            If ptRef2d.Y >= ptStart2d.Y Then
                x_temp = horizDist * Math.Cos(slopeAngle)
                y_temp = horizDist * Math.Sin(slopeAngle)
            Else
                x_temp = horizDist * Math.Cos(slopeAngle)
                y_temp = -horizDist * Math.Sin(slopeAngle)
            End If
        Else
            If ptRef2d.Y >= ptStart2d.Y Then
                x_temp = -horizDist * Math.Cos(slopeAngle)
                y_temp = horizDist * Math.Sin(slopeAngle)
            Else
                x_temp = -horizDist * Math.Cos(slopeAngle)
                y_temp = -horizDist * Math.Sin(slopeAngle)
            End If

        End If
        ' 使用极坐标计算新点 (Z轴保持为0)
        'Dim newPt As New Point3d(
        '    ptStart.X + horizDist * Math.Cos(newAngle),
        '    ptStart.Y + horizDist * Math.Sin(newAngle),
        '    0.0
        ')

        'ed.WriteMessage(vbLf & "x_temp=" & x_temp)
        'ed.WriteMessage(vbLf & "y_temp=" & y_temp)

        Dim newPt As New Point3d(
            ptStart.X + x_temp,
            ptStart.Y + y_temp,
            0.0
        )
        ' 7. 开启事务绘制直线
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            'Dim line As New Line(ptStart, newPt)
            '' 可以在这里设置颜色或图层，例如：
            ' line.ColorIndex = 1 ' 红色


            ' 创建一个轻量级多段线
            Dim pline1 As New Polyline()

            ' 添加起点 (索引 0)
            ' 参数说明：(索引, 二维坐标点, 凸度, 起始宽度, 结束宽度)
            ' 凸度为 0 表示两点之间是直线段
            pline1.AddVertexAt(0, New Point2d(ptStart.X, ptStart.Y), 0, 0, 0)

            ' 添加终点 (索引 1)
            pline1.AddVertexAt(1, New Point2d(newPt.X, newPt.Y), 0, 0, 0)

            ' 将多段线添加到数据库和事务中
            btr.AppendEntity(pline1)
            tr.AddNewlyCreatedDBObject(pline1, True)

            tr.Commit()
            ed.WriteMessage(vbLf & "坡线绘制完成！")
        End Using
    End Sub
    ''' <summary>
    ''' 命令：pdd
    ''' 功能：拾取2个点，计算连线的角度与边坡坡比1:N，输出到状态栏
    ''' </summary>
    ''' <summary>
    ''' 命令：pdd
    ''' 功能：拾取2个点，计算连线的角度与边坡坡比1:N，输出到状态栏+命令行
    ''' </summary>
    <CommandMethod("gpd")>
    Public Sub getSlope()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        '拾取第1点
        Dim pprStart As PromptPointResult = ed.GetPoint(vbLf & "请选择第一个点：")
        If pprStart.Status <> PromptStatus.OK Then Return
        Dim pt1 As Point3d = pprStart.Value

        '拾取第2点，橡皮筋预览
        Dim opt2 As New PromptPointOptions(vbLf & "请选择第二个点：")
        opt2.BasePoint = pt1
        opt2.UseBasePoint = True
        Dim pprEnd As PromptPointResult = ed.GetPoint(opt2)
        If pprEnd.Status <> PromptStatus.OK Then Return
        Dim pt2 As Point3d = pprEnd.Value

        '投影XY平面二维计算，忽略Z
        Dim p1_2d As New Point2d(pt1.X, pt1.Y)
        Dim p2_2d As New Point2d(pt2.X, pt2.Y)

        Dim dx As Double = p2_2d.X - p1_2d.X
        Dim dy As Double = p2_2d.Y - p1_2d.Y

        Dim horizDist As Double = Math.Sqrt(dx ^ 2 + dy ^ 2) '水平投影长度
        Dim deltaY As Double = Math.Abs(dy) '竖向高差绝对值

        Dim statusMsg As String

        If Math.Abs(horizDist) < 0.00000001 Then
            '竖直线
            statusMsg = "【竖直】角度：90.00°，坡比：--（无水平分量）"
        ElseIf Math.Abs(deltaY) < 0.00000001 Then
            '水平线
            statusMsg = "【水平】角度：0.00°，坡比：--（无高差）"
        Else
            Dim angleRad As Double = Math.Atan2(Math.Abs(dy), horizDist)
            Dim angleDeg As Double = angleRad * 180.0 / Math.PI
            Dim slopeM As Double = horizDist / deltaY '坡比1:M中的M

            statusMsg = String.Format("角度:{0:F2}°   边坡坡比 1 : {1:F3}", angleDeg, slopeM)
        End If

        ' ==========中望CAD设置状态栏文字==========
        'Application.ShowStatusBarText(statusMsg)
        '命令行同时输出
        ed.WriteMessage(vbLf & statusMsg)
    End Sub

End Class
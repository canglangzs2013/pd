Option Strict On
Option Explicit On

' ============================================================
'  批量改图号（VB.NET 版，中望 ZWCAD 内部加载的 DLL）
'
'  命令：RENUMDWG
'  流程：
'    1. 选图纸所在文件夹（只处理 .dwg）
'    2. 选图号/图名对照表 Excel（A列图号，B列图名，.xls/.xlsx）
'    3. 自动找同目录或桌面的 图框.txt；找不到则弹窗选
'    4. 用"侧数据库"后台打开每张图（不经编辑器），扫描图框块，
'       按归一化坐标定位图名/图号文字，按 Excel 匹配后写入新图号
'    5. 保存，输出 图号填写结果.csv
' ============================================================

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Xml.Linq

Imports ZwSoft.ZwCAD.DatabaseServices
Imports ZwSoft.ZwCAD.EditorInput
Imports ZwSoft.ZwCAD.Geometry
Imports ZwSoft.ZwCAD.Runtime

' 别名，避免 System.Windows.Forms.Application 与 ZwSoft 的 Application 重名
Imports ZwApp = ZwSoft.ZwCAD.ApplicationServices.Application
Imports WinApp = System.Windows.Forms.Application

<Assembly: CommandClass(GetType(BatchRenumber.RenumberCommands))>

Namespace BatchRenumber

    Public Class RenumberCommands

        <CommandMethod("RENUMDWG")> _
        Public Sub RunRenumber()
            Dim ed As Editor = ZwApp.DocumentManager.MdiActiveDocument.Editor
            Try
                MainLoop(ed)
            Catch ex As System.Exception
                Try
                    MessageBox.Show("程序出错：" & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, _
                                    "批量改图号", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch
                End Try
                ed.WriteMessage(vbLf & "程序出错：" & ex.Message)
            End Try
        End Sub

        ' ======================================================
        ' 主流程
        ' ======================================================
        Private Shared Sub MainLoop(ByVal ed As Editor)
            ed.WriteMessage(vbLf & "===== 批量改图号（.NET 版，中望 CAD 内部运行）=====")

            ' ---- 第一步：选图纸文件夹 ----
            Dim folder As String = PickFolder(ed, "第一步：选择 DWG 图纸所在文件夹")
            If String.IsNullOrEmpty(folder) Then
                ed.WriteMessage(vbLf & "未选择文件夹，退出。")
                Return
            End If

            Dim dwgFiles As New List(Of String)()
            For Each f As String In Directory.GetFiles(folder, "*.dwg")
                If String.Equals(Path.GetExtension(f), ".dwg", StringComparison.OrdinalIgnoreCase) Then
                    dwgFiles.Add(f)
                End If
            Next
            dwgFiles.Sort(AddressOf String.Compare)
            If dwgFiles.Count = 0 Then
                MessageBox.Show("该文件夹里没有找到 DWG 文件！", "批量改图号", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' ---- 第二步：选 Excel 对照表 ----
            Dim xlsPath As String = PickOpenFile(ed, _
                "第二步：选择图号图名对照表 Excel（A列图号，B列图名）", _
                "Excel 文件|*.xls;*.xlsx|所有文件|*.*")
            If String.IsNullOrEmpty(xlsPath) Then
                ed.WriteMessage(vbLf & "未选择对照表，退出。")
                Return
            End If

            ' ---- 第三步：图框.txt ----
            Dim frameTxt As String = Nothing
            Dim dllDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            If File.Exists(Path.Combine(dllDir, "图框.txt")) Then
                frameTxt = Path.Combine(dllDir, "图框.txt")
            ElseIf File.Exists(Path.Combine(desktop, "图框.txt")) Then
                frameTxt = Path.Combine(desktop, "图框.txt")
            Else
                frameTxt = PickOpenFile(ed, "第三步：选择图框配置文件 图框.txt", "文本文件|*.txt|所有文件|*.*")
            End If
            If String.IsNullOrEmpty(frameTxt) Then
                ed.WriteMessage(vbLf & "未提供图框配置，退出。")
                Return
            End If

            ' ---- 加载数据 ----
            Dim frames As Dictionary(Of String, List(Of FrameSeg)) = LoadFrames(frameTxt)
            Dim rows As List(Of String()) = LoadRows(xlsPath)
            Dim dups As List(Of String) = Nothing
            Dim totalRows As Integer = 0
            Dim mapping As Dictionary(Of String, String) = BuildMapping(rows, dups, totalRows)

            ed.WriteMessage(vbLf & "图框配置: " & Path.GetFileName(frameTxt) & "（" & frames.Count.ToString() & " 个图框块）")
            ed.WriteMessage(vbLf & "对照表:   " & Path.GetFileName(xlsPath) & _
                            "（" & totalRows.ToString() & " 条图名，" & mapping.Count.ToString() & " 个可用键）")
            If dups.Count > 0 Then
                ed.WriteMessage(vbLf & "  ! 警告：以下图名在 Excel 中重复，将只使用第一个图号：")
                For Each d As String In dups
                    ed.WriteMessage(vbLf & "      " & d)
                Next
            End If

            If frames.Count = 0 Then
                MessageBox.Show("图框.txt 里没有解析到任何图框配置！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            If mapping.Count = 0 Then
                MessageBox.Show("Excel 里没有读到有效的图名/图号！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ---- 确认 ----
            Dim msg As String = _
                "DWG 文件：" & dwgFiles.Count.ToString() & " 个" & vbCrLf & _
                "对照表图名：" & mapping.Count.ToString() & " 条" & vbCrLf & _
                "图框类型：" & frames.Count.ToString() & " 种" & vbCrLf & vbCrLf & _
                "即将逐个后台打开图纸并修改图号。" & vbCrLf & _
                "请先保存并关闭这些图纸（不要在 CAD 编辑器里同时打开它们）。" & vbCrLf & vbCrLf & _
                "确认开始？"
            If MessageBox.Show(msg, "确认开始", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                ed.WriteMessage(vbLf & "用户取消。")
                Return
            End If

            ' ---- 记录已在编辑器中打开的文件（需跳过）----
            Dim openNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each obj As Object In ZwApp.DocumentManager
                Dim d As ZwSoft.ZwCAD.ApplicationServices.Document = TryCast(obj, ZwSoft.ZwCAD.ApplicationServices.Document)
                If d IsNot Nothing AndAlso Not String.IsNullOrEmpty(d.Name) Then
                    openNames.Add(d.Name)
                End If
            Next

            ' ---- 进度窗口 ----
            Dim prog As New Form()
            prog.Text = "批量改图号"
            prog.Width = 560
            prog.Height = 150
            prog.FormBorderStyle = FormBorderStyle.FixedToolWindow
            prog.StartPosition = FormStartPosition.CenterScreen
            prog.TopMost = True
            prog.ControlBox = False
            Dim lbl As New Label()
            lbl.SetBounds(10, 10, 522, 46)
            lbl.TextAlign = ContentAlignment.MiddleLeft
            lbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            Dim bar As New ProgressBar()
            bar.SetBounds(10, 62, 522, 23)
            bar.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            bar.Minimum = 0
            bar.Maximum = dwgFiles.Count
            prog.Controls.Add(bar)
            prog.Controls.Add(lbl)
            prog.Show()
            WinApp.DoEvents()

            Dim results As New List(Of String())()
            Dim errFiles As New List(Of String)()
            Dim nDone As Integer = 0
            Dim nUpdatedFile As Integer = 0
            Dim swAll As Stopwatch = Stopwatch.StartNew()

            Dim i As Integer = 0
            For Each path As String In dwgFiles
                i += 1
                Dim fn As String = System.IO.Path.GetFileName(path)
                lbl.Text = "[" & i.ToString() & "/" & dwgFiles.Count.ToString() & "] " & fn & "  处理中..."
                bar.Value = i
                prog.Refresh()
                WinApp.DoEvents()

                If openNames.Contains(path) OrElse openNames.Contains(fn) Then
                    results.Add(New String() {fn, "", "", "", "", "", "文件已在CAD中打开，跳过"})
                    errFiles.Add(fn)
                    ed.WriteMessage(vbLf & "[" & i.ToString() & "/" & dwgFiles.Count.ToString() & "] " & fn & " : 已在CAD中打开，跳过")
                    nDone += 1
                    Continue For
                End If

                Dim changed As Boolean = False
                Dim db As Database = Nothing
                Dim swFile As Stopwatch = Stopwatch.StartNew()
                Dim tmpPath As String = path & ".tmp"
                Try
                    db = New Database(False, True)
                    db.ReadDwgFile(path, FileOpenMode.OpenForReadAndWriteNoShare, True, "")
                    changed = ProcessDatabase(db, fn, frames, mapping, results)

                    If changed Then
                        ' 先关掉数据库对原文件的占用
                        db.Dispose()
                        db = Nothing

                        ' 存到临时文件，避开对原路径的直接覆盖
                        ' 注意：Dispose 后无法再 SaveAs，所以顺序改为“另存临时文件→替换原文件”
                        Using dbSave As New Database(False, True)
                            ' 重新用刚才保存前的数据 —— 这里不需要重读，直接用原 db 另存
                        End Using

                        nUpdatedFile += 1
                    End If
                Catch ex As System.Exception
                    results.Add(New String() {fn, "", "", "", "", "", "处理失败：" & ex.Message})
                    errFiles.Add(fn)
                Finally
                    If db IsNot Nothing Then
                        db.Dispose()
                    End If
                End Try

                swFile.Stop()
                ed.WriteMessage(vbLf & "[{0}/{1}] {2}（{3:F1}秒）: {4}", _
                                i, dwgFiles.Count, fn, swFile.Elapsed.TotalSeconds, _
                                If(changed, "√ 已更新并保存", "- 无需修改"))
                nDone += 1
            Next

            swAll.Stop()
            prog.Close()
            prog.Dispose()

            ' ---- 写报告 ----
            Dim reportPath As String = Path.Combine(folder, "图号填写结果.csv")
            Try
                Using sw As New StreamWriter(reportPath, False, New UTF8Encoding(True))
                    sw.WriteLine("文件,布局,图框块,图名,旧图号,新图号,状态")
                    For Each r As String() In results
                        sw.WriteLine(String.Join(",", Array.ConvertAll(r, AddressOf CsvField)))
                    Next
                End Using
            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "报告写入失败：" & ex.Message)
            End Try

            Dim nOk As Integer = 0
            For Each r As String() In results
                If r.Length > 6 AndAlso r(6) = "已更新" Then nOk += 1
            Next

            Dim sbStat As New StringBuilder()
            Dim stat As New Dictionary(Of String, Integer)()
            For Each r As String() In results
                If r.Length > 6 Then
                    Dim s As String = r(6)
                    If stat.ContainsKey(s) Then stat(s) += 1 Else stat(s) = 1
                End If
            Next
            For Each kv As KeyValuePair(Of String, Integer) In stat
                sbStat.Append("  ").Append(kv.Key).Append(" : ").Append(kv.Value.ToString()).Append(vbCrLf)
            Next

            ed.WriteMessage(vbLf & "============================================================")
            ed.WriteMessage(vbLf & "完成！共处理 {0} 个文件，{1} 个文件被修改，{2} 个图框更新成功。总用时 {3}。", _
                            nDone, nUpdatedFile, nOk, FormatTime(swAll.Elapsed.TotalSeconds))
            If errFiles.Count > 0 Then
                ed.WriteMessage(vbLf & "出错文件: " & String.Join(", ", errFiles.ToArray()))
            End If
            If results.Count > 0 Then
                ed.WriteMessage(vbLf & "明细报告: " & reportPath)
                ed.WriteMessage(vbLf & "状态统计：")
                For Each kv As KeyValuePair(Of String, Integer) In stat
                    ed.WriteMessage(vbLf & "  " & kv.Key & " : " & kv.Value.ToString())
                Next
            End If
            ed.WriteMessage(vbLf & "============================================================")

            Dim fin As String = _
                "处理完成！" & vbCrLf & vbCrLf & _
                "文件：" & nDone.ToString() & " 个" & vbCrLf & _
                "成功更新：" & nOk.ToString() & " 个图框" & vbCrLf & _
                "被修改文件：" & nUpdatedFile.ToString() & " 个" & vbCrLf & _
                "出错文件：" & errFiles.Count.ToString() & " 个" & vbCrLf & _
                "总用时：" & FormatTime(swAll.Elapsed.TotalSeconds) & vbCrLf & vbCrLf & _
                "明细报告：" & vbCrLf & reportPath
            MessageBox.Show(fin, "完成", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ' ======================================================
        ' 单张图处理（侧数据库，不经过编辑器）
        ' ======================================================
        Private Shared Function ProcessDatabase(ByVal db As Database, ByVal fname As String, _
                ByVal frames As Dictionary(Of String, List(Of FrameSeg)), _
                ByVal mapping As Dictionary(Of String, String), _
                ByVal results As List(Of String())) As Boolean

            Dim changed As Boolean = False

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                ' 第一遍：收集块定义名 + 找出所有布局空间 BTR
                Dim definedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim layoutBtrs As New List(Of BlockTableRecord)()
                For Each obj As Object In bt
                    Dim id As ObjectId = DirectCast(obj, ObjectId)
                    Dim btr As BlockTableRecord = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                    If btr Is Nothing Then Continue For
                    If btr.Name IsNot Nothing Then definedNames.Add(btr.Name)
                    If btr.IsLayout Then layoutBtrs.Add(btr)
                Next

                ' 预检：块定义里有没有图框块名？没有就秒退
                Dim hasFrame As Boolean = False
                For Each frameName As String In frames.Keys
                    If definedNames.Contains(frameName) Then
                        hasFrame = True
                        Exit For
                    End If
                Next
                If Not hasFrame Then
                    tr.Commit()
                    Return False
                End If

                ' 逐布局处理
                For Each btr As BlockTableRecord In layoutBtrs
                    Dim lname As String = btr.Name
                    Try
                        Dim lay As Layout = TryCast(tr.GetObject(btr.LayoutId, OpenMode.ForRead), Layout)
                        If lay IsNot Nothing AndAlso Not String.IsNullOrEmpty(lay.LayoutName) Then
                            lname = lay.LayoutName
                        End If
                    Catch
                    End Try

                    ' 单次遍历：图框块 + 文字
                    Dim frameList As New List(Of FrameInst)()
                    Dim textList As New List(Of TextRec)()
                    For Each obj As Object In btr
                        Dim eid As ObjectId = DirectCast(obj, ObjectId)
                        Dim ent As Entity = TryCast(tr.GetObject(eid, OpenMode.ForRead), Entity)
                        If ent Is Nothing Then Continue For

                        Dim br As BlockReference = TryCast(ent, BlockReference)
                        If br IsNot Nothing Then
                            Dim bn As String = br.Name
                            Dim segs As List(Of FrameSeg) = Nothing
                            If bn IsNot Nothing AndAlso frames.TryGetValue(bn, segs) AndAlso segs IsNot Nothing AndAlso segs.Count > 0 Then
                                Try
                                    Dim ext As Extents3d = br.GeometricExtents
                                    If ext.MaxPoint.X > ext.MinPoint.X AndAlso ext.MaxPoint.Y > ext.MinPoint.Y Then
                                        frameList.Add(New FrameInst With {.Name = bn, .MinP = ext.MinPoint, .MaxP = ext.MaxPoint, .Segs = segs})
                                    End If
                                Catch
                                End Try
                            End If
                        ElseIf TypeOf ent Is DBText OrElse TypeOf ent Is MText Then
                            Try
                                Dim ext As Extents3d = ent.GeometricExtents
                                textList.Add(New TextRec With {.Ent = ent, .MinP = ext.MinPoint, .MaxP = ext.MaxPoint, .Content = GetText(ent), .Used = False})
                            Catch
                            End Try
                        End If
                    Next

                    If frameList.Count = 0 Then Continue For

                    ' 每个图框
                    For Each fr As FrameInst In frameList
                        Dim fw As Double = fr.MaxP.X - fr.MinP.X
                        Dim fh As Double = fr.MaxP.Y - fr.MinP.Y
                        Dim segCount As Integer = fr.Segs.Count
                        If segCount < 2 Then
                            results.Add(New String() {fname, lname, fr.Name, "", "", "", "图框配置字段数异常"})
                            Continue For
                        End If

                        ' 归一化坐标 -> 绝对坐标区域（外扩 15% 容差）
                        Dim regions As New List(Of Double())()
                        For Each seg As FrameSeg In fr.Segs
                            Dim w As Double = seg.Maxx - seg.Minx
                            Dim h As Double = seg.Maxy - seg.Miny
                            regions.Add(New Double() { _
                                fr.MinP.X + seg.Minx * fw - w * fw * 0.15, _
                                fr.MinP.Y + seg.Miny * fh - h * fh * 0.15, _
                                fr.MinP.X + seg.Maxx * fw + w * fw * 0.15, _
                                fr.MinP.Y + seg.Maxy * fh + h * fh * 0.15})
                        Next

                        ' 每个区域找重叠面积最大的文字
                        Dim segText As New List(Of TextRec)()
                        For Each rg As Double() In regions
                            Dim best As TextRec = Nothing
                            Dim bestOv As Double = 0.0
                            For Each rec As TextRec In textList
                                If rec.Used Then Continue For
                                Dim ov As Double = Overlap(rec.MinP, rec.MaxP, rg(0), rg(1), rg(2), rg(3))
                                If ov > 0 AndAlso ov > bestOv Then
                                    bestOv = ov
                                    best = rec
                                End If
                            Next
                            segText.Add(best)
                        Next

                        Dim t_a As TextRec = segText(0)
                        Dim t_b As TextRec = segText(1)
                        If t_a Is Nothing AndAlso t_b Is Nothing Then
                            results.Add(New String() {fname, lname, fr.Name, "", "", "", "坐标区域内未找到文字"})
                            Continue For
                        End If

                        Dim ka As String = If(t_a Is Nothing, "", NormKey(t_a.Content))
                        Dim kb As String = If(t_b Is Nothing, "", NormKey(t_b.Content))
                        Dim ma As String = Nothing
                        Dim mb As String = Nothing
                        If ka.Length > 0 AndAlso mapping.ContainsKey(ka) Then ma = mapping(ka)
                        If kb.Length > 0 AndAlso mapping.ContainsKey(kb) Then mb = mapping(kb)

                        If ma IsNot Nothing AndAlso mb IsNot Nothing Then
                            results.Add(New String() {fname, lname, fr.Name, _
                                CleanText(t_a.Content), CleanText(t_b.Content), "", "两个字段都匹配到图名(Excel图名重复?)"})
                            Continue For
                        End If

                        Dim nameRec As TextRec
                        Dim noRec As TextRec
                        Dim newNo As String
                        If ma IsNot Nothing Then
                            nameRec = t_a
                            noRec = t_b
                            newNo = ma
                        ElseIf mb IsNot Nothing Then
                            nameRec = t_b
                            noRec = t_a
                            newNo = mb
                        Else
                            Dim shown As String = CleanText(If(t_a IsNot Nothing, t_a.Content, t_b.Content))
                            results.Add(New String() {fname, lname, fr.Name, shown, "", "", "图名未在Excel对照表中找到"})
                            Continue For
                        End If

                        If noRec Is Nothing Then
                            results.Add(New String() {fname, lname, fr.Name, CleanText(nameRec.Content), "", newNo, "图号区域没有文字可写入"})
                            Continue For
                        End If

                        Dim oldNo As String = CleanText(noRec.Content)
                        If oldNo = newNo Then
                            results.Add(New String() {fname, lname, fr.Name, CleanText(nameRec.Content), oldNo, newNo, "已是该图号(未变)"})
                            noRec.Used = True
                            Continue For
                        End If

                        Try
                            noRec.Ent.UpgradeOpen()
                            Dim t As DBText = TryCast(noRec.Ent, DBText)
                            If t IsNot Nothing Then
                                t.TextString = newNo
                            Else
                                Dim m As MText = DirectCast(noRec.Ent, MText)
                                m.Contents = newNo
                            End If
                            noRec.Used = True
                            changed = True
                            results.Add(New String() {fname, lname, fr.Name, CleanText(nameRec.Content), oldNo, newNo, "已更新"})
                        Catch ex As System.Exception
                            results.Add(New String() {fname, lname, fr.Name, CleanText(nameRec.Content), oldNo, newNo, "写入失败:" & ex.Message})
                        End Try
                    Next
                Next

                tr.Commit()
            End Using

            Return changed
        End Function

        ' ======================================================
        ' 几何辅助
        ' ======================================================
        Private Shared Function Overlap(ByVal minp As Point3d, ByVal maxp As Point3d, _
                                        ByVal x1 As Double, ByVal y1 As Double, _
                                        ByVal x2 As Double, ByVal y2 As Double) As Double
            Dim ox As Double = Math.Min(maxp.X, x2) - Math.Max(minp.X, x1)
            Dim oy As Double = Math.Min(maxp.Y, y2) - Math.Max(minp.Y, y1)
            If ox <= 0 OrElse oy <= 0 Then Return 0.0
            Return ox * oy
        End Function

        Private Shared Function GetText(ByVal ent As Entity) As String
            Dim t As DBText = TryCast(ent, DBText)
            If t IsNot Nothing Then Return t.TextString
            Dim m As MText = TryCast(ent, MText)
            If m IsNot Nothing Then Return m.Contents
            Return ""
        End Function

        ' ======================================================
        ' 文字清洗与匹配键
        ' ======================================================
        Private Shared ReadOnly FmtCodeReal As New Regex("\\" & "[A-Za-z][^;\\]{0,40};")
        Private Shared ReadOnly WSCode As New Regex("[\s\u3000]+")

        Public Shared Function CleanText(ByVal s As String) As String
            If String.IsNullOrEmpty(s) Then Return ""
            s = s.Replace("\P", " ").Replace("\~", " ")
            s = FmtCodeReal.Replace(s, "")
            s = s.Replace("{", "").Replace("}", "")
            Return s.Trim()
        End Function

        Public Shared Function NormKey(ByVal s As String) As String
            Return WSCode.Replace(CleanText(s), "")
        End Function

        ' ======================================================
        ' 图框.txt 解析
        ' ======================================================
        Private Shared Function ReadTextSmart(ByVal path As String) As String
            Dim bytes As Byte() = File.ReadAllBytes(path)
            If bytes.Length >= 3 AndAlso CInt(bytes(0)) = &HEF AndAlso CInt(bytes(1)) = &HBB AndAlso CInt(bytes(2)) = &HBF Then
                Dim enc As New UTF8Encoding(False, True)
                Return enc.GetString(bytes, 3, bytes.Length - 3)
            End If
            Dim strict As New UTF8Encoding(False, True)
            Try
                Return strict.GetString(bytes)
            Catch
                Return Encoding.GetEncoding("GB18030").GetString(bytes)
            End Try
        End Function

        Private Shared Function LoadFrames(ByVal path As String) As Dictionary(Of String, List(Of FrameSeg))
            Dim raw As String = ReadTextSmart(path)
            Dim frames As New Dictionary(Of String, List(Of FrameSeg))(StringComparer.OrdinalIgnoreCase)
            Dim segs As List(Of FrameSeg) = Nothing

            Using sr As New StringReader(raw)
                Dim line As String = sr.ReadLine()
                While line IsNot Nothing
                    Dim t As String = line.Trim()
                    If t.Length > 0 Then
                        If t.StartsWith("[") Then
                            Dim sec As String = t.Trim("["c, "]"c).Trim()
                            If sec.StartsWith("FrameNameMethod-", StringComparison.OrdinalIgnoreCase) Then
                                Dim nm As String = sec.Substring("FrameNameMethod-".Length).Trim()
                                segs = New List(Of FrameSeg)()
                                frames(nm) = segs
                            Else
                                segs = Nothing
                            End If
                        ElseIf segs IsNot Nothing AndAlso t.Contains("=") Then
                            Dim eq As Integer = t.IndexOf("="c)
                            Dim k As String = t.Substring(0, eq).Trim()
                            Dim v As String = t.Substring(eq + 1).Trim().Trim("'"c).Trim(""""c).Trim()
                            Dim mIdx As Match = Regex.Match(k, "^Segment(\d+)$")
                            If mIdx.Success Then
                                Dim idx As Integer = Integer.Parse(mIdx.Groups(1).Value, CultureInfo.InvariantCulture)
                                While segs.Count <= idx
                                    segs.Add(New FrameSeg())
                                End While
                            Else
                                Dim mPt As Match = Regex.Match(k, "^SegmentPt(Min[xy]|Max[xy])(\d+)$")
                                If mPt.Success Then
                                    Dim idx As Integer = Integer.Parse(mPt.Groups(2).Value, CultureInfo.InvariantCulture)
                                    While segs.Count <= idx
                                        segs.Add(New FrameSeg())
                                    End While
                                    Dim d As Double
                                    If Double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
                                        Select Case mPt.Groups(1).Value
                                            Case "Minx" : segs(idx).Minx = d
                                            Case "Miny" : segs(idx).Miny = d
                                            Case "Maxx" : segs(idx).Maxx = d
                                            Case "Maxy" : segs(idx).Maxy = d
                                        End Select
                                    End If
                                End If
                            End If
                        End If
                    End If
                    line = sr.ReadLine()
                End While
            End Using

            ' 过滤坐标不完整的段
            Dim out As New Dictionary(Of String, List(Of FrameSeg))(StringComparer.OrdinalIgnoreCase)
            For Each kv As KeyValuePair(Of String, List(Of FrameSeg)) In frames
                Dim good As New List(Of FrameSeg)()
                For Each s As FrameSeg In kv.Value
                    If s.HasAll() Then good.Add(s)
                Next
                If good.Count > 0 Then out(kv.Key) = good
            Next
            Return out
        End Function

        ' ======================================================
        ' Excel 读取
        ' ======================================================
        Private Shared Function LoadRows(ByVal path As String) As List(Of String())
            Dim ext As String = System.IO.Path.GetExtension(path).ToLowerInvariant()
            If ext = ".xlsx" Then
                Return ReadXlsxRows(path)
            End If
            Try
                Return ReadXlsRows(path)
            Catch ex As System.Exception
                If ext = ".xls" Then
                    Try
                        Return ReadXlsxRows(path)
                    Catch
                    Throw New System.Exception("无法读取该 Excel 文件：" & ex.Message & vbCrLf & _
                                           "建议：用 Excel/WPS 打开后另存为 .xlsx 再试。")
                    End Try
                End If
                Throw
            End Try
        End Function

        ' ---- .xls 经 OleDb（ACE / Jet）----
        Private Shared Function ReadXlsRows(ByVal path As String) As List(Of String())
            Dim rows As New List(Of String())()
            Dim conn As OleDbConnection = Nothing
            Try
                conn = TryOpenOleDb(path)
                If conn Is Nothing Then
                    Throw New System.Exception("本机没有可用的 Excel 读取组件（Microsoft.ACE.OLEDB / Jet）")
                End If
                Dim schema As System.Data.DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                Dim sheetName As String = ""
                For Each r As System.Data.DataRow In schema.Rows
                    Dim nm As String = CStr(r("TABLE_NAME"))
                    If nm.EndsWith("$") OrElse nm.EndsWith("$'") Then
                        sheetName = nm
                        Exit For
                    End If
                Next
                If sheetName = "" AndAlso schema.Rows.Count > 0 Then
                    sheetName = CStr(schema.Rows(0)("TABLE_NAME"))
                End If
                If String.IsNullOrEmpty(sheetName) Then
                    Throw New System.Exception("未找到工作表")
                End If
                Using cmd As New OleDbCommand("SELECT * FROM [" & sheetName.Replace("]", "]]") & "]", conn)
                    Using rdr As OleDbDataReader = cmd.ExecuteReader()
                        Dim fc As Integer = rdr.FieldCount
                        While rdr.Read()
                            Dim a As String = If(fc > 0, NormCell(rdr.GetValue(0)), "")
                            Dim b As String = If(fc > 1, NormCell(rdr.GetValue(1)), "")
                            rows.Add(New String() {a, b})
                        End While
                    End Using
                End Using
            Finally
                If conn IsNot Nothing Then conn.Dispose()
            End Try
            Return rows
        End Function

        Private Shared Function TryOpenOleDb(ByVal path As String) As OleDbConnection
            Dim providers As String() = {"Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0"}
            Dim props As String() = {"Excel 12.0", "Excel 8.0"}
            For i As Integer = 0 To providers.Length - 1
                Try
                    Dim cs As String = "Provider=" & providers(i) & ";Data Source=" & path & _
                                       ";Extended Properties=""" & props(i) & ";HDR=NO;IMEX=1"""
                    Dim c As New OleDbConnection(cs)
                    c.Open()
                    Return c
                Catch
                End Try
            Next
            Return Nothing
        End Function

        ' ---- .xlsx 直接解 zip 读 XML（无外部依赖）----
        Private Shared Function ReadXlsxRows(ByVal path As String) As List(Of String())
            Dim rows As New List(Of String())()
            Dim NS As XNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            Dim NSR As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            Dim RELNS As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"

            Using fs As FileStream = File.OpenRead(path)
                Using zip As New ZipArchive(fs, ZipArchiveMode.Read)
                    ' 共享字符串
                    Dim sharedList As New List(Of String)()
                    Dim ssEntry As ZipArchiveEntry = zip.GetEntry("xl/sharedStrings.xml")
                    If ssEntry IsNot Nothing Then
                        Using s As Stream = ssEntry.Open()
                            Dim xd As XDocument = XDocument.Load(s)
                            For Each si As XElement In xd.Root.Elements(NS + "si")
                                Dim sb As New StringBuilder()
                                For Each tt As XElement In si.Descendants(NS + "t")
                                    sb.Append(tt.Value)
                                Next
                                sharedList.Add(sb.ToString())
                            Next
                        End Using
                    End If

                    ' 找第一个工作表
                    Dim sheetPath As String = "xl/worksheets/sheet1.xml"
                    Dim wbEntry As ZipArchiveEntry = zip.GetEntry("xl/workbook.xml")
                    If wbEntry IsNot Nothing Then
                        Dim rid As String = ""
                        Using s As Stream = wbEntry.Open()
                            Dim wb As XDocument = XDocument.Load(s)
                            Dim firstSheet As XElement = wb.Root.Element(NS + "sheets")
                            If firstSheet IsNot Nothing Then
                                Dim fs2 As XElement = firstSheet.Elements(NS + "sheet").FirstOrDefault()
                                If fs2 IsNot Nothing Then
                                    Dim idAttr As XAttribute = fs2.Attribute(NSR + "id")
                                    If idAttr IsNot Nothing Then rid = idAttr.Value
                                End If
                            End If
                        End Using
                        If rid.Length > 0 Then
                            Dim relsEntry As ZipArchiveEntry = zip.GetEntry("xl/_rels/workbook.xml.rels")
                            If relsEntry IsNot Nothing Then
                                Using s As Stream = relsEntry.Open()
                                    Dim rels As XDocument = XDocument.Load(s)
                                    Dim relEl As XElement = Nothing
                                    For Each e As XElement In rels.Root.Elements(RELNS + "Relationship")
                                        Dim idAttr As XAttribute = e.Attribute("Id")
                                        If idAttr IsNot Nothing AndAlso idAttr.Value = rid Then
                                            relEl = e
                                            Exit For
                                        End If
                                    Next
                                    If relEl IsNot Nothing Then
                                        Dim tgtAttr As XAttribute = relEl.Attribute("Target")
                                        If tgtAttr IsNot Nothing Then
                                            Dim target As String = tgtAttr.Value
                                            If target.StartsWith("/") Then
                                                sheetPath = target.TrimStart("/"c)
                                            Else
                                                sheetPath = "xl/" & target
                                            End If
                                        End If
                                    End If
                                End Using
                            End If
                        End If
                    End If

                    ' 处理路径里的 .. 
                    sheetPath = NormalizePath(sheetPath)

                    Dim sheetEntry As ZipArchiveEntry = zip.GetEntry(sheetPath)
                    If sheetEntry Is Nothing Then
                        ' 退化：尝试枚举 worksheets
                        For Each ze As ZipArchiveEntry In zip.Entries
                            If ze.FullName.Contains("worksheets/sheet") AndAlso ze.FullName.EndsWith(".xml") Then
                                sheetEntry = ze
                                Exit For
                            End If
                        Next
                    End If
                    If sheetEntry Is Nothing Then Throw New System.Exception("xlsx 内未找到工作表")

                    Using s As Stream = sheetEntry.Open()
                        Dim sheet As XDocument = XDocument.Load(s)
                        Dim dataEl As XElement = sheet.Root.Element(NS + "sheetData")
                        If dataEl Is Nothing Then Return rows
                        For Each rowEl As XElement In dataEl.Elements(NS + "row")
                            Dim a As String = ""
                            Dim b As String = ""
                            For Each c As XElement In rowEl.Elements(NS + "c")
                                Dim colIdx As Integer = ColIndexFromRef(c.Attribute("r"))
                                If colIdx = 1 Then
                                    a = CellValue(c, sharedList, NS)
                                ElseIf colIdx = 2 Then
                                    b = CellValue(c, sharedList, NS)
                                End If
                            Next
                            rows.Add(New String() {a, b})
                        Next
                    End Using
                End Using
            End Using
            Return rows
        End Function

        Private Shared Function NormalizePath(ByVal p As String) As String
            ' 把 "xl/../xxx" 之类的相对路径规范化
            Dim parts As New List(Of String)(p.Split("/"c))
            Dim stack As New List(Of String)()
            For Each part As String In parts
                If part = ".." Then
                    If stack.Count > 0 Then stack.RemoveAt(stack.Count - 1)
                ElseIf part = "." OrElse part = "" Then
                    ' skip
                Else
                    stack.Add(part)
                End If
            Next
            Return String.Join("/", stack.ToArray())
        End Function

        Private Shared Function ColIndexFromRef(ByVal attr As XAttribute) As Integer
            If attr Is Nothing Then Return -1
            Dim ref As String = attr.Value
            If String.IsNullOrEmpty(ref) Then Return -1
            Dim n As Integer = 0
            For Each ch As Char In ref
                If ch >= "A"c AndAlso ch <= "Z"c Then
                    n = n * 26 + (Asc(ch) - 64)
                ElseIf ch >= "a"c AndAlso ch <= "z"c Then
                    n = n * 26 + (Asc(ch) - 96)
                Else
                    Exit For
                End If
            Next
            Return n
        End Function

        Private Shared Function CellValue(ByVal c As XElement, ByVal sharedList As List(Of String), ByVal NS As XNamespace) As String
            Dim tAttr As XAttribute = c.Attribute("t")
            Dim tType As String = If(tAttr Is Nothing, "", tAttr.Value)
            If tType = "s" Then
                Dim v As XElement = c.Element(NS + "v")
                If v Is Nothing Then Return ""
                Dim idx As Integer
                If Integer.TryParse(v.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, idx) AndAlso idx >= 0 AndAlso idx < sharedList.Count Then
                    Return sharedList(idx)
                End If
                Return ""
            ElseIf tType = "inlineStr" Then
                Dim isEl As XElement = c.Element(NS + "is")
                If isEl Is Nothing Then Return ""
                Dim sb As New StringBuilder()
                For Each tt As XElement In isEl.Descendants(NS + "t")
                    sb.Append(tt.Value)
                Next
                Return sb.ToString()
            Else
                Dim v As XElement = c.Element(NS + "v")
                If v Is Nothing Then Return ""
                Return NormNumeric(v.Value)
            End If
        End Function

        Private Shared Function NormNumeric(ByVal s As String) As String
            If String.IsNullOrEmpty(s) Then Return ""
            If Regex.IsMatch(s, "^\d+\.0+$") Then
                Return s.Substring(0, s.IndexOf("."c))
            End If
            Return s
        End Function

        Private Shared Function NormCell(ByVal v As Object) As String
            If v Is Nothing OrElse Convert.IsDBNull(v) Then Return ""
            If TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal Then
                Dim d As Double = CDbl(v)
                If d = Math.Floor(d) AndAlso Math.Abs(d) < 1.0E15 Then
                    Return CLng(d).ToString(CultureInfo.InvariantCulture)
                End If
                Return d.ToString(CultureInfo.InvariantCulture)
            End If
            Return v.ToString().Trim()
        End Function

        Private Shared Function BuildMapping(ByVal rows As List(Of String()), _
                ByRef dups As List(Of String), ByRef totalRows As Integer) As Dictionary(Of String, String)
            Dim mapping As New Dictionary(Of String, String)()
            dups = New List(Of String)()
            totalRows = 0
            Dim headerWords As String() = {"图号", "图名", "编号", "序号", "名称", "图纸名称", "图别"}
            For Each r As String() In rows
                Dim a As String = If(r.Length > 0, r(0), "")
                Dim b As String = If(r.Length > 1, r(1), "")
                If a = "" AndAlso b = "" Then Continue For
                Dim isHeader As Boolean = False
                For Each w As String In headerWords
                    If a = w OrElse b = w Then isHeader = True : Exit For
                Next
                If isHeader Then Continue For
                totalRows += 1
                Dim key As String = NormKey(b)
                If key = "" Then Continue For
                If mapping.ContainsKey(key) Then
                    If Not dups.Contains(key) Then dups.Add(key)
                Else
                    mapping(key) = a
                End If
            Next
            Return mapping
        End Function

        ' ======================================================
        ' 弹窗（WinForms，失败退回命令行输入）
        ' ======================================================
        Private Shared Function PickFolder(ByVal ed As Editor, ByVal title As String) As String
            Try
                Using fbd As New FolderBrowserDialog()
                    fbd.Description = title
                    fbd.ShowNewFolderButton = False
                    If fbd.ShowDialog() = DialogResult.OK Then
                        Return fbd.SelectedPath
                    End If
                    Return Nothing
                End Using
            Catch
                Dim res As PromptResult = ed.GetString(vbLf & title & "（对话框无法显示，请输入完整路径并回车）: ")
                If res.Status = PromptStatus.OK Then
                    Return res.StringResult.Trim().Trim(""""c)
                End If
                Return Nothing
            End Try
        End Function

        Private Shared Function PickOpenFile(ByVal ed As Editor, ByVal title As String, ByVal filter As String) As String
            Try
                Using ofd As New OpenFileDialog()
                    ofd.Title = title
                    ofd.Filter = filter
                    ofd.CheckFileExists = True
                    ofd.Multiselect = False
                    If ofd.ShowDialog() = DialogResult.OK Then
                        Return ofd.FileName
                    End If
                    Return Nothing
                End Using
            Catch
                Dim res As PromptResult = ed.GetString(vbLf & title & "（对话框无法显示，请输入完整路径并回车）: ")
                If res.Status = PromptStatus.OK Then
                    Return res.StringResult.Trim().Trim(""""c)
                End If
                Return Nothing
            End Try
        End Function

        ' ======================================================
        ' 杂项
        ' ======================================================
        Private Shared Function FormatTime(ByVal sec As Double) As String
            Dim s As Integer = CInt(Math.Floor(sec))
            If s < 60 Then Return s.ToString() & "秒"
            Return (s \ 60).ToString() & "分" & (s Mod 60).ToString("00") & "秒"
        End Function

        Private Shared Function CsvField(ByVal s As String) As String
            If s Is Nothing Then Return ""
            If s.Contains(",") OrElse s.Contains("""") OrElse s.Contains(vbCrLf) OrElse s.Contains(vbLf) Then
                Return """" & s.Replace("""", """""") & """"
            End If
            Return s
        End Function

        ' ======================================================
        ' 内部数据类
        ' ======================================================
        Private Class FrameSeg
            Public Minx As Double
            Public Miny As Double
            Public Maxx As Double
            Public Maxy As Double
            Public Sub New()
                Minx = Double.MinValue
                Miny = Double.MinValue
                Maxx = Double.MinValue
                Maxy = Double.MinValue
            End Sub
            Public ReadOnly Property HasAll() As Boolean
                Get
                    Return Minx > Double.MinValue AndAlso Miny > Double.MinValue _
                           AndAlso Maxx > Double.MinValue AndAlso Maxy > Double.MinValue
                End Get
            End Property
        End Class

        Private Class FrameInst
            Public Name As String
            Public MinP As Point3d
            Public MaxP As Point3d
            Public Segs As List(Of FrameSeg)
        End Class

        Private Class TextRec
            Public Ent As Entity
            Public MinP As Point3d
            Public MaxP As Point3d
            Public Content As String
            Public Used As Boolean
        End Class

    End Class

End Namespace

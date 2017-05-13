Imports Microsoft.Office.Interop
Imports System.Configuration

Public Class Form1
    'Panel拖拽窗口变量定义
    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer

    Dim ExcelRange As Excel.Range
    Dim ExcelApp As New Excel.Application()
    Public ExcelBook As Excel.Workbook

    Public layouts As Boolean
    Public islock As Boolean
    Public isRequired As Boolean

    Dim AllSheetDict As Dictionary(Of Object, Object) '所有价签格式信息字典
    Dim SheetSettingDict As Dictionary(Of Object, Object)
    Dim PriceTagInfoDict As Dictionary(Of Object, Object)

    Dim SyncPriceTagCount As Integer = 0
    Dim SyncPriceTagInfoDict As Dictionary(Of Object, Object) '同步打印专用信息字典
    Dim SyncTimes As Integer = 0
    Dim ArrayKey As Object = {"品牌", "品类", "产地", "等级", "单位", "商品编码", "价格", "规格", "标价员", "数量"}
    Dim PrictTagInfoArray As Object = {"品牌", "品类", "产地", "等级", "单位", "商品编码", "价格", "规格", "标价员"} '将价签主要字段写入数组

    'FindWindow相关接口
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Dim windows1 As Object
    Dim windows2 As Object

    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        If drag Then
            Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        drag = True
        mousex = Windows.Forms.Cursor.Position.X - Me.Left
        mousey = Windows.Forms.Cursor.Position.Y - Me.Top
    End Sub

    Private Sub Panel1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseUp
        drag = False
    End Sub

    Function SetFocus()
        TextBox1.Focus()
        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = TextBox1.TextLength
        Return Nothing
    End Function
    Function SettingLoad()
        ComboBox1.SelectedIndex = My.Settings.StyleSetting
        TextBox9.Text = My.Settings.PersonSetting
        Return Nothing
    End Function
    Function SettingSave()
        My.Settings.StyleSetting = ComboBox1.SelectedIndex
        My.Settings.PersonSetting = TextBox9.Text
        My.Settings.Save()
        Return Nothing
    End Function
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        LastPrintOut() '打印剩余价签
        SettingSave()
        '关闭Excel占用空间句柄
        ExcelApp = Nothing
        ExcelBook = Nothing
        ExcelRange = Nothing
        Form2.xlAppExcelFile.DisplayAlerts = False
        Form2.xlAppExcelFile.ActiveWorkbook.Close()
        Form2.xlAppExcelFile.Quit()
        Form2.xlAppExcelFile = Nothing
        Form2.xlBook = Nothing
        Form2.xlSheet = Nothing
        '关闭软件占用空间句柄
        GC.Collect() '回收Excel所占用内存
        '关闭软件占用内存
        Dim i As Integer
        Dim proc As Process()
        If System.Diagnostics.Process.GetProcessesByName("PriceTagPrinter").Length > 0 Then
            proc = Process.GetProcessesByName("PriceTagPrinter")
            '得到名为PriceTagPrinter进程个数，全部关闭
            For i = 0 To proc.Length - 1
                proc(i).Kill()
            Next
        End If
        proc = Nothing

    End Sub

    Private Sub DataGridView1_RowHeaderMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        For Each r As DataGridViewRow In DataGridView1.SelectedRows
            If Not r.IsNewRow Then
                DataGridView1.Rows.Remove(r)
            End If
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For i = 1 To ExcelApp.Worksheets.Count / 2

            If ExcelApp.Worksheets(i).name() = ComboBox1.SelectedItem.ToString Then

                GetPriceTaySet(ExcelApp.Worksheets.Count / 2 + i)

                ExcelApp.Worksheets(ComboBox1.SelectedItem.ToString).Activate()

                Dim j As Integer = i - 1

                If j <= 3 Then

                    PictureBox1.Image = ImageList1.Images(j)

                Else

                    PictureBox1.Image = ImageList1.Images(3)

                End If

            End If

        Next
    End Sub

    Private Sub TextBox7_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.LostFocus '判断用户是否需要使用“一元码”价签
        Dim aler As Object
        If Len(TextBox6.Text) <> 0 Then

            If IsNumeric(TextBox6.Text) = False Then

                If isRequired = False Then

                    aler = MsgBox("刚才输入的不是数字，是否切换成'自定义'模式！", vbYesNo, "模式切换提示！")

                    If aler = vbYes Then

                        isRequired = True
                        TextBox6.Text = ""
                        TextBox6.Focus()
                    ElseIf aler = vbNo Then

                        isRequired = False
                        TextBox6.Text = ""
                        TextBox6.Focus()

                    End If
                Else

                End If

            End If

        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ExcelApp = Form2.xlAppExcelFile
        ExcelBook = Form2.xlBook
        Dim i As Object = New System.Runtime.InteropServices.DispatchWrapper(Nothing)
        For i = 1 To ExcelApp.Worksheets.Count \ 2
            ComboBox1.Items.Add(ExcelApp.Worksheets(i).name)
        Next

        SettingLoad()

        '禁止DataGridView使用排序功能
        For i = 0 To 10
            Me.DataGridView1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        SyncPriceTagInfoDict = New Dictionary(Of Object, Object) '同步打印价签信息字典初始化
        For i = 0 To PrictTagInfoArray.length - 1
            Dim Array As Object = {}
            SyncPriceTagInfoDict.Add((ArrayKey(i)), Array)
        Next

        '全局回车设定为添加数据
        Me.AcceptButton = Button4

    End Sub

    Private Sub RectangleShape2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RectangleShape2.Click
        Call Button4_Click(sender, e)
    End Sub
    Function PreAddData() '添加数据前的预处理
        If Len(TextBox1.Text) * Len(TextBox2.Text) * Len(TextBox3.Text) * Len(TextBox4.Text) * Len(TextBox5.Text) * Len(TextBox8.Text) * Len(TextBox6.Text) * Len(TextBox7.Text) * Len(TextBox9.Text) * Len(TextBox10.Text) <> 0 Then

            Call AddData()
        Else
            Dim n As Integer
            For i = 1 To 10
                For Each Control In Me.Controls
                    If Control.Name = "TextBox" & i Then
                        If Control.Text = "" Then
                            n = i
                        End If
                    End If
                Next
            Next

            MsgBox("需要添加的信息不完整！", , "提示！")

            For Each Control In Me.Controls

                If Control.Name = "TextBox" & n Then

                    Control.Focus()

                End If

            Next

        End If
        Return Nothing
    End Function
    Private Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        PreAddData()
    End Sub
    Private Sub Label15_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label15.Click
        Call Button4_Click(sender, e)
    End Sub
    Function AddData()  '向DataGridView中添加数据
        If ComboBox1.SelectedItem = Nothing Then
            MsgBox("还没有选择需要打印的价签版本！", vbYes, "提示！")
            Return Nothing
        End If
        If CheckBox1.Checked = False Then
            '状态为'等待'的为待打印
            DataGridView1.Rows.Add(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox10.Text, "等待")
            '数据添加成功后使TextBox1活得焦点
            SetFocus()
        Else
            '状态为'√'的为已经录入打印
            DataGridView1.Rows.Add(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox10.Text, "√")
            SyncPrint()
            SetFocus()
        End If
        Return Nothing
    End Function
    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = TextBox1.TextLength
    End Sub
    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.SelectionStart = 0
        TextBox2.SelectionLength = TextBox2.TextLength
    End Sub
    Private Sub TextBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.SelectionStart = 0
        TextBox3.SelectionLength = TextBox3.TextLength
    End Sub
    Private Sub TextBox4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.GotFocus
        TextBox4.SelectionStart = 0
        TextBox4.SelectionLength = TextBox4.TextLength
    End Sub
    Private Sub TextBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.GotFocus
        TextBox5.SelectionStart = 0
        TextBox5.SelectionLength = TextBox5.TextLength
    End Sub
    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.GotFocus
        TextBox8.SelectionStart = 0
        TextBox8.SelectionLength = TextBox8.TextLength
    End Sub
    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.SelectionStart = 0
        TextBox6.SelectionLength = TextBox6.TextLength
    End Sub
    Private Sub TextBox8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.SelectionStart = 0
        TextBox7.SelectionLength = TextBox7.TextLength
    End Sub
    Private Sub TextBox9_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.GotFocus
        TextBox9.SelectionStart = 0
        TextBox9.SelectionLength = TextBox9.TextLength
    End Sub
    Private Sub TextBox10_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox10.GotFocus
        TextBox10.SelectionStart = 0
        TextBox10.SelectionLength = TextBox10.TextLength
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click '清除DataGridView中所有数据
        DataGridView1.Rows.Clear()
    End Sub
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean '重写Form中按下回车指向“添加数据”方法
        If keyData = Keys.Enter Then
            Call PreAddData()
            Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
    Function GetPriceTaySet(ByVal SheetNo As Integer)
        '初始化价签主要字段定位信息数组
        Dim ArrayValue As Object = {}
        '初始化所选价签主要字段定位信息字典
        SheetSettingDict = New Dictionary(Of Object, Object)
        '激活所选价签版本配置信息sheet
        ExcelApp.Worksheets(SheetNo).Activate()
        '将价签配置信息写入字典中
        SheetSettingDict.Add("单排数量", ExcelApp.Cells(1, 1).Value)

        SheetSettingDict.Add("单张数量", ExcelApp.Cells(2, 1).Value)

        For j = 3 To 11

            For i = 1 To CInt(ExcelApp.Cells(2, 1).Value)

                ReDim Preserve ArrayValue(i - 1)

                ArrayValue(i - 1) = ExcelApp.Cells(j, i).Value

            Next

            SheetSettingDict.Add(PrictTagInfoArray(j - 3), ArrayValue)

        Next

        Return Nothing

    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If AddDataToDict() = True Then
            If PriceTagInfoDict.Item("品牌").Length = 0 Then
                MsgBox("并没有'等待'打印的数据可以打印！请检查！", vbYes, "提示！")
            Else

                CleanSheet()
                AddDataToSheet(Button1.Name)
            End If
        End If
    End Sub
    Function SaveAsFile(ByVal style As String)
        Dim filename As String = ""
        Dim FilePath As String = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        Dim nowtime As String
        nowtime = Format(Now(), "yyyy-mm-dd hh.mm.ss") & System.DateTime.Now.Millisecond
        If Dir(FilePath & "\价签电子版" & "(" & style & ")", vbDirectory) = "" Then
            MkDir(FilePath & "\价签电子版" & "(" & style & ")")
        End If

        If ExcelApp.Range("C1").Value.ToString <> filename Then
            filename = ExcelApp.Range("C1").Value
            If Dir(FilePath & "\价签电子版" & "(" & style & ")\" & filename, vbDirectory) <> "" Then
                ExcelBook.SaveAs(Filename:=FilePath & "\价签电子版" & "(" & style & ")\" & filename & "\" & filename & " " & nowtime & ".xls", FileFormat:=56, ReadOnlyRecommended:=False, CreateBackup:=False, Password:="")
            Else
                MkDir(FilePath & "\价签电子版" & "(" & style & ")\" & filename)
                ExcelBook.SaveAs(Filename:=FilePath & "\价签电子版" & "(" & style & ")\" & filename & "\" & filename & " " & nowtime & ".xls", FileFormat:=56, ReadOnlyRecommended:=False, CreateBackup:=False, Password:="")
            End If

        Else

            If Dir(FilePath & "\价签电子版" & "(" & style & ")\" & filename, vbDirectory) <> "" Then
                ExcelBook.SaveAs(Filename:=FilePath & "\价签电子版" & "(" & style & ")\" & filename & "\" & filename & " " & nowtime & ".xls", FileFormat:=56, ReadOnlyRecommended:=False, CreateBackup:=False, Password:="")
            Else
                MkDir(FilePath & "\价签电子版" & "(" & style & ")\" & filename)
                ExcelBook.SaveAs(Filename:=FilePath & "\价签电子版" & "(" & style & ")\" & filename & "\" & filename & " " & nowtime & ".xls", FileFormat:=56, ReadOnlyRecommended:=False, CreateBackup:=False, Password:="")
            End If
        End If
        Return Nothing
    End Function
    Function SyncPrint()

        For j = 0 To CInt(DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(9).Value) - 1
            For k = 0 To PrictTagInfoArray.length - 1
                ReDim Preserve SyncPriceTagInfoDict.Item(ArrayKey(k))(SyncPriceTagCount)
                SyncPriceTagInfoDict.Item(ArrayKey(k))(SyncPriceTagCount) = DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(k).Value
            Next
            SyncPriceTagCount = SyncPriceTagCount + 1
        Next
        If (SyncPriceTagCount - CInt((SheetSettingDict.Item("单张数量"))) * SyncTimes) \ CInt(SheetSettingDict.Item("单张数量")) > 0 Then
            For i = 1 To SyncPriceTagCount \ CInt(SheetSettingDict.Item("单张数量")) - SyncTimes
                For j = 0 To CInt(SheetSettingDict.Item("单张数量")) - 1
                    For k = 0 To 8
                        ExcelApp.Range(SheetSettingDict.Item(PrictTagInfoArray(k))(j)).Value = SyncPriceTagInfoDict.Item(PrictTagInfoArray(k))(CInt(SheetSettingDict.Item("单张数量")) * SyncTimes + j)
                    Next

                Next
                SyncTimes = SyncTimes + 1
                PrintOut()
                CleanSheet()
                SetFocus()
            Next
        End If

        Return Nothing
    End Function

    Function AddDataToDict() As Boolean

        '将DataGridView中待打印的数据添加到PriceTagInfoDict字典中
        If DataGridView1.RowCount = 0 Then
            MsgBox("您还没有添加任何需要打印的数据！", vbYesNo, "提示！")
            Return False
        Else
            Dim n As Integer = 0
            PriceTagInfoDict = New Dictionary(Of Object, Object)

            Dim ArrayValue As Object = {}
            For k = 0 To PrictTagInfoArray.length - 1
                For i = 1 To DataGridView1.RowCount
                    If DataGridView1.Rows(i - 1).Cells(10).Value = "等待" Then
                        For j = 0 To CInt(DataGridView1.Rows(i - 1).Cells(9).Value) - 1
                            ReDim Preserve ArrayValue(n)
                            ArrayValue(n) = DataGridView1.Rows(i - 1).Cells(k).Value
                            n = n + 1
                        Next
                    End If
                Next
                '按照选择价签模板进行凑整操作
                If n Mod CInt(SheetSettingDict.Item("单排数量")) <> 0 Then
                    For h = 1 To CInt(SheetSettingDict.Item("单排数量")) - (n Mod CInt(SheetSettingDict.Item("单排数量")))
                        ReDim Preserve ArrayValue(n)
                        ArrayValue(n) = DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(k).Value
                        n = n + 1
                    Next
                End If
                n = 0 '价签个数计数变量归零
                PriceTagInfoDict.Add(ArrayKey(k), ArrayValue) '将所有待打印价签信息录入到价签字典中
                ArrayValue = {} '价签字段信息临时数组清空
            Next
            Return True
        End If
    End Function
    Function CleanSheet()
        '首先清空表中数据，等待数据输入
        For i = 0 To CInt(SheetSettingDict.Item("单张数量")) - 1
            For j = 0 To PrictTagInfoArray.length - 1
                ExcelApp.Range(SheetSettingDict.Item(PrictTagInfoArray(j))(i)).Value = ""
            Next
        Next
        Return Nothing
    End Function
    Function AddDataToSheet(ByVal Method As String)
        Dim n As Integer = 0
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100

        For i = 0 To PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length - 1

            If (i + 1) Mod CInt(SheetSettingDict.Item("单张数量")) = 0 Then
                For j = 0 To PrictTagInfoArray.length - 1
                    ExcelApp.Range(SheetSettingDict.Item(PrictTagInfoArray(j))(i - n * (CInt(SheetSettingDict.Item("单张数量"))))).Value = PriceTagInfoDict.Item(PrictTagInfoArray(j))(i)
                Next
                n = n + 1
                If Method = "Button1" Then
                    PrintOut()
                    CleanSheet()
                    SetFocus()
                Else
                    SaveAsFile(ComboBox1.SelectedItem.ToString)
                    CleanSheet()
                    SetFocus()
                End If

            ElseIf i = PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length - 1 Then

                For j = 0 To PrictTagInfoArray.length - 1
                    ExcelApp.Range(SheetSettingDict.Item(PrictTagInfoArray(j))(i - n * (CInt(SheetSettingDict.Item("单张数量"))))).Value = PriceTagInfoDict.Item(PrictTagInfoArray(j))(i)
                Next
                Select Case Method
                    Case "Button1"
                        PrintOut()
                        CleanSheet()
                        SetFocus()
                    Case "Button3"
                        SaveAsFile(ComboBox1.SelectedItem.ToString)
                        CleanSheet()
                        SetFocus()
                End Select

            Else

                For j = 0 To PrictTagInfoArray.length - 1
                    ExcelApp.Range(SheetSettingDict.Item(PrictTagInfoArray(j))(i - n * (CInt(SheetSettingDict.Item("单张数量"))))).Value = PriceTagInfoDict.Item(PrictTagInfoArray(j))(i)
                Next

            End If

            ChangeStatus()
            ProgressBar1.Value = i / (PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length - 1) * 100
        Next
        If Method = "Button3" Then
            MsgBox("文件已经存放到桌面‘价签电子版" & "(" & ComboBox1.SelectedItem.ToString & ")" & "’", vbYes, "提示！")
        Else
            If PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length Mod CInt(SheetSettingDict.Item("单张数量")) = 0 Then
                MsgBox("请放入" & PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length \ CInt(SheetSettingDict.Item("单张数量")) & "张" & ComboBox1.SelectedItem.ToString, vbYes, "提示！")
            ElseIf PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length / CInt(SheetSettingDict.Item("单张数量")) = 0 Then
                MsgBox("请放入" & PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length \ CInt(SheetSettingDict.Item("单排数量")) & "排" & ComboBox1.SelectedItem.ToString, vbYes, "提示！")
            Else
                MsgBox("请放入" & PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length \ CInt(SheetSettingDict.Item("单张数量")) & "整张 + " & (PriceTagInfoDict.Item(PrictTagInfoArray(0)).Length \ CInt(SheetSettingDict.Item("单排数量"))) Mod CInt(SheetSettingDict.Item("单张数量")) \ CInt(SheetSettingDict.Item("单排数量")) & "排" & ComboBox1.SelectedItem.ToString)
            End If
        End If
        Return Nothing
    End Function
    Function ChangeStatus()
        For i = 0 To DataGridView1.RowCount - 1
            DataGridView1.Rows(i).Cells(10).Value = "√"
        Next
        Return Nothing
    End Function
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If AddDataToDict() = True Then
            CleanSheet()
            AddDataToSheet(Button3.Name)
        End If
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Button2_Click(sender, e)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Form2.CheckBoxLoaded = True And CheckBox1.Checked = True And DataGridView1.RowCount <> 0 Then
            Button1_Click(sender, e)
            SyncPrint()
        Else
            Return
        End If
    End Sub
    Function PrintOut()
        Dim aler As Object
        Try
            ExcelApp.ActiveSheet.PrintOut()
        Catch ex As Exception
            aler = MsgBox(ex.Message, vbYesNo, "提示！")
            If aler = vbYes Then
                PrintOut()
            Else
                Return Nothing
            End If
        End Try
        Return Nothing
    End Function
    Function LastPrintOut()
        If SyncPriceTagCount - SyncTimes * CInt(SheetSettingDict.Item("单张数量")) Then
            MsgBox("等待最后价签打印......", vbYes, "提示！")
            CleanSheet()
            For i = 0 To (SyncPriceTagCount - SyncTimes * CInt(SheetSettingDict.Item("单张数量")) - 1)
                For k = 0 To 8
                    ExcelApp.Range(SheetSettingDict.Item(PrictTagInfoArray(k))(i)).Value = SyncPriceTagInfoDict.Item(PrictTagInfoArray(k))(CInt(SheetSettingDict.Item("单张数量")) * SyncTimes + i)
                Next
            Next
            PrintOut()

        End If
        Return Nothing
    End Function
    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged
        If IsNumeric(TextBox10.Text) = False Then
            If Len(TextBox10.Text) <> 0 Then
                MsgBox("刚才输入的不是数字，请重新输入！", vbOKOnly, "输入错误！")
                TextBox10.Text = ""
                TextBox10.Focus()

            End If
        ElseIf InStr(TextBox10.Text, "0") = 1 Then
            If Len(TextBox10.Text) <> 0 Then
                MsgBox("数量第一位不能为0，请重新输入！", , "提示！")
                TextBox10.Text = ""
                TextBox10.Focus()

            End If
        Else

        End If
    End Sub
    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        Me.Close()
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        Button3_Click(sender, e)
    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        MsgBox("作者：徐安飞", vbYes, "作者信息")
    End Sub

    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click
        Select Case Form2.CheckBoxValue
            Case True
                Form2.CheckBoxValue = Not Form2.CheckBoxValue
                CheckBox1.Checked = Form2.CheckBoxValue
                PictureBox10.Image = ImageList2.Images(1)
            Case False
                Form2.CheckBoxValue = Not Form2.CheckBoxValue
                CheckBox1.Checked = Form2.CheckBoxValue
                PictureBox10.Image = ImageList2.Images(0)
        End Select
    End Sub

    Private Sub PictureBox10_LoadCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles PictureBox10.LoadCompleted

    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Button4_Click(sender, e)
    End Sub
End Class

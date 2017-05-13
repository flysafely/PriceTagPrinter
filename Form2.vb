Imports Microsoft.Office.Interop
Public Class Form2
    '全局变量
    Public xlAppExcelFile As Excel.Application
    Public xlBook As Excel.Workbook
    Public xlSheet As Excel.Worksheet
    Public CheckBoxLoaded As Boolean = False
    Public CheckBoxValue As Boolean
    Dim ExcelPath As String

    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = 91
        ExcelPath = Application.StartupPath & "\PriceTagModelFile.mrx"

    End Sub
    Function CheckSync()
        Dim aler As Object

        Timer1.Stop() '计时器停止

        Timer1.Enabled = False '计时器停用

        delay(0.7) '等待进度条完成后执行下列操作

        Me.Hide() '进度条隐藏

        aler = MsgBox("是否启用‘同步打印’功能", vbQuestion + vbYesNo, "提示！") '用户选择是否开启同步打印功能

        If aler = vbYes Then

            Form1.CheckBox1.Checked = True
            CheckBoxValue = Form1.CheckBox1.Checked
            CheckBoxPicload()
            Form1.Show()
            Form1.SetFocus()
            CheckBoxLoaded = True '确认checkbox已经初始化成功
        Else

            Form1.CheckBox1.Checked = False
            CheckBoxValue = Form1.CheckBox1.Checked
            CheckBoxPicload()
            Form1.Show()
            Form1.SetFocus()
            CheckBoxLoaded = True '确认checkbox已经初始化成功
        End If
        Return Nothing
    End Function
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Me.ProgressBar1.Value <= 90 Then

            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1

        Else

            CheckSync()

        End If
    End Sub
    Function OpenFileDialog()
        '默认打开路径失效，手动选取文件路径
        OpenFileDialog1.InitialDirectory = "c:\"

        OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"

        OpenFileDialog1.FilterIndex = 2

        OpenFileDialog1.RestoreDirectory = True

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            ExcelPath = OpenFileDialog1.FileName '获取选取文件路径，赋值给全局变量

        End If
        Return Nothing
    End Function
    Function GetExcel()

        Dim xlap As Excel.Application

        xlap = CreateObject("Excel.Application")

        'xlAppExcelFile = Nothing '重新调用清空，释放内存

        xlAppExcelFile = xlap '生成Excel程序实例

        xlAppExcelFile.DisplayAlerts = False '设置Excel程序中间提示窗口是否显示

        xlBook = xlAppExcelFile.Workbooks.Open(ExcelPath, Password:="q63785095")

        xlAppExcelFile.Visible = False '设置Excel程序是否显示窗口

        Timer1.Enabled = True  '开启定时器线程

        Timer1.Start()  '开始计时

        For i = 1 To xlAppExcelFile.Sheets.Count / 2  '开始执行版本页边距调整，预计时间3秒

            xlSheet = xlAppExcelFile.Worksheets(i)

            With (xlSheet.PageSetup)
                .LeftMargin = 0    '页面的左边距
                .RightMargin = 0     '页面的右左边距
                .TopMargin = 0    '页面的顶部边距
                .BottomMargin = 0    '页面的底部边距
                .HeaderMargin = 36.01    '页面顶端到页眉的距离
                .FooterMargin = 36.01    '页脚到页面底端的距离
            End With

        Next
        Return Nothing
    End Function

    Private Sub Form2_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        Dim xlap As Excel.Application

        xlap = CreateObject("Excel.Application")

        xlAppExcelFile = xlap '生成Excel程序实例

        xlAppExcelFile.DisplayAlerts = False '设置Excel程序中间提示窗口是否显示

        Try '尝试按照设定路径打开Excel模板
            xlBook = xlAppExcelFile.Workbooks.Open(ExcelPath, Password:="q63785095")
            xlAppExcelFile.Visible = False '设置Excel程序是否显示窗口
            Timer1.Enabled = True
            Timer1.Start()
            For i = 1 To xlAppExcelFile.Sheets.Count / 2 '设定价签版本的初始值以及打印页边距的需求格式

                xlSheet = xlAppExcelFile.Worksheets(i)

                With (xlSheet.PageSetup)
                    .LeftMargin = 0    '页面的左边距
                    .RightMargin = 0     '页面的右左边距
                    .TopMargin = 0    '页面的顶部边距
                    .BottomMargin = 0    '页面的底部边距
                    .HeaderMargin = 36.01    '页面顶端到页眉的距离
                    .FooterMargin = 36.01    '页脚到页面底端的距离
                End With

            Next

        Catch exc As Exception

            If MsgBox(exc.Message & "  " & "需要重新选择请选择‘是’", vbYesNo, "提示！") = vbYes Then

                OpenFileDialog() '重新选择价签模板路径

                GetExcel() '重新打开选择价签模板
            Else
                Me.xlAppExcelFile.DisplayAlerts = False
                Me.xlAppExcelFile.Quit()
                Me.xlAppExcelFile = Nothing
                Application.Exit()

            End If
        End Try
    End Sub
    Public Sub delay(ByRef Interval As Double) '延时方法

        On Error Resume Next
        Dim time As DateTime = DateTime.Now
        Dim Span As Double = Interval * 10000000   '因为时间是以100纳秒为单位。   
        While ((DateTime.Now.Ticks - time.Ticks) < Span)
            Application.DoEvents()
        End While
    End Sub
    Function CheckBoxPicload()
        Select Case Me.CheckBoxValue
            Case True
                Form1.PictureBox10.Image = Form1.ImageList2.Images(0)
            Case False
                Form1.PictureBox10.Image = Form1.ImageList2.Images(1)
        End Select
        Return Nothing
    End Function

End Class
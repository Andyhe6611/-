
Option Explicit

Sub Mac图片批量排版最终版()

    Dim doc As Document
    Dim tbl As Table
    Dim folderPath As String
    Dim fileName As String
    Dim files() As String
    Dim fileCount As Long
    Dim pairCount As Long
    Dim i As Long, j As Long
    Dim temp As String
    Dim rowPic As Long, rowDesc As Long
    Dim col As Long
    Dim cellRng As Range
    Dim pic As InlineShape
    Dim maxW As Single, maxH As Single
    Dim usableW As Single
    Dim dlgResult As String
    Dim tableRange As Range

    On Error GoTo ErrorHandler

    '=================================
    ' 1. 选择图片文件夹
    '=================================

    dlgResult = MacScript( _
        "return POSIX path of (choose folder with prompt " & _
        """请选择存放图片的文件夹"")")

    If dlgResult = "" Then Exit Sub

    folderPath = dlgResult

    If Right(folderPath, 1) <> "/" Then
        folderPath = folderPath & "/"
    End If

    '=================================
    ' 2. 读取图片文件
    '=================================

    fileCount = 0

    fileName = Dir(folderPath & "*")

    Do While fileName <> ""

        If IsMacImageFile(fileName) Then

            fileCount = fileCount + 1

            ReDim Preserve files(1 To fileCount)

            files(fileCount) = fileName

        End If

        fileName = Dir()

    Loop

    If fileCount = 0 Then

        MsgBox "没有找到支持的图片文件！", vbExclamation

        Exit Sub

    End If

    '=================================
    ' 3. 按文件名自然排序
    '=================================

    For i = 1 To fileCount - 1

        For j = i + 1 To fileCount

            If NaturalCompareMac(files(i), files(j)) > 0 Then

                temp = files(i)
                files(i) = files(j)
                files(j) = temp

            End If

        Next j

    Next i

    '=================================
    ' 4. 创建新Word文档
    '=================================

    Set doc = Documents.Add

    With doc.PageSetup

        .PaperSize = wdPaperA4

        ' 竖版
        .Orientation = wdOrientPortrait

        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1)
        .RightMargin = CentimetersToPoints(1)

    End With

    usableW = doc.PageSetup.PageWidth _
              - doc.PageSetup.LeftMargin _
              - doc.PageSetup.RightMargin

    ' 图片最大尺寸
    maxW = CentimetersToPoints(9.1)
    maxH = CentimetersToPoints(6.5)

    '=================================
    ' 5. 创建表格
    ' 每组两行：图片行+描述行
    '=================================

    pairCount = (fileCount + 1) \ 2

    Set tableRange = doc.Range(0, 0)

    Set tbl = doc.Tables.Add( _
        Range:=tableRange, _
        NumRows:=pairCount * 2, _
        NumColumns:=2)

    With tbl

        .AllowAutoFit = False

        .Columns(1).Width = usableW / 2
        .Columns(2).Width = usableW / 2

        ' 表格边距尽量贴近图片
        .TopPadding = CentimetersToPoints(0.02)
        .BottomPadding = CentimetersToPoints(0.02)
        .LeftPadding = CentimetersToPoints(0.02)
        .RightPadding = CentimetersToPoints(0.02)

        ' 显示表格边框
        .Borders.Enable = True

    End With

    '=================================
    ' 6. 插入图片
    '=================================

    For i = 1 To fileCount

        rowPic = ((i - 1) \ 2) * 2 + 1
        rowDesc = rowPic + 1

        col = ((i - 1) Mod 2) + 1

        ' 获取图片单元格
        Set cellRng = tbl.Cell(rowPic, col).Range

        cellRng.End = cellRng.End - 1

        cellRng.Collapse wdCollapseStart

        ' 插入图片
        Set pic = doc.InlineShapes.AddPicture( _
            FileName:=folderPath & files(i), _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Range:=cellRng)

        ' 保持原比例
        pic.LockAspectRatio = msoTrue

        ' 限制宽度
        If pic.Width > maxW Then
            pic.Width = maxW
        End If

        ' 限制高度
        If pic.Height > maxH Then
            pic.Height = maxH
        End If

        ' 图片居中
        With tbl.Cell(rowPic, col).Range.ParagraphFormat

            .Alignment = wdAlignParagraphCenter

            .SpaceBefore = 0
            .SpaceAfter = 0

        End With

        tbl.Cell(rowPic, col).VerticalAlignment = _
            wdCellAlignVerticalCenter

        '=================================
        ' 7. 描述栏留空
        '=================================

        With tbl.Cell(rowDesc, col).Range

            .Text = ""

            .Font.Name = "宋体"
            .Font.Size = 10

            .ParagraphFormat.Alignment = wdAlignParagraphLeft

            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0

        End With

        tbl.Cell(rowDesc, col).VerticalAlignment = _
            wdCellAlignVerticalCenter

    Next i

    '=================================
    ' 8. 设置描述栏高度
    '=================================

    For i = 1 To pairCount

        rowDesc = i * 2

        tbl.Rows(rowDesc).HeightRule = wdRowHeightAtLeast

        tbl.Rows(rowDesc).Height = _
            CentimetersToPoints(1.2)

    Next i

    '=================================
    ' 9. 每页三行图片，自动分页
    '=================================

    For i = 1 To pairCount

        rowPic = (i - 1) * 2 + 1

        If i > 1 And (i - 1) Mod 3 = 0 Then

            tbl.Cell(rowPic, 1).Range. _
                ParagraphFormat.PageBreakBefore = True

        End If

    Next i

    '=================================
    ' 10. 完成
    '=================================

    doc.Activate

    MsgBox "图片排版完成！" & vbCrLf & vbCrLf & _
           "图片总数：" & fileCount & vbCrLf & _
           "图片组数：" & pairCount & vbCrLf & _
           "页面：A4竖版" & vbCrLf & _
           "每行：2张图片" & vbCrLf & _
           "每页：最多6张图片" & vbCrLf & _
           "图片比例：保持原比例" & vbCrLf & _
           "描述栏：空白" & vbCrLf & _
           "表格边框：已添加", _
           vbInformation, "排版完成"

    Exit Sub

ErrorHandler:

    MsgBox "运行发生错误：" & vbCrLf & _
           Err.Description & vbCrLf & _
           "错误编号：" & Err.Number, _
           vbExclamation, "VBA错误"

End Sub


'=================================
' 判断图片格式
'=================================

Private Function IsMacImageFile(ByVal f As String) As Boolean

    Dim ext As String
    Dim p As Long

    If Left(f, 2) = "._" Then Exit Function

    If Left(f, 2) = "~$" Then Exit Function

    p = InStrRev(f, ".")

    If p = 0 Then Exit Function

    ext = LCase(Mid(f, p + 1))

    Select Case ext

        Case "jpg", "jpeg", "png", "bmp", _
             "gif", "tif", "tiff"

            IsMacImageFile = True

        Case Else

            IsMacImageFile = False

    End Select

End Function


'=================================
' 文件名自然排序
' 例如：1、2、3、10、11
'=================================

Private Function NaturalCompareMac( _
    ByVal a As String, _
    ByVal b As String) As Long

    Dim i As Long, j As Long

    Dim ca As String, cb As String

    Dim na As String, nb As String

    Dim va As Double, vb As Double

    i = 1
    j = 1

    Do While i <= Len(a) And j <= Len(b)

        ca = Mid(a, i, 1)

        cb = Mid(b, j, 1)

        If ca Like "#" And cb Like "#" Then

            na = ""
            nb = ""

            Do While i <= Len(a)

                ca = Mid(a, i, 1)

                If Not (ca Like "#") Then Exit Do

                na = na & ca

                i = i + 1

            Loop

            Do While j <= Len(b)

                cb = Mid(b, j, 1)

                If Not (cb Like "#") Then Exit Do

                nb = nb & cb

                j = j + 1

            Loop

            va = Val(na)

            vb = Val(nb)

            If va < vb Then

                NaturalCompareMac = -1

                Exit Function

            ElseIf va > vb Then

                NaturalCompareMac = 1

                Exit Function

            End If

        Else

            If StrComp(ca, cb, vbTextCompare) < 0 Then

                NaturalCompareMac = -1

                Exit Function

            ElseIf StrComp(ca, cb, vbTextCompare) > 0 Then

                NaturalCompareMac = 1

                Exit Function

            End If

            i = i + 1

            j = j + 1

        End If

    Loop

    If i > Len(a) And j > Len(b) Then

        NaturalCompareMac = 0

    ElseIf i > Len(a) Then

        NaturalCompareMac = -1

    Else

        NaturalCompareMac = 1

    End If

End Function
<img width="539" height="782" alt="image" src="https://github.com/user-attachments/assets/1fa04cf5-0d10-4fa7-8901-891423a5d039" />

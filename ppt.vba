Option Explicit

Sub AddProgressBar()
    Dim X As Integer
    Dim S As shape
    Dim slideW As Single
    Dim sectionCount As Integer
    Dim sectionIndex As Integer
    Dim sectionName As String

    Dim validSectionNames As Collection
    Dim i As Integer

    Dim currentValidSectionPos As Integer
    Dim currentSectionStart As Integer
    Dim currentSectionEnd As Integer
    Dim currentSectionSlideCount As Integer
    Dim currentPosInSection As Integer

    Dim sectionWidth As Single
    Dim progressWidth As Single

    With ActivePresentation
        slideW = .PageSetup.slideWidth

        Set validSectionNames = New Collection

        ' 收集所有 section 名称
        For i = 1 To .SectionProperties.Count
            validSectionNames.Add .SectionProperties.Name(i)
        Next i

        ' 移除名为“目录”的 section
        For i = validSectionNames.Count To 1 Step -1
            If StrComp(validSectionNames(i), "目录", vbTextCompare) = 0 Then
                validSectionNames.Remove i
            End If
        Next i

        ' 移除第一个 section
        If validSectionNames.Count > 0 Then
            validSectionNames.Remove 1
        End If

        ' 移除最后一个 section
        If validSectionNames.Count > 0 Then
            validSectionNames.Remove validSectionNames.Count
        End If

        If validSectionNames.Count = 0 Then Exit Sub

        sectionCount = validSectionNames.Count
        sectionWidth = slideW / sectionCount

        ' 从第 3 页到倒数第 2 页添加进度条
        For X = 3 To .Slides.Count - 1

            ' 删除旧进度条
            On Error Resume Next
            Do
                .Slides(X).Shapes("PB").Delete
            Loop Until .Slides(X).Shapes("PB") Is Nothing

            Do
                .Slides(X).Shapes("PC").Delete
            Loop Until .Slides(X).Shapes("PC") Is Nothing
            On Error GoTo 0

            ' 灰色底条
            Set S = .Slides(X).Shapes.AddLine(0, 0, slideW, 0)
            S.Line.Weight = 6
            S.Line.ForeColor.RGB = RGB(205, 205, 205)
            S.Name = "PB"

            ' 当前 slide 所属 section
            sectionIndex = .Slides(X).sectionIndex
            sectionName = .SectionProperties.Name(sectionIndex)

            ' 找到当前 section 在有效 section 中的位置
            currentValidSectionPos = 0
            For i = 1 To validSectionNames.Count
                If validSectionNames(i) = sectionName Then
                    currentValidSectionPos = i
                    Exit For
                End If
            Next i

            ' 当前 slide 不属于有效 section，只保留灰色底条
            If currentValidSectionPos = 0 Then
                GoTo NextSlide
            End If

            ' 当前 section 起始页
            currentSectionStart = .SectionProperties.FirstSlide(sectionIndex)

            ' 当前 section 结束页
            If sectionIndex < .SectionProperties.Count Then
                currentSectionEnd = .SectionProperties.FirstSlide(sectionIndex + 1) - 1
            Else
                currentSectionEnd = .Slides.Count
            End If

            ' 限制在正文范围内：第 3 页到倒数第 2 页
            If currentSectionStart < 3 Then currentSectionStart = 3
            If currentSectionEnd > .Slides.Count - 1 Then currentSectionEnd = .Slides.Count - 1

            currentSectionSlideCount = currentSectionEnd - currentSectionStart + 1
            currentPosInSection = X - currentSectionStart + 1

            If currentSectionSlideCount <= 0 Then GoTo NextSlide

            ' 按 section 分段推进
            progressWidth = _
                (currentValidSectionPos - 1) * sectionWidth + _
                currentPosInSection * sectionWidth / currentSectionSlideCount

            If progressWidth < 0 Then progressWidth = 0
            If progressWidth > slideW Then progressWidth = slideW

            ' 黄色当前进度
            Set S = .Slides(X).Shapes.AddLine(0, 0, progressWidth, 0)
            S.Line.Weight = 6
            S.Line.ForeColor.RGB = RGB(255, 255, 0)
            S.Name = "PC"

NextSlide:
        Next X
    End With
End Sub


Sub AddSectionNamesToHeader()
    Dim slide As slide
    Dim sectionIndex As Integer
    Dim sectionName As String
    Dim headerShape As shape
    Dim sectionNames As Collection
    Dim i As Integer
    Dim currentSectionName As String
    Dim sectionSlideIndex As Integer

    Set sectionNames = New Collection

    ' 收集所有 section 名称
    For i = 1 To ActivePresentation.SectionProperties.Count
        sectionNames.Add ActivePresentation.SectionProperties.Name(i)
    Next i

    ' 移除名为“目录”的 section
    For i = sectionNames.Count To 1 Step -1
        If StrComp(sectionNames(i), "目录", vbTextCompare) = 0 Then
            sectionNames.Remove i
        End If
    Next i

    ' 移除第一个 section
    If sectionNames.Count > 0 Then
        sectionNames.Remove 1
    End If

    ' 移除最后一个 section
    If sectionNames.Count > 0 Then
        sectionNames.Remove sectionNames.Count
    End If

    ' 记录每个 section 的起始页
    Dim sectionStartSlides As Object
    Set sectionStartSlides = CreateObject("Scripting.Dictionary")

    For i = 1 To ActivePresentation.SectionProperties.Count
        If Not sectionStartSlides.Exists(ActivePresentation.SectionProperties.Name(i)) Then
            sectionStartSlides.Add _
                ActivePresentation.SectionProperties.Name(i), _
                ActivePresentation.SectionProperties.FirstSlide(i)
        End If

        Debug.Print ActivePresentation.SectionProperties.Name(i) & " " & ActivePresentation.SectionProperties.FirstSlide(i)
    Next i

    ' 给每一页添加页眉 section 名称
    For Each slide In ActivePresentation.Slides
        sectionIndex = slide.sectionIndex
        currentSectionName = ActivePresentation.SectionProperties.Name(sectionIndex)

        ' 删除旧页眉
        Dim j As Integer
        For j = slide.Shapes.Count To 1 Step -1
            Set headerShape = slide.Shapes(j)
            If Left(headerShape.Name, 17) = "HeaderSectionName" Or _
               Left(headerShape.Name, 15) = "HeaderSeparator" Then
                headerShape.Delete
            End If
        Next j

        ' 添加页眉 section 名称
        Dim portion As Single
        If sectionNames.Count > 0 Then
            portion = ActivePresentation.PageSetup.slideWidth / sectionNames.Count

            For i = 1 To sectionNames.Count
                Set headerShape = slide.Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    (i - 1) * portion, _
                    6, _
                    portion, _
                    18)

                headerShape.Name = "HeaderSectionName" & i
                headerShape.TextFrame.TextRange.Font.NameFarEast = "黑体"
                headerShape.TextFrame.TextRange.Font.Name = "Times New Roman"
                headerShape.TextFrame.TextRange.Font.Size = 16
                headerShape.TextFrame.TextRange.Text = sectionNames(i)
                headerShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                headerShape.TextFrame.TextRange.Font.Color.RGB = RGB(205, 205, 205)

                If sectionNames(i) = currentSectionName Then
                    headerShape.TextFrame.TextRange.Font.Bold = msoTrue
                    headerShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 0)
                End If

                ' 添加跳转链接
                If sectionStartSlides.Exists(sectionNames(i)) Then
                    sectionSlideIndex = sectionStartSlides(sectionNames(i))

                    headerShape.ActionSettings(ppMouseClick).Hyperlink.Address = ""
                    headerShape.ActionSettings(ppMouseClick).Action = ppActionHyperlink
                    headerShape.ActionSettings(ppMouseClick).Hyperlink.SubAddress = _
                        ActivePresentation.Slides(sectionSlideIndex).SlideID & "," & _
                        sectionSlideIndex & "," & _
                        ActivePresentation.Slides(sectionSlideIndex).Name

                    headerShape.TextFrame.TextRange.Font.Underline = msoFalse
                End If

                ' 添加分隔符
                If i < sectionNames.Count Then
                    Dim sepShape As shape

                    Set sepShape = slide.Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        i * portion - 10, _
                        6, _
                        20, _
                        18)

                    sepShape.Name = "HeaderSeparator" & i
                    sepShape.TextFrame.TextRange.Font.NameFarEast = "黑体"
                    sepShape.TextFrame.TextRange.Font.Name = "Times New Roman"
                    sepShape.TextFrame.TextRange.Font.Size = 16
                    sepShape.TextFrame.TextRange.Text = "|"
                    sepShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                    sepShape.TextFrame.TextRange.Font.Color.RGB = RGB(205, 205, 205)
                End If
            Next i
        End If
    Next slide
End Sub


Sub UpdatePageFormat()
    Dim slide As slide
    Dim shape As shape

    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            Debug.Print shape.Name

            If shape.Type = msoPlaceholder Then
                If shape.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                    shape.TextFrame.TextRange.Text = _
                        slide.slideIndex & " / " & ActivePresentation.Slides.Count

                    Debug.Print "Slide Number Placeholder " & shape.Name & " " & shape.TextFrame.TextRange.Text

                    shape.Width = 60
                    shape.Left = ActivePresentation.PageSetup.slideWidth - 60
                    shape.TextFrame.TextRange.Font.NameFarEast = "黑体"
                    shape.TextFrame.TextRange.Font.Name = "Times New Roman"
                    shape.TextFrame.TextRange.Font.Size = 14
                    shape.TextFrame.TextRange.Font.Color.RGB = RGB(25, 25, 25)
                End If
            End If
        Next shape

        Debug.Print "----------------------"
    Next slide
End Sub


Sub AddNavigationLinks()
    Dim slide As slide
    Dim navShape As shape
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim linkNames As Variant
    Dim linkActions As Variant
    Dim linkIcons As Variant
    Dim i As Integer

    linkNames = Array("上一页", "下一页", "首页", "目录", "尾页")
    linkActions = Array( _
        ppActionPreviousSlide, _
        ppActionNextSlide, _
        ppActionFirstSlide, _
        ppActionHyperlink, _
        ppActionLastSlide)

    linkIcons = Array(ChrW(9194), ChrW(9193), ChrW(9198), ChrW(9208), ChrW(9197))

    For Each slide In ActivePresentation.Slides
        Debug.Print slide.Name

        slideWidth = ActivePresentation.PageSetup.slideWidth
        slideHeight = ActivePresentation.PageSetup.slideHeight

        ' 删除旧导航按钮
        Dim j As Integer
        For j = slide.Shapes.Count To 1 Step -1
            Set navShape = slide.Shapes(j)
            If Left(navShape.Name, 14) = "NavigationLink" Then
                navShape.Delete
            End If
        Next j

        For i = LBound(linkNames) To UBound(linkNames)
            Set navShape = slide.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, _
                10 + (i * 30), _
                slideHeight - 30, _
                30, _
                16)

            navShape.Name = "NavigationLink" & i
            navShape.TextFrame.TextRange.Text = linkIcons(i)
            navShape.TextFrame.TextRange.Font.Name = "Segoe UI Symbol"
            navShape.TextFrame.TextRange.Font.Size = 12
            navShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
            navShape.TextFrame.TextRange.Font.Color.RGB = RGB(205, 205, 205)

            If linkNames(i) = "目录" Then
                navShape.ActionSettings(ppMouseClick).Action = ppActionHyperlink
                navShape.ActionSettings(ppMouseClick).Hyperlink.Address = ""
                navShape.ActionSettings(ppMouseClick).Hyperlink.SubAddress = _
                    ActivePresentation.Slides(2).SlideID & ",2," & ActivePresentation.Slides(2).Name
                navShape.TextFrame.TextRange.Font.Underline = msoFalse
            Else
                navShape.ActionSettings(ppMouseClick).Action = linkActions(i)
            End If
        Next i
    Next slide
End Sub


Sub RunAllFunctions()
    AddProgressBar
    AddSectionNamesToHeader
    UpdatePageFormat
    AddNavigationLinks
End Sub

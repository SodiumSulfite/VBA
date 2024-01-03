# visio导出图片脚本，zotero与ppt文献
``` vbs
Sub AddMacroButton()
    Dim oBar As CommandBar
    Dim oControl As CommandBarButton
    Dim oButtonStyle(1 To 3) As Variant

    '这里是按钮显示名
    oButtonStyle(1) = "MacroFunc"
    '这里是按钮链接宏名
    oButtonStyle(2) = "Export_Selected_As_PNG_1200DPI_and_emz"
    '这里是按钮图案ID
    oButtonStyle(3) = 261

    Set oBar = Application.CommandBars.Add("MacroBar", , , True)
    oBar.Visible = True

    Set oControl = oBar.Controls.Add(1)
    oControl.Caption = oButtonStyle(1)
    oControl.OnAction = oButtonStyle(2)
    oControl.FaceId = oButtonStyle(3)
End Sub

Sub Export_Selected_As_PNG_1200DPI_and_emz()
    ' 代码参考自 https://zhuanlan.zhihu.com/p/632069994
    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

    ' 输入导出图片子目录
    Dim Export_SubFileName As String
    Export_SubFileName = InputBox("请输入导出图片子目录", "图片导出(PNG 1200 DPI)", "1_OPT")

    If StrPtr(Export_SubFileName) = 0 Then
        Export_SubFileName = ""
    End If

    ' 输入导出图片名称
    Dim Export_ImageName As String
    Export_ImageName = InputBox("请输入导出图片名称", "图片导出(PNG 1200 DPI)", "Fig. ")

    If StrPtr(Export_ImageName) = 0 Then
        ' User cancels
     Exit Sub
    End If

    Dim Export_ImagePath As String '获取当前文件的全名 包含路径
    Export_ImagePath = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path
    Export_ImageType = "png" '指定导出图片的格式
    Export_ImageType1 = "emz" '指定导出图片的格式

    ' 获取选中形状ID
    Dim vsoShape As Visio.Shape
    Dim vsoSelection As Visio.Selection
    Set vsoSelection = ActiveWindow.Selection ' 获取当前选择集合

    If (vsoSelection.Count > 0) Then ' 如果大于一个形状被选中

        For Each vsoShape In vsoSelection ' 遍历当前选择集合中的所有形状

            ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(vsoShape.ID), visSelect

        Next vsoShape

    Else ' 如果没有形状被选中
        MsgBox "未选中对象或选中对象数目大于1", vbCritical
     Exit Sub
    End If


    Dim Export_FolderPath As String

    If StrPtr(Export_SubFileName) = 0 Then
        Export_FolderPath = Export_ImagePath
    Else
        Export_FolderPath = Export_ImagePath + "\" + Export_SubFileName + "\"
    End If

    Dim fso, fld
    Set fso = CreateObject("scripting.filesystemobject")

    If fso.FolderExists(Export_FolderPath) = False Then
        Set fld = fso.createfolder(Export_FolderPath)
    End If

    ' 导出格式设置 其他格式自行修改(使用录制宏)
    If Export_ImageType = "png" Then
        Application.Settings.SetRasterExportResolution visRasterUsePrinterResolution, 1200#, 1200#, visRasterPixelsPerInch
        Application.Settings.SetRasterExportSize visRasterFitToSourceSize, 3.383333, 3.183333, visRasterInch
        Application.Settings.RasterExportDataFormat = visRasterInterlace
        Application.Settings.RasterExportColorFormat = visRaster24Bit
        Application.Settings.RasterExportRotation = visRasterNoRotation
        Application.Settings.RasterExportFlip = visRasterNoFlip
        Application.Settings.RasterExportBackgroundColor = 16777215
        Application.Settings.RasterExportTransparencyColor = 16777215
        Application.Settings.RasterExportUseTransparencyColor = False

        Application.ActiveWindow.Selection.Export Export_FolderPath + Export_ImageName + "." + Export_ImageType
    End If

    Application.ActiveWindow.Selection.Export Export_FolderPath + Export_ImageName + "." + Export_ImageType1

End Sub
```

# zotero 与 PPT，文献整理
- https://www.zhihu.com/question/35443903
``` vbs
Sub CollectAllRefs()
    Dim i, j, p, num_slides, count, slide_id As Long
    Dim oSld As Slide
    Dim shp As Shape
    Dim flag As Boolean
    Dim tb_name, tb_text, ret, slide_name As String
    Dim to_write() As String
    Dim tmp As String
    Dim ref_pages() As String
    Dim ref_p() As String

    'tb_name = InputBox("Please enter the name of textboxes that contain references", "Processing references", "tb_ref")
    tb_name = InputBox("请输入含有参考文献的文本框名称", "批量整理参考文献", "tb_ref")
    If StrPtr(tb_name) = 0 Then
        ' User cancels
     Exit Sub
    End If

    ReDim Preserve to_write(0 To 0)
    ReDim Preserve ref_pages(0 To 0)

    num_slides = ActivePresentation.Slides.Count
    For i = 1 To num_slides
        Set oSld = ActivePresentation.Slides(i)

        ' Skip hidden slides
        If oSld.SlideShowTransition.Hidden = msoFalse Then
            For Each oShp In oSld.Shapes
                ' Check To see If shape has a text frame And text
                If oShp.Name = tb_name And oShp.HasTextFrame And oShp.TextFrame.HasText Then
                    For p = 1 To oShp.TextFrame.TextRange.Paragraphs.Count
                        tb_text = oShp.TextFrame.TextRange.Paragraphs(p).Text
                        tb_text = Replace(tb_text, vbCrLf, "")
                        tb_text = Replace(tb_text, vbCr, "")
                        tb_text = Replace(tb_text, vbLf, "")
                        tb_text = Replace(tb_text, vbNewLine, "")
                        If Trim(tb_text & vbNullString) <> vbNullString Then
                            ' Not an empty string
                            ret = ProcessOneString(tb_text, to_write)

                            If Left(ret, 1) <> "*" Then
                                ' Add a New reference
                                ReDim Preserve to_write(0 To (UBound(to_write, 1) + 1))
                                j = UBound(to_write, 1)
                                to_write(j) = ret
                                ReDim Preserve ref_pages(0 To (UBound(ref_pages, 1) + 1))
                                ref_pages(UBound(ref_pages, 1)) = CStr(i)
                            Else
                                ' Found an existing reference
                                j = CLng(Mid(ret, 2, Len(ret)))
                                ret = to_write(j)
                                ref_pages(j) = ref_pages(j) & "," & CStr(i)
                            End If

                            ' Modify the numbering in each slide
                            count= InStr(1, oShp.TextFrame.TextRange.Paragraphs(p).Text, ret, vbTextCompare)
                            If count = 0 Then
                                ' This should Not happen For normal cases
                                Debug.Print tb_text & vbNewLine & ret & vbNewLine & "==========" & vbNewLine
                            End If
                            With oShp.TextFrame.TextRange.Paragraphs(p)
                                .Characters(1, count - 1) = ""
                                .InsertBefore("[" & j - LBound(to_write) & "] ")
                            End With
                        End If
                    Next p
                End If
            Next oShp
        End If
    Next i

    Set oSld = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.count + 1, Layout:=ppLayoutBlank)
    With oSld.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=ActivePresentation.PageSetup.SlideWidth * 0.1, _
        Top:=ActivePresentation.PageSetup.SlideHeight * 0.1, _
        Width:=ActivePresentation.PageSetup.SlideWidth * 0.8, _
        Height:=ActivePresentation.PageSetup.SlideHeight * 0.8 _
        ).TextFrame
        count = 0
        For i = LBound(to_write) To UBound(to_write)
            If Trim(to_write(i) & vbNullString) <> vbNullString Then
                count = count + 1
                tmp = "[" & count & "] " & to_write(i) & " "
                With .TextRange.InsertAfter(tmp)
                    .Font.Superscript = msoFalse
                End With

                ' Add link To pages that refer To the references
                ref_p = Split(ref_pages(i), ",")
                flag = msoFalse
                For j = LBound(ref_p) To UBound(ref_p)
                    If Trim(ref_p(j) & vbNullString) <> vbNullString Then
                        If flag Then
                            With .TextRange.InsertAfter(", ")
                                .Font.Superscript = msoTrue
                            End With
                        End If
                        With .TextRange.InsertAfter(ref_p(j))
                            slide_id = ActivePresentation.Slides(CLng(ref_p(j))).SlideID
                            slide_name = ActivePresentation.Slides(CLng(ref_p(j))).Name
                            .ActionSettings(1).Hyperlink.SubAddress = slide_id & "," & ref_p(j) & "," & slide_name
                            .Font.Superscript = msoTrue
                        End With
                        flag = msoTrue
                    End If
                Next j
                If i < UBound(to_write) Then
                    With .TextRange.InsertAfter(vbNewLine)
                        .Font.Superscript = msoFalse
                    End With
                End If
            End If
        Next i
        .AutoSize = ppAutoSizeShapeToFitText
    End With

    If count > 0 Then
        'MsgBox "Added " & count & " references at the end"
        MsgBox "已在尾页添加 " & count & " 条参考文献"
    Else
        oSld.Delete
        'MsgBox "No reference is found"
        MsgBox "未找到参考文献"
    End If
End Sub


Function ProcessOneString(in_text As Variant, all_text As Variant) As String
    Dim j As Long
    Dim found_match As Boolean
    Dim record As String

    Dim strPattern As String
    #If Mac Then
    in_text = Replace(in_text, """", "\""")
    sMacScript = "Set s To """ & in_text & """" & vbNewLine & _
    "Set srpt To ""echo \"""" & s & ""\"" | sed -r \""s/^([0-9]+([[:space:]]|\\.)|[[【][0-9]+[]】]|\\*)[[:space:]]*//\""""" & vbNewLine & _
    "return (Do shell script srpt)"

    Debug.Print sMacScript
    in_text = MacScript(sMacScript)
    #Else
    strPattern = "^([0-9]+[\s\.]|[\[【][0-9]+[\]】]|\*)\s*"
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = strPattern
    End With

    If regEx.TEST(in_text) Then
        in_text = regEx.Replace(in_text, "")
    End If
    #End If

    found_match = msoFalse
    For j = LBound(all_text) To UBound(all_text)
        record = CStr(all_text(j))
        If Trim(record & vbNullString) <> vbNullString Then
            If InStr(1, record, in_text, vbTextCompare) > 0 Then
                found_match = msoTrue
                ProcessOneString = "*" & CStr(j)
             Exit For
            End If
        End If
    Next j
    If found_match = msoFalse Then
        ProcessOneString = in_text
    End If

End Function
```

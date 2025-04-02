' 定义常量
Const API_URL As String = "https://api.siliconflow.cn/v1/chat/completions" ' SiliconFlow API地址
Const API_KEY As String = "sk-xTYifeB7vMG2mUp4111816A4C6124f8cA0Dc741528E04cDd" ' API密钥
Const SYSTEM_PROMPT As String = "您是一位拥有20年丰富经验的资深写作助手，曾担任秘书、记者、编辑、公务员文案师和市场营销师等职务。请根据我提供的要求为我提供帮助。请以Markdown格式呈现内容。"
Const MODEL_NAME As String = "deepseek-ai/DeepSeek-V3" ' 模型名称


' 可靠的替代方案，申请免费API地址：https://cloud.siliconflow.cn/i/CCl5Mnrb 新用户注册即得 2000万 Tokens。


' 清理输入文本中的特殊字符
Function CleanInputText(ByVal text As String) As String
    Dim cleanedText As String
    cleanedText = text
    
    ' 替换特殊控制字符
    cleanedText = Replace(cleanedText, vbCrLf, " ") ' 换行符替换为空格
    cleanedText = Replace(cleanedText, vbCr, " ")   ' 回车符替换为空格
    cleanedText = Replace(cleanedText, vbLf, " ")   ' 换行符替换为空格
    cleanedText = Replace(cleanedText, vbTab, " ")  ' 制表符替换为空格
    cleanedText = Replace(cleanedText, Chr(11), " ") ' 垂直制表符(\v)替换为空格
    
    ' 转义JSON所需的特殊字符
    cleanedText = Replace(cleanedText, "\", "\\")   ' 反斜杠转义
    cleanedText = Replace(cleanedText, """", "\""") ' 双引号转义
    
    CleanInputText = cleanedText
End Function

' API调用函数
Function CallDeepSeekAPI(apiKey As String, inputText As String) As String
    Dim requestBody As String
    Dim httpRequest As Object
    Dim statusCode As Integer
    Dim response As String
    
    ' 输入验证
    If Len(Trim(inputText)) = 0 Then
        CallDeepSeekAPI = "Error: Input text is empty."
        Exit Function
    End If
    
    ' 构造请求体
    requestBody = "{""model"": """ & MODEL_NAME & """, ""messages"": [{""role"":""system"", ""content"":""" & SYSTEM_PROMPT & """}, {""role"":""user"", ""content"":""" & inputText & """}], ""stream"": false}"
    
    On Error Resume Next
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    If Err.Number <> 0 Then
        CallDeepSeekAPI = "Error: Unable to create HTTP object - " & Err.Description
        Exit Function
    End If
    On Error GoTo 0
    
    With httpRequest
        .Open "POST", API_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .setTimeouts 30000, 30000, 300000, 300000 ' 设置超时：解析、连接、发送、接收（毫秒）
        On Error Resume Next
        .send requestBody
        If Err.Number <> 0 Then
            CallDeepSeekAPI = "Error: Network issue - " & Err.Description
            Exit Function
        End If
        On Error GoTo 0
        statusCode = .Status
        response = .responseText
    End With
    
    ' 根据状态码返回详细错误
    Select Case statusCode
        Case 200
            CallDeepSeekAPI = response
        Case 401
            CallDeepSeekAPI = "Error: Invalid API key."
        Case 429
            CallDeepSeekAPI = "Error: API rate limit exceeded."
        Case Else
            CallDeepSeekAPI = "Error: " & statusCode & " - " & response
    End Select
    
    Set httpRequest = Nothing
End Function

' Unicode转义解码函数
Function DecodeUnicode(ByVal text As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim i As Long
    Dim decodedText As String
    Dim unicodeChar As String
    Dim charCode As Long
    
    decodedText = text
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .Pattern = "\\u([0-9A-Fa-f]{4})"
    End With
    
    Set matches = regex.Execute(text)
    For i = 0 To matches.Count - 1
        unicodeChar = matches(i).SubMatches(0)
        charCode = CLng("&H" & unicodeChar)
        decodedText = Replace(decodedText, "\u" & unicodeChar, ChrW(charCode))
    Next i
    
    DecodeUnicode = decodedText
End Function

' 设置行距
Sub SetLineSpacing(selection As Object)
    With selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 1.5 * selection.Font.Size ' 1.5倍行距
    End With
End Sub

' 设置字体样式
Sub ApplyStyle(selection As Object, styleName As String, fontName As String, fontSize As Single, Optional isBold As Boolean = False)
    With selection
        .Style = ActiveDocument.Styles(styleName)
        .Font.Name = fontName
        .Font.Size = fontSize
        .Font.Bold = isBold
        SetLineSpacing selection
    End With
End Sub

' Markdown解析并插入Word
Sub ParseMarkdownToWord(content As String, selection As Object)
    Dim lines As Variant
    Dim i As Integer
    Dim line As String
    Dim isList As Boolean
    
    content = Replace(content, "\n", vbCrLf) ' 处理换行符
    lines = Split(content, vbCrLf)
    isList = False
    
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i)) ' 去除前后空格
        If line = "" Then
            selection.TypeParagraph ' 空行
        ElseIf Left(line, 2) = "# " Then
            ' 主标题：小二号黑体
            ApplyStyle selection, wdStyleTitle, "黑体", 18
            selection.TypeText text:=Replace(Mid(line, 3), ".", "")
            selection.TypeParagraph
            isList = False
        ElseIf Left(line, 3) = "## " Then
            ' 副标题：四号黑体
            ApplyStyle selection, wdStyleHeading1, "黑体", 14
            selection.TypeText text:=Replace(Mid(line, 4), ".", "")
            selection.TypeParagraph
            isList = False
        ElseIf Left(line, 4) = "### " Or Left(line, 5) = "#### " Then
            ' 三四级标题：小四号仿宋
            ApplyStyle selection, wdStyleHeading2, "仿宋", 12
            If Left(line, 4) = "### " Then
                selection.TypeText text:=Replace(Mid(line, 5), ".", "")
            Else
                selection.TypeText text:=Replace(Mid(line, 6), ".", "")
            End If
            selection.TypeParagraph
            isList = False
        ElseIf Left(line, 2) = "- " Then
            ' 无序列表：小四号仿宋
            If Not isList Then
                ApplyStyle selection, wdStyleListParagraph, "仿宋", 12
                isList = True
            End If
            selection.TypeText text:=Mid(line, 3)
            selection.TypeParagraph
        ElseIf Left(line, 3) = "1. " Then
            ' 有序列表：小四号仿宋
            If Not isList Then
                ApplyStyle selection, wdStyleListParagraph, "仿宋", 12
                isList = True
            End If
            selection.TypeText text:=Mid(line, 4)
            selection.TypeParagraph
        Else
            ' 普通段落：小四号仿宋，首行缩进
            If isList Then
                selection.Style = ActiveDocument.Styles(wdStyleNormal)
                isList = False
            End If
            With selection
                ApplyStyle selection, wdStyleNormal, "仿宋", 12
                .ParagraphFormat.FirstLineIndent = CentimetersToPoints(0.74) ' 首行缩进两个字符
                .TypeText text:=line
                .TypeParagraph
            End With
        End If
    Next i
End Sub

' 主函数
Sub DeepSeekV3()
    Dim inputText As String
    Dim response As String
    Dim regex As Object
    Dim matches As Object
    Dim originalSelection As Object
    Dim content As String
    
    ' 输入验证
    If API_KEY = "" Then
        MsgBox "请填写API密钥。", vbExclamation
        Exit Sub
    ElseIf selection.Type <> wdSelectionNormal Then
        MsgBox "请选择需要处理的文本。", vbExclamation
        Exit Sub
    End If
    
    ' 保存原始选区
    Set originalSelection = selection.Range.Duplicate
    ' 清理输入文本
    inputText = CleanInputText(selection.text)
    
    ' 在状态栏显示提示
    Application.ScreenUpdating = False
    Application.StatusBar = "正在调用API，请稍候..."
    DoEvents ' 确保状态栏更新
    
    ' 调用API
    response = CallDeepSeekAPI(API_KEY, inputText)
    
    If Left(response, 5) <> "Error" Then
        ' 解析API返回的JSON
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = """content"":""(.*?)"""
        End With
        
        Set matches = regex.Execute(response)
        If matches.Count > 0 Then
            content = matches(0).SubMatches(0)
            content = Replace(Replace(content, "\""", Chr(34)), "\""", Chr(34))
            content = DecodeUnicode(content)
            
            ' 取消选中原始文本并插入新行
            selection.Collapse Direction:=wdCollapseEnd
            selection.TypeParagraph
            
            ' 解析并插入Markdown内容
            ParseMarkdownToWord content, selection
            
            ' 光标移回原始选区末尾
            originalSelection.Select
        Else
            Application.StatusBar = "" ' 清除状态栏
            MsgBox "无法解析API返回的内容。", vbExclamation
        End If
    Else
        Application.StatusBar = "" ' 清除状态栏
        MsgBox response, vbCritical
    End If
    
    ' 清除状态栏并恢复屏幕更新
    Application.StatusBar = ""
    Application.ScreenUpdating = True
End Sub

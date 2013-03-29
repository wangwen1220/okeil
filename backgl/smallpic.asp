<%  
'
'功能：按照指定图片生成缩略图  
'注意：以下提到的“路径”都是值相对于调用本函数的文件的相对路径  
'参数：  
'’    s_OriginalPath:     原图片路径 例:images/image1.gif  
'’    s_BuildBasePath:    生成图片的基路径,不论是否以“/”结尾均可 例:images或images/  
'’    n_MaxWidth:         生成图片最大宽度  
'’                        如果在前台显示的缩略图是 100*100,这里 n_MaxWidth=100,n_MaxHeight=100.  
'’    n_MaxHeight:        生成图片最大高度  
'’返回值：  
'’    返回生成后的缩略图的路径  
'’错误处理：  
'’    如果函数执行过程中出现错误,将返回错误代码,错误代码以 “Error”开头  
'’        Error_01:创建AspJpeg组件失败,没有正确安装注册该组件  
'’        Error_02:原图片不存在,检查s_OriginalPath参数传入值  
'’        Error_03:缩略图存盘失败.可能原因:缩略图保存基地址不存在,检查s_OriginalPath参数传入值;对目录没有写权限;磁盘空间不足  
'’        Error_Other:未知错误  
'’调用例子:  
'’    Dim sSmallPath ’缩略图路径  
'’    sSmallPath = BuildSmallPic("images/image1.gif", "images", 100, 100)      
'’
Function BuildSmallPic(s_OriginalPath, s_BuildBasePath, n_MaxWidth, n_MaxHeight)  
    Err.Clear  
    On Error Resume Next  
      
    '检查组件是否已经注册  
    Dim AspJpeg  
    Set AspJpeg = Server.Createobject("Persits.Jpeg")  
    If Err.Number <> 0 Then  
        Err.Clear  
        BuildSmallPic = "Error_01"  
        Exit Function  
    End If  
    '检查原图片是否存在  
    Dim s_MapOriginalPath  
    s_MapOriginalPath = Server.MapPath(s_OriginalPath)  
    AspJpeg.Open s_MapOriginalPath '打开原图片  
    If Err.Number <> 0 Then  
        Err.Clear  
        BuildSmallPic = "Error_02"  
        Exit Function  
    End If  
    '按比例取得缩略图宽度和高度  
    Dim n_OriginalWidth, n_OriginalHeight '原图片宽度、高度  
    Dim n_BuildWidth, n_BuildHeight '缩略图宽度、高度  
    Dim div1, div2  
    Dim n1, n2  
    n_OriginalWidth = AspJpeg.Width  
    n_OriginalHeight = AspJpeg.Height  
    div1 = n_OriginalWidth / n_OriginalHeight  
    div2 = n_OriginalHeight / n_OriginalWidth  
    n1 = 0  
    n2 = 0  
    If n_OriginalWidth > n_MaxWidth Then  
        n1 = n_OriginalWidth / n_MaxWidth  
    Else  
        n_BuildWidth = n_OriginalWidth  
    End If  
    If n_OriginalHeight > n_MaxHeight Then  
        n2 = n_OriginalHeight / n_MaxHeight  
    Else  
        n_BuildHeight = n_OriginalHeight  
    End If  
    If n1 <> 0 Or n2 <> 0 Then  
        If n1 > n2 Then  
            n_BuildWidth = n_MaxWidth  
            n_BuildHeight = n_MaxWidth * div2  
        Else  
            n_BuildWidth = n_MaxHeight * div1  
            n_BuildHeight = n_MaxHeight  
        End If  
    End If  
    '指定宽度和高度生成  
    AspJpeg.Width = n_BuildWidth  
    AspJpeg.Height = n_BuildHeight  
      
    '--将缩略图存盘开始--  
    Dim pos, s_OriginalFileName, s_OriginalFileExt '位置、原文件名、原文件扩展名  
    pos = InStrRev(s_OriginalPath, "/") + 1  
    s_OriginalFileName = Mid(s_OriginalPath, pos)  
    pos = InStrRev(s_OriginalFileName, ".")  
    s_OriginalFileExt = Mid(s_OriginalFileName, pos)  
    Dim s_MapBuildBasePath, s_MapBuildPath, s_BuildFileName '缩略图绝对路径、缩略图文件名  
    Dim s_EndFlag '小图片文件名结尾标识 例: 如果大图片文件名是“image1.gif”,结尾标识是“_small”,那么小图片文件名就是“image1_small.gif”  
    If Right(s_BuildBasePath, 1) <> "/" Then s_BuildBasePath = s_BuildBasePath & "/"  
    s_MapBuildBasePath = Server.MapPath(s_BuildBasePath)  
    s_EndFlag = "" '可以自定义,只要能区别大小图片即可  
    s_BuildFileName = Replace(s_OriginalFileName, s_OriginalFileExt, "") & s_EndFlag & s_OriginalFileExt  
    s_MapBuildPath = s_MapBuildBasePath & "\" & s_BuildFileName  
      
    AspJpeg.Save s_MapBuildPath '保存  
    If Err.Number <> 0 Then  
        Err.Clear  
        BuildSmallPic = "Error_03"  
        Exit Function  
    End If  
    '--将缩略图存盘结束--  
    '注销实例  
    Set AspJpeg = Nothing  
    If Err.Number <> 0 Then  
        BuildSmallPic = "Error_Other"  
        Err.Clear  
    End If  
    BuildSmallPic = s_BuildBasePath & s_BuildFileName  
End Function  
%>
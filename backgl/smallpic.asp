<%  
'
'���ܣ�����ָ��ͼƬ��������ͼ  
'ע�⣺�����ᵽ�ġ�·��������ֵ����ڵ��ñ��������ļ������·��  
'������  
'��    s_OriginalPath:     ԭͼƬ·�� ��:images/image1.gif  
'��    s_BuildBasePath:    ����ͼƬ�Ļ�·��,�����Ƿ��ԡ�/����β���� ��:images��images/  
'��    n_MaxWidth:         ����ͼƬ�����  
'��                        �����ǰ̨��ʾ������ͼ�� 100*100,���� n_MaxWidth=100,n_MaxHeight=100.  
'��    n_MaxHeight:        ����ͼƬ���߶�  
'������ֵ��  
'��    �������ɺ������ͼ��·��  
'��������  
'��    �������ִ�й����г��ִ���,�����ش������,��������� ��Error����ͷ  
'��        Error_01:����AspJpeg���ʧ��,û����ȷ��װע������  
'��        Error_02:ԭͼƬ������,���s_OriginalPath��������ֵ  
'��        Error_03:����ͼ����ʧ��.����ԭ��:����ͼ�������ַ������,���s_OriginalPath��������ֵ;��Ŀ¼û��дȨ��;���̿ռ䲻��  
'��        Error_Other:δ֪����  
'����������:  
'��    Dim sSmallPath ������ͼ·��  
'��    sSmallPath = BuildSmallPic("images/image1.gif", "images", 100, 100)      
'��
Function BuildSmallPic(s_OriginalPath, s_BuildBasePath, n_MaxWidth, n_MaxHeight)  
    Err.Clear  
    On Error Resume Next  
      
    '�������Ƿ��Ѿ�ע��  
    Dim AspJpeg  
    Set AspJpeg = Server.Createobject("Persits.Jpeg")  
    If Err.Number <> 0 Then  
        Err.Clear  
        BuildSmallPic = "Error_01"  
        Exit Function  
    End If  
    '���ԭͼƬ�Ƿ����  
    Dim s_MapOriginalPath  
    s_MapOriginalPath = Server.MapPath(s_OriginalPath)  
    AspJpeg.Open s_MapOriginalPath '��ԭͼƬ  
    If Err.Number <> 0 Then  
        Err.Clear  
        BuildSmallPic = "Error_02"  
        Exit Function  
    End If  
    '������ȡ������ͼ��Ⱥ͸߶�  
    Dim n_OriginalWidth, n_OriginalHeight 'ԭͼƬ��ȡ��߶�  
    Dim n_BuildWidth, n_BuildHeight '����ͼ��ȡ��߶�  
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
    'ָ����Ⱥ͸߶�����  
    AspJpeg.Width = n_BuildWidth  
    AspJpeg.Height = n_BuildHeight  
      
    '--������ͼ���̿�ʼ--  
    Dim pos, s_OriginalFileName, s_OriginalFileExt 'λ�á�ԭ�ļ�����ԭ�ļ���չ��  
    pos = InStrRev(s_OriginalPath, "/") + 1  
    s_OriginalFileName = Mid(s_OriginalPath, pos)  
    pos = InStrRev(s_OriginalFileName, ".")  
    s_OriginalFileExt = Mid(s_OriginalFileName, pos)  
    Dim s_MapBuildBasePath, s_MapBuildPath, s_BuildFileName '����ͼ����·��������ͼ�ļ���  
    Dim s_EndFlag 'СͼƬ�ļ�����β��ʶ ��: �����ͼƬ�ļ����ǡ�image1.gif��,��β��ʶ�ǡ�_small��,��ôСͼƬ�ļ������ǡ�image1_small.gif��  
    If Right(s_BuildBasePath, 1) <> "/" Then s_BuildBasePath = s_BuildBasePath & "/"  
    s_MapBuildBasePath = Server.MapPath(s_BuildBasePath)  
    s_EndFlag = "" '�����Զ���,ֻҪ�������СͼƬ����  
    s_BuildFileName = Replace(s_OriginalFileName, s_OriginalFileExt, "") & s_EndFlag & s_OriginalFileExt  
    s_MapBuildPath = s_MapBuildBasePath & "\" & s_BuildFileName  
      
    AspJpeg.Save s_MapBuildPath '����  
    If Err.Number <> 0 Then  
        Err.Clear  
        BuildSmallPic = "Error_03"  
        Exit Function  
    End If  
    '--������ͼ���̽���--  
    'ע��ʵ��  
    Set AspJpeg = Nothing  
    If Err.Number <> 0 Then  
        BuildSmallPic = "Error_Other"  
        Err.Clear  
    End If  
    BuildSmallPic = s_BuildBasePath & s_BuildFileName  
End Function  
%>
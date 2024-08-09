import win32com.client

def copy_slide(source_ppt, target_ppt, source_slide_index, target_slide_index):
    # 打开 PowerPoint 应用程序
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True

    # 打开源文件和目标文件
    presentation_a = powerpoint.Presentations.Open(source_ppt)
    presentation_b = powerpoint.Presentations.Open(target_ppt)

    # 获取源文件的第 source_slide_index 张幻灯片（索引从 1 开始）
    slide_to_copy = presentation_a.Slides(source_slide_index)

    # 复制幻灯片到目标文件
    slide_to_copy.Copy()

    # 粘贴到目标文件的第 target_slide_index 张幻灯片之前
    presentation_b.Slides.Paste(Index=target_slide_index)

    # 保存并关闭文件
    presentation_b.Save()
    presentation_a.Close()
    presentation_b.Close()

    # 退出 PowerPoint 应用程序
    powerpoint.Quit()

# 文件路径
source_ppt = r'D:\pyproj\yz_win_server\test\标杆案例-谜底.pptx'
target_ppt = r'D:\pyproj\yz_win_server\test\首面.pptx'

# 调用函数，将 a.pptx 的第 4 张幻灯片复制到 b.pptx 的第 6 张幻灯片位置
copy_slide(source_ppt, target_ppt, 4, 6)

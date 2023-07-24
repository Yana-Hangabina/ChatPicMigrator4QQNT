import os
import sys
import shutil
import time
import psutil
import PySimpleGUI as sg
import csv

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_all_files(directory):
    files_list = []
    for root, _, files in os.walk(directory):
        for filename in files:
            filepath = os.path.join(root, filename)
            files_list.append(filepath)
    return files_list

def get_file_timestamp(file_path):
    create_time = os.path.getctime(file_path)
    modify_time = os.path.getmtime(file_path)
    access_time = os.path.getatime(file_path)

    # 选择最早的那个时间作为文件的时间戳
    timestamp = min(create_time, modify_time, access_time)
    return timestamp

def get_qq_dir():
    # 显示源文件夹提示框
    sg.popup("请选择您的QQ存储目录，其通常位于您的文档中。\n路径形如：C:\\Users\\[您的用户名]\\Documents\\Tencent Files\\[您的QQ号]", title=app_name)

    qq_dir = sg.popup_get_folder("选择QQ存储目录")
    return qq_dir

def copy_files(isVid):
    qq_dir = get_qq_dir()
    src_dir = os.path.join(qq_dir, 'Video') if isVid else os.path.join(qq_dir, 'Image')
    dst_dir = os.path.join(qq_dir, 'nt_qq/nt_data/Video') if isVid else os.path.join(qq_dir, 'nt_qq/nt_data/Pic')
    if not src_dir:
        sg.popup("未选择源文件夹！", title=app_name)
        return
    if not dst_dir:
        sg.popup("QQ NT文件夹不存在！", title=app_name)
        return

    # 获取所有文件列表
    files_list = get_all_files(src_dir)

    # 判断目标盘号的剩余空间是否大于源文件夹大小
    source_size = 0
    for file_path in files_list:
        source_size += os.path.getsize(file_path)
    dest_drive = os.path.splitdrive(dst_dir)[0]
    dest_free_space = psutil.disk_usage(dest_drive).free
    if dest_free_space < source_size:
        sg.popup("磁盘空间不足！", title=app_name)
        return

    total_files = len(files_list)
    copied_files = 0
    skipped_files = 0
    copied_files_list = []

    layout = [[sg.Text(f"{'视频' if isVid else '图片'}文件迁移中...", justification='center', key='prompt')],
            [sg.ProgressBar(total_files, orientation='h', size=(100, 15), key='progressbar')],
            [sg.Cancel()]]

    window = sg.Window('迁移旧版QQ文件', layout, no_titlebar=True, grab_anywhere=True)

    # 遍历所有文件
    for file_path in files_list:
        event, values = window.read(timeout=10)
        if event == 'Cancel' or event == sg.WIN_CLOSED:
            if sg.popup_yes_no("是否取消复制操作？") == "Yes":
                sg.popup("复制已取消！", title=app_name)
                break
        
        # 获取文件名和路径
        file_name = os.path.basename(file_path)
        file_dir = os.path.dirname(file_path)
        
        window['progressbar'].update_bar(copied_files + skipped_files)

        if os.path.basename(file_dir) == 'Thumbnails':
            skipped_files += 1
            window['prompt'].update(f"跳过Thubnail，进度：{copied_files+skipped_files}/{total_files}")
            continue

        # 排除没有后缀名的文件
        if not os.path.splitext(file_name)[1]:
            skipped_files += 1
            window['prompt'].update(f"跳过无后缀文件，进度：{copied_files+skipped_files}/{total_files}")
            continue

        if isVid and os.path.splitext(file_name)[1] != '.mp4':
            skipped_files += 1
            window['prompt'].update(f"跳过非视频文件，进度：{copied_files+skipped_files}/{total_files}")
            continue

        # 获取文件时间戳
        timestamp = get_file_timestamp(file_path)
        # 排除时间戳为0的文件
        if timestamp == 0:
            skipped_files += 1
            continue

        # 将时间戳转换为struct_time格式
        time_struct = time.localtime(timestamp)
        # 获取年份和月份
        year_month = time.strftime("%Y-%m", time_struct)
        # 构建目标文件夹路径
        dest_dir = os.path.join(dst_dir, year_month, "Ori")
        # 如果目标文件夹不存在，则创建它
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir)

        # 构建目标文件路径
        dest_path = os.path.join(dest_dir, file_name)

        # 复制文件到目标文件夹中
        shutil.copy(file_path, dest_path)
        # 将复制成功的文件路径添加到列表中
        copied_files_list.append(dest_path)
        copied_files += 1
        window['prompt'].update(f"{dest_path} 复制成功，进度：{copied_files+skipped_files}/{total_files}")

        # 将源文件路径和目标文件路径添加到列表中
        copied_files_list.append((file_path, dest_path))

    window.close()

    # 保存复制文件列表为CSV文件
    fields = ["源文件路径", "目标文件路径"]
    with open(os.path.join(qq_dir, f"migrated_{'vid' if isVid else 'pic'}s_{time.strftime('%Y%m%d-%H%M%S')}.csv"), "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(fields)
        writer.writerows(copied_files_list)

    # 显示复制结果
    sg.popup(f"共复制 {copied_files} 个文件。", title=app_name)

def copy_filerecvs():
    qq_dir = get_qq_dir()
    src_dir = os.path.join(qq_dir, 'FileRecv')
    if not src_dir:
        sg.popup("未选择源文件夹！", title=app_name)
        return
    sg.popup("请选择新版QQ NT的文件存储目录，位置可以查看新版QQ设置，一般位于：C:\\Users\\[您的用户名]\\Documents\\Tencent Files\\nt_qq_files", title=app_name)
    dst_dir = sg.popup_get_folder("选择新版QQ的文件存储目录")
    if not dst_dir:
        sg.popup("新版QQ的文件存储目录不存在！", title=app_name)
        return

    # 获取所有文件列表
    files_list = get_all_files(src_dir)

    # 判断目标盘号的剩余空间是否大于源文件夹大小
    source_size = 0
    for file_path in files_list:
        source_size += os.path.getsize(file_path)
    dest_drive = os.path.splitdrive(dst_dir)[0]
    dest_free_space = psutil.disk_usage(dest_drive).free
    if dest_free_space < source_size:
        sg.popup("磁盘空间不足！", title=app_name)
        return

    total_files = len(files_list)
    copied_files = 0
    skipped_files = 0
    copied_files_list = []

    layout = [[sg.Text(f"接收文件迁移中...", justification='center', key='prompt')],
            [sg.ProgressBar(total_files, orientation='h', size=(100, 15), key='progressbar')],
            [sg.Cancel()]]

    window = sg.Window('迁移旧版QQ接收到的文件', layout, no_titlebar=True, grab_anywhere=True)

    # 遍历所有文件
    for file_path in files_list:
        event, values = window.read(timeout=10)
        if event == 'Cancel' or event == sg.WIN_CLOSED:
            if sg.popup_yes_no("是否取消复制操作？") == "Yes":
                sg.popup("复制已取消！", title=app_name)
                break
        
        # 获取文件名和路径
        file_name = os.path.basename(file_path)
        file_dir = os.path.dirname(file_path)
        
        window['progressbar'].update_bar(copied_files + skipped_files)

        # 构建目标文件路径
        dst_path = os.path.join(dst_dir, file_name)

        # 复制文件到目标文件夹中
        shutil.copy(file_path, dst_path)
        # 将复制成功的文件路径添加到列表中
        copied_files_list.append(dst_path)
        copied_files += 1
        window['prompt'].update(f"{dst_path} 复制成功，进度：{copied_files+skipped_files}/{total_files}")

        # 将源文件路径和目标文件路径添加到列表中
        copied_files_list.append((file_path, dst_path))

    window.close()

    # 保存复制文件列表为CSV文件
    fields = ["源文件路径", "目标文件路径"]
    with open(os.path.join(qq_dir, f"migrated_recvs_{time.strftime('%Y%m%d-%H%M%S')}.csv"), "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(fields)
        writer.writerows(copied_files_list)

    # 显示复制结果
    sg.popup(f"共复制 {copied_files} 个文件。", title=app_name)
    
def delete_old_qq_files():
    src_dir = get_qq_dir()
    if not src_dir:
        sg.popup("未选择源文件夹！", title=app_name)
        return

    total_files = len(os.listdir(src_dir))

    layout = [[sg.Text('文件删除中...', justification='center', key='prompt')],
            [sg.ProgressBar(total_files, orientation='h', size=(100, 15), key='progressbar')],
            [sg.Cancel()]]

    window = sg.Window('删除旧版QQ文件', layout, no_titlebar=True, grab_anywhere=True)

    deleted_files = 0
    skipped_files = 0
    for filename in os.listdir(src_dir):
        event, values = window.read(timeout=10)
        if event == 'Cancel' or event == sg.WIN_CLOSED:
            return
        
        filepath = os.path.join(src_dir, filename)
        if 'nt_qq' not in filepath:
            if os.path.isfile(filepath):
                try:
                    os.remove(filepath)
                    deleted_files += 1
                    window['prompt'].update(f"删除文件 {filepath}")
                except Exception as e:
                    skipped_files += 1
                    window['prompt'].update(f"无法删除文件 {filepath}")
            elif os.path.isdir(filepath):
                try:
                    shutil.rmtree(filepath)
                    deleted_files += 1
                    window['prompt'].update(f"删除文件夹 {filepath}")
                except Exception as e:
                    skipped_files += 1
                    window['prompt'].update(f"无法删除文件夹 {filepath}")
        else:
            skipped_files += 1
            window['prompt'].update(f"跳过QQ NT文件 {filepath}")
        window['progressbar'].update_bar(deleted_files + skipped_files)

    window.close()
    sg.popup(f"删除旧版QQ文件成功！已删除 {deleted_files} 个文件/文件夹。", title=app_name)

def revert_copy():
    csv_file = sg.popup_get_file("选择保存的CSV文件记录")
    if not csv_file:
        sg.popup("未找到已复制的CSV文件！", title=app_name)
        return
    
    with open(csv_file, "r") as f:
        reader = csv.reader(f)
        next(reader) # 跳过标题行
        dest_paths = [row[1] for row in reader]
    
    total_files = len(dest_paths)
    deleted_files = 0
    
    layout = [[sg.Text('撤回迁移中...', justification='center', key='prompt')],
            [sg.ProgressBar(total_files, orientation='h', size=(100, 15), key='progressbar')],
            [sg.Cancel()]]

    window = sg.Window('撤回迁移文件', layout, no_titlebar=True, grab_anywhere=True)

    for path in dest_paths:
        event, values = window.read(timeout=10)
        if event == 'Cancel' or event == sg.WIN_CLOSED:
            return
        
        if os.path.exists(path):
            os.remove(path)
            deleted_files += 1
            window['prompt'].update(f"撤回迁移 {path}")
        # 更新进度条
        window['progressbar'].update_bar(deleted_files)
    
    window.close()
    sg.popup(f"共撤回 {deleted_files} 个复制的文件。", title=app_name)
    
if __name__ == "__main__":
    
    app_name = "ChatPicMigrator4QQNT"
    sg.theme("DefaultNoMoreNagging")
    sg.set_options(font=("微软雅黑", 12), icon=get_resource_path("icon.ico"))

    layout = [
        [sg.Text("请确保您同时安装了旧版QQ PC与新版QQ NT，并完成了新版聊天记录导入。\n请选择需要执行的操作：")],
        [sg.Button("复制图片"), sg.Button("复制视频"), sg.Button("复制接收的文件"), sg.Button("删除旧版QQ文件"), sg.Button("撤回复制")],
    ]

    window = sg.Window(app_name, layout, element_justification='c')

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "复制图片":
            copy_files(False)        
        elif event == "复制视频":
            copy_files(True)
        elif event == "复制接收的文件":
            copy_filerecvs()
        elif event == "删除旧版QQ文件":
            delete_old_qq_files()
        elif event == "撤回复制":
            revert_copy()

    window.close()

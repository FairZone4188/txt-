import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from pathlib import Path
from datetime import datetime


def create_progress_window(title, total_steps, parent=None):
    """创建进度窗口"""
    progress_root = tk.Toplevel(parent or tk.Tk())
    progress_root.title(title)
    progress_root.geometry("400x150")
    progress_root.resizable(False, False)

    # 添加进度条
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(
        progress_root,
        orient="horizontal",
        length=350,
        mode="determinate",
        variable=progress_var,
        maximum=total_steps
    )
    progress_bar.pack(pady=20)

    # 添加状态标签
    status_label = tk.Label(progress_root, text="准备开始...", font=("Arial", 10))
    status_label.pack(pady=5)

    # 添加取消按钮
    cancel_button = tk.Button(
        progress_root,
        text="取消操作",
        command=lambda: progress_root.destroy()
    )
    cancel_button.pack(pady=10)

    # 更新进度条
    def update_progress(step, message):
        if not progress_root.winfo_exists():  # 检查窗口是否被关闭
            return False
        progress_var.set(step)
        status_label.config(text=message)
        progress_root.update()
        return True

    return progress_root, update_progress


def perform_batch_rename(folder_path, keep_chars, new_extension, progress_window=None, update_progress=None):
    """执行批量重命名操作（可带进度更新）"""
    # 获取文件夹信息
    files = [f for f in Path(folder_path).iterdir() if f.is_file()]
    file_count = len(files)

    if file_count == 0:
        return 0, "文件夹中没有找到任何文件"

    # 处理每个文件
    renamed_count = 0
    for idx, file in enumerate(files):
        # 更新进度条（如果提供）
        if progress_window and update_progress:
            if not update_progress(idx + 1, f"重命名: {file.name}..."):
                return renamed_count, "用户取消了操作"

        # 获取文件名（不含扩展名）
        stem = file.stem[:keep_chars]  # 只取前指定字符数

        # 构建新文件名
        new_path = file.with_name(f"{stem}.{new_extension}")

        # 处理文件名冲突
        counter = 1
        while new_path.exists():
            new_path = file.with_name(f"{stem}_{counter}.{new_extension}")
            counter += 1

        # 重命名文件
        file.rename(new_path)
        renamed_count += 1

    return renamed_count, None


def perform_merge(folder_path, progress_window=None, update_progress=None):
    """执行合并操作（可带进度更新）"""
    # 获取所有txt文件并按名称排序
    txt_files = []
    for file in os.listdir(folder_path):
        if file.lower().endswith('.txt'):
            txt_files.append(file)
    txt_files.sort()

    if not txt_files:
        return 0, None, "文件夹中没有找到TXT文件"

    # 创建合并后的文件名（带时间戳）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(folder_path, f"合并结果_{timestamp}.txt")

    # 合并文件并添加分隔符
    with open(output_file, 'w', encoding='utf-8') as outfile:
        for i, filename in enumerate(txt_files):
            # 更新进度条（如果提供）
            if progress_window and update_progress:
                if not update_progress(i + 1, f"合并: {filename} ({i + 1}/{len(txt_files)})"):
                    return len(txt_files), output_file, "用户取消了操作"

            filepath = os.path.join(folder_path, filename)

            # 添加文件分隔标记
            if i > 0:  # 第一个文件前不加分隔符
                separator = f"\n\n{'=' * 60}\n"
                separator += f"文件: {filename} (开始)\n"
                separator += f"{'=' * 60}\n\n"
                outfile.write(separator)

            # 写入文件内容
            try:
                with open(filepath, 'r', encoding='utf-8') as infile:
                    content = infile.read()
                    outfile.write(content)

                    # 添加文件结束标记
                    end_marker = f"\n\n{'=' * 60}\n"
                    end_marker += f"文件: {filename} (结束)\n"
                    end_marker += f"{'=' * 60}\n"
                    outfile.write(end_marker)
            except UnicodeDecodeError:
                try:
                    with open(filepath, 'r', encoding='gbk') as infile:
                        content = infile.read()
                        outfile.write(content)
                        # 同上添加结束标记
                        end_marker = f"\n\n{'=' * 60}\n"
                        end_marker += f"文件: {filename} (结束)\n"
                        end_marker += f"{'=' * 60}\n"
                        outfile.write(end_marker)
                except Exception as e:
                    error_msg = f"\n\n[错误] 无法读取文件: {filename} - {str(e)}\n\n"
                    outfile.write(error_msg)
            except Exception as e:
                error_msg = f"\n\n[错误] 处理文件时出错: {filename} - {str(e)}\n\n"
                outfile.write(error_msg)

    return len(txt_files), output_file, None


def rename_files(parent=None):
    """批量重命名文件并转换为指定格式"""
    # 创建图形界面
    root = parent or tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 选择文件夹对话框
    folder_path = filedialog.askdirectory(title="选择要处理的文件夹")
    if not folder_path:
        print("未选择文件夹，操作取消")
        if parent: parent.deiconify()
        return

    # 获取文件夹信息
    files = [f for f in Path(folder_path).iterdir() if f.is_file()]
    file_count = len(files)

    if file_count == 0:
        messagebox.showwarning("警告", "文件夹中没有找到任何文件")
        if parent: parent.deiconify()
        return

    # 获取用户输入：保留字符数和扩展名
    keep_chars = simpledialog.askinteger(
        "重命名设置",
        "请输入要保留的文件名前几位字符:",
        initialvalue=5,
        minvalue=1,
        parent=root
    )

    if keep_chars is None:  # 用户取消
        if parent: parent.deiconify()
        return

    new_extension = simpledialog.askstring(
        "重命名设置",
        "请输入新的文件扩展名 (例如: txt, jpg):",
        initialvalue="txt",
        parent=root
    )

    if new_extension is None:  # 用户取消
        if parent: parent.deiconify()
        return

    # 处理扩展名格式（确保没有点）
    new_extension = new_extension.strip().lstrip('.')
    if not new_extension:
        new_extension = "txt"  # 默认扩展名

    # 添加确认步骤
    confirm = messagebox.askyesno(
        "确认操作",
        f"将在以下文件夹执行批量重命名操作：\n{folder_path}\n\n"
        f"共发现 {file_count} 个文件\n"
        f"所有文件将被重命名为: 前{keep_chars}个字符 + '.{new_extension}'\n\n"
        "是否继续？",
        parent=root
    )
    if not confirm:
        messagebox.showinfo("操作取消", "批量重命名操作已取消", parent=root)
        if parent: parent.deiconify()
        return

    # 创建进度窗口
    progress_root, update_progress = create_progress_window("重命名文件", file_count, root)

    # 执行批量重命名
    renamed_count, error = perform_batch_rename(folder_path, keep_chars, new_extension, progress_root, update_progress)

    # 关闭进度窗口
    progress_root.destroy()

    # 显示结果
    if error:
        messagebox.showerror("操作失败", error, parent=root)
    else:
        result = f"操作完成！已将 {renamed_count} 个文件转换为 .{new_extension} 格式\n"
        result += f"处理后的文件夹: {folder_path}"
        messagebox.showinfo("操作完成", result, parent=root)

    # 任务完成后重新显示主菜单
    if parent:
        parent.deiconify()


def merge_txt_files(parent=None):
    """合并多个TXT文件"""
    # 创建隐藏主窗口
    root = parent or tk.Tk()
    root.withdraw()

    # 选择文件夹对话框
    folder_path = filedialog.askdirectory(title="选择包含TXT文件的文件夹")
    if not folder_path:
        print("未选择文件夹，操作取消")
        if parent: parent.deiconify()
        return

    # 获取所有txt文件并按名称排序
    txt_files = []
    for file in os.listdir(folder_path):
        if file.lower().endswith('.txt'):
            txt_files.append(file)

    if not txt_files:
        messagebox.showwarning("警告", "文件夹中没有找到TXT文件", parent=root)
        if parent: parent.deiconify()
        return

    # 添加确认步骤
    file_list = "\n".join(txt_files[:5])  # 显示前5个文件名
    if len(txt_files) > 5:
        file_list += f"\n...等共 {len(txt_files)} 个文件"

    confirm = messagebox.askyesno(
        "确认操作",
        f"将在以下文件夹执行合并操作：\n{folder_path}\n\n"
        f"将合并以下文件：\n{file_list}\n\n"
        "是否继续？",
        parent=root
    )
    if not confirm:
        messagebox.showinfo("操作取消", "合并操作已取消", parent=root)
        if parent: parent.deiconify()
        return

    # 创建进度窗口
    progress_root, update_progress = create_progress_window("合并TXT文件", len(txt_files), root)

    # 执行合并操作
    merged_count, output_file, error = perform_merge(folder_path, progress_root, update_progress)

    # 关闭进度窗口
    progress_root.destroy()

    # 显示结果
    if error:
        messagebox.showerror("操作失败", error, parent=root)
    else:
        result = f"合并完成！共合并 {merged_count} 个文件\n"
        result += f"输出文件: {output_file}"
        messagebox.showinfo("操作完成", result, parent=root)

    # 任务完成后重新显示主菜单
    if parent:
        parent.deiconify()


def rename_single_file(parent=None):
    """重命名单个文件"""
    # 创建图形界面
    root = parent or tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 选择文件对话框
    file_path = filedialog.askopenfilename(title="选择要重命名的文件")
    if not file_path:
        print("未选择文件，操作取消")
        if parent: parent.deiconify()
        return

    file_path = Path(file_path)

    # 检查文件是否存在
    if not file_path.is_file():
        messagebox.showerror("错误", "选择的文件不存在", parent=root)
        if parent: parent.deiconify()
        return

    # 获取原始文件名信息
    original_name = file_path.name
    original_stem = file_path.stem
    original_extension = file_path.suffix

    # 创建输入对话框获取新文件名设置
    keep_chars = simpledialog.askinteger(
        "设置文件名",
        "请输入要保留的文件名前几位字符:",
        initialvalue=5,
        minvalue=1,
        maxvalue=len(original_stem),
        parent=root
    )

    if keep_chars is None:  # 用户取消
        if parent: parent.deiconify()
        return

    new_extension = simpledialog.askstring(
        "设置文件扩展名",
        "请输入新的文件扩展名 (例如: txt, jpg):",
        initialvalue=original_extension.lstrip('.'),
        parent=root
    )

    if new_extension is None:  # 用户取消
        if parent: parent.deiconify()
        return

    # 处理扩展名格式（确保没有点）
    new_extension = new_extension.strip().lstrip('.')
    if not new_extension:
        new_extension = "txt"  # 默认扩展名

    # 构建新文件名
    new_stem = original_stem[:keep_chars]
    new_file_name = f"{new_stem}.{new_extension}"
    new_path = file_path.parent / new_file_name

    # 处理文件名冲突
    counter = 1
    while new_path.exists():
        new_file_name = f"{new_stem}_{counter}.{new_extension}"
        new_path = file_path.parent / new_file_name
        counter += 1

    # 显示确认信息
    confirm = messagebox.askyesno(
        "确认重命名",
        f"原始文件名: {original_name}\n"
        f"新文件名: {new_file_name}\n\n"
        "是否执行重命名操作？",
        parent=root
    )

    if not confirm:
        messagebox.showinfo("操作取消", "文件重命名操作已取消", parent=root)
        if parent: parent.deiconify()
        return

    # 执行重命名
    try:
        file_path.rename(new_path)
        messagebox.showinfo(
            "操作成功",
            f"文件重命名成功！\n\n"
            f"原始文件: {original_name}\n"
            f"新文件: {new_file_name}",
            parent=root
        )
    except Exception as e:
        messagebox.showerror(
            "操作失败",
            f"文件重命名失败:\n{str(e)}",
            parent=root
        )

    # 任务完成后重新显示主菜单
    if parent:
        parent.deiconify()


def batch_rename_and_merge(parent=None):
    """批量重命名并合并文件"""
    # 创建图形界面
    root = parent or tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 选择文件夹对话框
    folder_path = filedialog.askdirectory(title="选择要处理的文件夹")
    if not folder_path:
        print("未选择文件夹，操作取消")
        if parent: parent.deiconify()
        return

    # 获取文件夹信息
    files = [f for f in Path(folder_path).iterdir() if f.is_file()]
    file_count = len(files)

    if file_count == 0:
        messagebox.showwarning("警告", "文件夹中没有找到任何文件", parent=root)
        if parent: parent.deiconify()
        return

    # 获取用户输入：保留字符数和扩展名
    keep_chars = simpledialog.askinteger(
        "重命名设置",
        "请输入要保留的文件名前几位字符:",
        initialvalue=5,
        minvalue=1,
        parent=root
    )

    if keep_chars is None:  # 用户取消
        if parent: parent.deiconify()
        return

    new_extension = simpledialog.askstring(
        "重命名设置",
        "请输入新的文件扩展名 (例如: txt, jpg):",
        initialvalue="txt",
        parent=root
    )

    if new_extension is None:  # 用户取消
        if parent: parent.deiconify()
        return

    # 处理扩展名格式（确保没有点）
    new_extension = new_extension.strip().lstrip('.')
    if not new_extension:
        new_extension = "txt"  # 默认扩展名

    # 添加确认步骤
    confirm = messagebox.askyesno(
        "确认操作",
        f"将在以下文件夹执行批量操作：\n{folder_path}\n\n"
        f"共发现 {file_count} 个文件\n\n"
        "操作步骤：\n"
        f"1. 所有文件将被重命名为前{keep_chars}个字符的{new_extension}文件\n"
        f"2. 所有{new_extension}文件将被合并为一个文件\n\n"
        "是否继续？",
        parent=root
    )
    if not confirm:
        messagebox.showinfo("操作取消", "批量操作已取消", parent=root)
        if parent: parent.deiconify()
        return

    # 创建总进度窗口
    total_steps = file_count + 1  # 重命名 + 合并（合并作为一步）
    progress_root, update_progress = create_progress_window("批量重命名并合并", total_steps, root)

    # 第一步：执行批量重命名
    update_progress(1, "正在重命名文件...")
    renamed_count, error = perform_batch_rename(folder_path, keep_chars, new_extension)

    if error:
        progress_root.destroy()
        messagebox.showerror("操作失败", f"重命名失败: {error}", parent=root)
        if parent: parent.deiconify()
        return

    # 第二步：执行合并操作
    update_progress(file_count + 1, f"正在合并 .{new_extension} 文件...")
    merged_count, output_file, merge_error = perform_merge(folder_path)

    # 关闭进度窗口
    progress_root.destroy()

    # 显示结果
    if merge_error:
        messagebox.showerror("操作失败", f"合并失败: {merge_error}", parent=root)
    else:
        result = f"批量操作完成！\n\n"
        result += f"重命名: 已将 {renamed_count} 个文件转换为 .{new_extension} 格式\n"
        result += f"合并: 共合并 {merged_count} 个文件\n"
        result += f"输出文件: {output_file}"
        messagebox.showinfo("操作完成", result, parent=root)

    # 任务完成后重新显示主菜单
    if parent:
        parent.deiconify()


def show_main_menu():
    """显示主菜单界面"""
    # 创建主窗口
    root = tk.Tk()
    root.title("文件处理工具")
    root.geometry("300x300")

    # 设置窗口图标（可选）
    try:
        root.iconbitmap("icon.ico")  # 如果有图标文件可以添加
    except:
        pass

    # 创建标题标签
    title_label = tk.Label(root, text="请选择要执行的功能", font=("Arial", 14))
    title_label.pack(pady=10)

    # 创建功能按钮框架
    frame = tk.Frame(root)
    frame.pack(pady=5)

    # 批量重命名按钮
    batch_rename_btn = tk.Button(
        frame,
        text="批量重命名并转换",
        command=lambda: [root.withdraw(), rename_files(root)],
        width=25,
        height=2
    )
    batch_rename_btn.pack(pady=5)

    # 合并文件按钮
    merge_btn = tk.Button(
        frame,
        text="合并TXT文件",
        command=lambda: [root.withdraw(), merge_txt_files(root)],
        width=25,
        height=2
    )
    merge_btn.pack(pady=5)

    # 单文件重命名按钮
    single_rename_btn = tk.Button(
        frame,
        text="单文件重命名",
        command=lambda: [root.withdraw(), rename_single_file(root)],
        width=25,
        height=2
    )
    single_rename_btn.pack(pady=5)

    # 批量重命名并合并按钮
    batch_merge_btn = tk.Button(
        frame,
        text="批量重命名并合并",
        command=lambda: [root.withdraw(), batch_rename_and_merge(root)],
        width=25,
        height=2,
        bg="#e0e0ff"  # 浅蓝色背景突出显示
    )
    batch_merge_btn.pack(pady=5)

    # 添加退出按钮
    exit_btn = tk.Button(
        root,
        text="退出程序",
        command=root.destroy,
        width=15,
        height=1
    )
    exit_btn.pack(pady=10)

    # 运行主循环
    root.mainloop()


if __name__ == "__main__":
    show_main_menu()
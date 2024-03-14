import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox
import queue  

# 设置环境变量
RESOLVE_SCRIPT_API = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting"
RESOLVE_SCRIPT_LIB = "/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so"
PYTHONPATH = "$PYTHONPATH:$RESOLVE_SCRIPT_API/Modules/"

# 把上面的字典添加脚本库路径
sys.path.append(RESOLVE_SCRIPT_API)
sys.path.append(RESOLVE_SCRIPT_API + "/Modules/")

try:
    import DaVinciResolveScript as dvr_script
except ImportError:
    print("未找到 DaVinciResolveScript这个模块。请确保已安装 DaVinci Resolve 并正确配置脚本环境。")
    sys.exit(1)

# 把时码转换成帧
def timecode_to_frames(timecode, frame_rate):
    h, m, s, f = map(int, timecode.split(':'))
    return h * 3600 * frame_rate + m * 60 * frame_rate + s * frame_rate + f

def frames_to_timecode(frames, frame_rate):
    h = frames // (3600 * frame_rate)
    frames %= (3600 * frame_rate)
    m = frames // (60 * frame_rate)
    frames %= (60 * frame_rate)
    s = frames // frame_rate
    f = frames % frame_rate
    return f"{h:02}:{m:02}:{s:02}:{f:02}"

# 获取DaVinci Resolve实例和当前时间线
def get_resolve():
    resolve = dvr_script.scriptapp("Resolve")
    if resolve is None:
        print("未能连接到DaVinci Resolve。请确保Resolve正在运行。")
        sys.exit(1)
    return resolve

def get_current_timeline(resolve):
    project_manager = resolve.GetProjectManager()
    project = project_manager.GetCurrentProject()
    timeline = project.GetCurrentTimeline()
    if timeline is None:
        print("未能获取当前时间线。请确保你有一个打开的项目，并且时间线是可用的。")
        sys.exit(1)
    return timeline

# 来到下一个剪辑点
def get_next_edit_point(timeline, current_frames, skip_frame, track_type="video", track_index=1):
    items = timeline.GetItemsInTrack(track_type, track_index)
    next_edit_frame = None
    for i, item in enumerate(items.values()):
        start_frame = item.GetStart()
        if start_frame > skip_frame:
            next_edit_frame = start_frame
            break
    return next_edit_frame

# 主处理函数
def process(frame_rate, interval, stop_event, message_queue):
    resolve = get_resolve()
    timeline = get_current_timeline(resolve)

    # 获取当前时间线上的当前时间码
    current_timecode = timeline.GetCurrentTimecode()
    current_frames = timecode_to_frames(current_timecode, frame_rate)

    # 使用当前帧作为下一个要检查的帧
    next_frame_to_check = current_frames

    while not stop_event.is_set():
        # 检索下一个剪辑点
        next_edit_frame = get_next_edit_point(timeline, next_frame_to_check, current_frames)

        if next_edit_frame is not None:
            # 移动到下一个剪辑点
            timeline.SetCurrentTimecode(frames_to_timecode(next_edit_frame, frame_rate))
            print(f"Moved to next edit point at frame: {next_edit_frame}")
            time.sleep(interval)  # 使用用户指定的时间间隔

            # 往前挪一帧
            previous_frame = next_edit_frame - 1
            timeline.SetCurrentTimecode(frames_to_timecode(previous_frame, frame_rate))
            print(f"Moved back one frame to: {previous_frame}")
            time.sleep(interval)  # 使用用户指定的时间间隔

            # 更新下一次要检查的帧和跳过帧的值
            next_frame_to_check = next_edit_frame + 1
            current_frames = next_edit_frame
        else:
            message_queue.put("结束啦！")  # 将结束消息放入队列
            break

# Tkinter的大界面
class App:
    def __init__(self, root):
        self.root = root
        root.title('自动质检小助手')
        root.geometry("400x250")  # 设置窗口大小

        self.frame_rate_var = tk.StringVar()
        self.frame_rate_label = tk.Label(root, text='输入帧率 (例如25):')
        self.frame_rate_label.pack(pady=5)  # 增加上下间距

        self.frame_rate_entry = tk.Entry(root, textvariable=self.frame_rate_var)
        self.frame_rate_entry.pack(pady=5)

        self.interval_var = tk.StringVar()
        self.interval_label = tk.Label(root, text='输入切点跳转的时间间隔 (秒，例如0.5):')
        self.interval_label.pack(pady=5)

        self.interval_entry = tk.Entry(root, textvariable=self.interval_var)
        self.interval_entry.pack(pady=5)

        self.start_button = tk.Button(root, text='开始', command=self.start_processing)
        self.start_button.pack(pady=5)

        self.stop_button = tk.Button(root, text='结束', command=self.stop_processing)
        self.stop_button.pack(pady=5)

        self.process_thread = None
        self.stop_event = threading.Event()
        self.message_queue = queue.Queue()  # 创建一个消息队列
        self.root.after(100, self.check_message_queue)  # 每隔100毫秒检查一次消息队列

    def start_processing(self):
        frame_rate_str = self.frame_rate_var.get()
        interval_str = self.interval_var.get()
        try:
            frame_rate = int(frame_rate_str)
            interval = float(interval_str)
            self.stop_event.clear()
            self.process_thread = threading.Thread(target=process, args=(frame_rate, interval, self.stop_event, self.message_queue))
            self.process_thread.start()
        except ValueError:
            messagebox.showwarning('不对', '请输入有效的帧率和时间间隔!')

    def stop_processing(self):
        if self.process_thread and self.process_thread.is_alive():
            self.stop_event.set()
            self.process_thread.join()
        messagebox.showinfo('信息', '处理已结束或暂停!')

    def check_message_queue(self):
        """检查消息队列，并显示任何消息"""
        try:
            while True:  # 尝试获取所有消息
                message = self.message_queue.get_nowait()
                messagebox.showinfo('信息', message)
        except queue.Empty:  # 如果队列为空，捕获异常
            pass
        finally:
            self.root.after(100, self.check_message_queue)  # 继续定期巡检队列

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

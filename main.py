import pyautogui
import time
import threading
import sys
from datetime import datetime
from CursorProcess import *

# 解决Python 3.6编码问题
try:
    sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)
except Exception:
    pass

# 配置参数
DIRECTION_CHANGE_THRESHOLD = 3  # 触发晃动的方向改变次数
MIN_MOVE_DISTANCE = 30  # 单次最小移动距离(像素)
DETECTION_WINDOW = 1.5  # 检测窗口时间(秒)
CHECK_INTERVAL = 0.03  # 检测间隔(秒)
COOLDOWN_AFTER_STOP = 0.5  # 晃动停止后冷却时间(秒)
RUNNING = True


def changeCursor():
    # 2. 生成标准放大CUR（缩小scale至3倍，避免尺寸超限）
    scaled_cur_path = "D:\\pycode\\findMyCursor\\1.cur"
    original_path = get_arrow_cursor_path()
    if not scaled_cur_path:

        if not original_path:
            print("无法获取原始光标，退出")
            exit(1)
        print(f"原始光标：{original_path}")
        resize_and_convert_cursor(original_path, 2, scaled_cur_path)
        # exit(0)
        if not scaled_cur_path:
            print("生成放大光标失败，退出")
            exit(1)

    # 3. 设置并还原
    print("设置放大光标...")
    if set_arrow_cursor(scaled_cur_path):
        print("放大光标生效，5秒后还原...")
        time.sleep(1)
        print("还原原始光标...")
        set_arrow_cursor(original_path)
        print("还原完成")
    else:
        print("设置光标失败")


def monitor_shaking():
    """带防抖节流的鼠标来回晃动监控"""
    prev_x, prev_y = pyautogui.position()

    # 水平方向变量
    h_changes = []  # 方向改变记录 [(时间戳, 方向)]
    last_h_dir = 0  # 上一次水平方向 (1:右, -1:左)
    h_shaking_detected = False  # 水平晃动已检测标记

    # 垂直方向变量
    v_changes = []  # 方向改变记录 [(时间戳, 方向)]
    last_v_dir = 0  # 上一次垂直方向 (1:下, -1:上)
    v_shaking_detected = False  # 垂直晃动已检测标记

    # 防抖节流控制
    last_move_time = time.time()  # 最后一次有效移动时间
    in_cooldown = False  # 冷却状态标记

    print("带防抖节流的鼠标晃动监控已启动")
    print(f"判定条件: {DIRECTION_CHANGE_THRESHOLD}次方向改变（{DETECTION_WINDOW}秒内）")
    print("特点：一次晃动过程只输出一次结果，按Ctrl+C停止")

    try:
        while RUNNING:
            curr_x, curr_y = pyautogui.position()
            current_time = time.time()
            dx = curr_x - prev_x
            dy = curr_y - prev_y
            moved = False  # 标记是否有有效移动

            # 水平方向检测
            if abs(dx) > MIN_MOVE_DISTANCE:
                moved = True
                last_move_time = current_time
                curr_h_dir = 1 if dx > 0 else -1

                # 检测方向改变
                if last_h_dir != 0 and curr_h_dir != last_h_dir:
                    h_changes.append((current_time, curr_h_dir))
                    # 清理过期记录
                    h_changes = [
                        (t, d) for t, d in h_changes
                        if current_time - t <= DETECTION_WINDOW
                    ]
                    last_h_dir = curr_h_dir
                elif last_h_dir == 0:
                    last_h_dir = curr_h_dir

            # 垂直方向检测
            if abs(dy) > MIN_MOVE_DISTANCE:
                moved = True
                last_move_time = current_time
                curr_v_dir = 1 if dy > 0 else -1

                # 检测方向改变
                if last_v_dir != 0 and curr_v_dir != last_v_dir:
                    v_changes.append((current_time, curr_v_dir))
                    # 清理过期记录
                    v_changes = [
                        (t, d) for t, d in v_changes
                        if current_time - t <= DETECTION_WINDOW
                    ]
                    last_v_dir = curr_v_dir
                elif last_v_dir == 0:
                    last_v_dir = curr_v_dir

            # 检查是否符合晃动条件且不在冷却中
            if not in_cooldown:
                # 水平晃动判定（未触发过且次数达标）
                if not h_shaking_detected and len(h_changes) >= DIRECTION_CHANGE_THRESHOLD:
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    print(f"[{timestamp}] 检测到左右来回晃动（一次）")
                    h_shaking_detected = True  # 标记为已触发
                    changecursor_time = time.time()  # 最后一次有效移动时间
                    changeCursor()
                    print(f"执行{time.time() - changecursor_time}")
                    # v_shaking_detected = False

                # 垂直晃动判定（未触发过且次数达标）
                if not v_shaking_detected and len(v_changes) >= DIRECTION_CHANGE_THRESHOLD:
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    print(f"[{timestamp}] 检测到上下来回晃动（一次）")
                    v_shaking_detected = True  # 标记为已触发
                    changeCursor()
                    # v_shaking_detected = False

            # 防抖逻辑：当鼠标停止移动超过冷却时间后重置检测状态

            if current_time - last_move_time > COOLDOWN_AFTER_STOP:
                # print(f"已冷却时间{current_time - last_move_time}")
                # 只有当之前检测到过晃动才需要重置
                if h_shaking_detected or v_shaking_detected:
                    in_cooldown = False
                    h_shaking_detected = False
                    v_shaking_detected = False
                    h_changes = []
                    v_changes = []
                    last_h_dir = 0
                    last_v_dir = 0

            # 更新上一位置
            prev_x, prev_y = curr_x, curr_y
            time.sleep(CHECK_INTERVAL)

    except Exception as e:
        print(f"监控出错: {e}")


def handle_exit():
    """处理程序退出"""
    global RUNNING
    try:
        while RUNNING:
            time.sleep(0.1)
    except KeyboardInterrupt:
        RUNNING = False
        print("\n正在停止监控...")


if __name__ == "__main__":
    monitor_thread = threading.Thread(target=monitor_shaking, daemon=True)
    monitor_thread.start()
    handle_exit()
    monitor_thread.join()
    print("监控已停止")

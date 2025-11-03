基于python3.6
# 配置参数
DIRECTION_CHANGE_THRESHOLD = 3  # 触发晃动的方向改变次数
MIN_MOVE_DISTANCE = 30  # 单次最小移动距离(像素)
DETECTION_WINDOW = 1.5  # 检测窗口时间(秒)
CHECK_INTERVAL = 0.03  # 检测间隔(秒)
COOLDOWN_AFTER_STOP = 0.5  # 晃动停止后冷却时间(秒)

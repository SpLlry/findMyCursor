import winreg
import os
import subprocess
from PIL import Image
import ctypes
from ctypes import wintypes, Structure, c_uint16, c_uint32, byref
import struct
import win32gui
import win32con
import win32ui
import ctypes
from ctypes import wintypes

# 加载Windows系统库
user32 = ctypes.WinDLL('user32', use_last_error=True)
gdi32 = ctypes.WinDLL('gdi32', use_last_error=True)


# 定义CUR文件头部结构（必须符合Windows规范）
class ICONDIRENTRY(Structure):
    _fields_ = [
        ("bWidth", c_uint16),  # 宽度（1-255）
        ("bHeight", c_uint16),  # 高度（1-255）
        ("bColorCount", c_uint16),  # 颜色数（0=256色）
        ("bReserved", c_uint16),  # 保留（必须为0）
        ("wPlanes", c_uint16),  # 平面数（必须为1）
        ("wBitCount", c_uint16),  # 位深度（32=RGBA）
        ("dwBytesInRes", c_uint32),  # 图像数据大小
        ("dwImageOffset", c_uint32)  # 图像数据偏移量
    ]


class ICONDIR(Structure):
    _fields_ = [
        ("idReserved", c_uint16),  # 保留（必须为0）
        ("idType", c_uint16),  # 类型（2=CUR光标）
        ("idCount", c_uint16),  # 图标数量（1=单个光标）
        ("idEntries", ICONDIRENTRY * 1)  # 光标信息
    ]


# 系统消息常量
WM_SETTINGCHANGE = 0x001A
SPI_SETCURSORS = 0x0057
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02


def expand_environment_vars(path):
    """解析环境变量"""
    if not path:
        return ""
    try:
        result = subprocess.check_output(
            f'echo {path}', shell=True, stderr=subprocess.STDOUT, universal_newlines=True
        ).strip()
        return result
    except Exception:
        if "%SYSTEMROOT%" in path:
            return path.replace("%SYSTEMROOT%", os.environ.get("SYSTEMROOT", "C:\\Windows"))
        return path


def get_arrow_cursor_path():
    """获取原始Arrow光标路径"""
    try:
        reg_path = r"Control Panel\Cursors"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
        arrow_path, _ = winreg.QueryValueEx(key, "Arrow")
        winreg.CloseKey(key)
        return expand_environment_vars(arrow_path)
    except Exception as e:
        print(f"获取原始路径失败：{e}")
        return None


def png_to_cur(png_path, cur_path, hotspot=(0, 0)):
    """
    将PNG图像转换为CUR鼠标指针文件
    """
    # 打开PNG图像
    img = Image.open(png_path)

    # 确保图像是32x32或更小(标准鼠标指针尺寸)
    if img.width > 32 or img.height > 32:
        # 处理不同版本的Pillow
        try:
            resample_filter = Image.ANTIALIAS
        except AttributeError:
            resample_filter = Image.LANCZOS
        img = img.resize((32, 32), resample_filter)

    # 转换为RGBA模式
    img = img.convert("RGBA")

    # 创建CUR文件
    with open(cur_path, 'wb') as f:
        # CUR文件头 (6字节)
        f.write(struct.pack('<HHH', 0, 2, 1))

        # 图像目录项 (16字节)
        width = img.width if img.width < 256 else 0
        height = img.height if img.height < 256 else 0
        f.write(struct.pack('BBB', width, height, 0))
        f.write(struct.pack('B', 0))
        f.write(struct.pack('<HH', hotspot[0], hotspot[1]))
        image_data_size = img.width * img.height * 4 + 40
        f.write(struct.pack('<I', image_data_size))
        f.write(struct.pack('<I', 22))

        # 写入位图信息头 (40字节)
        f.write(struct.pack('<I', 40))
        f.write(struct.pack('<I', img.width))
        f.write(struct.pack('<I', img.height * 2))
        f.write(struct.pack('<H', 1))
        f.write(struct.pack('<H', 32))
        f.write(struct.pack('<I', 0))
        f.write(struct.pack('<I', img.width * img.height * 4))
        f.write(struct.pack('<I', 0))
        f.write(struct.pack('<I', 0))
        f.write(struct.pack('<I', 0))
        f.write(struct.pack('<I', 0))

        # 写入像素数据 (BGRA格式)
        pixels = []
        for y in range(img.height - 1, -1, -1):  # CUR文件从下到上存储
            for x in range(img.width):
                r, g, b, a = img.getpixel((x, y))
                pixels.extend([b, g, r, a])  # 转换为BGRA格式

        f.write(bytes(pixels))

        # 创建正确的AND掩码
        and_mask = []
        for y in range(img.height - 1, -1, -1):
            for x in range(0, img.width, 8):
                byte_val = 0
                for bit in range(8):
                    if x + bit < img.width:
                        r, g, b, a = img.getpixel((x + bit, y))
                        # 如果alpha通道为0(完全透明)，则AND位设为1
                        if a == 0:
                            byte_val |= (1 << (7 - bit))
                and_mask.append(byte_val)

        # 确保AND掩码大小正确(每行需要4字节对齐)
        row_size = ((img.width + 31) // 32) * 4
        and_mask_padded = []
        for i in range(0, len(and_mask), img.width // 8 if img.width % 8 == 0 else img.width // 8 + 1):
            row = and_mask[i:i + (img.width // 8 if img.width % 8 == 0 else img.width // 8 + 1)]
            and_mask_padded.extend(row)
            # 填充到4字节边界
            while len(and_mask_padded) % 4 != 0:
                and_mask_padded.append(0)

        f.write(bytes(and_mask_padded))


def resize_and_convert_cursor(original_path, scale=3, output_name="scaled_arrow"):
    """放大图像并转换为标准CUR"""
    # 1. 放大图像保存为PNG
    output_name = output_name.split(".")[0]
    png_path = os.path.join(os.getcwd(), f"{output_name}.png")
    png_path2 = os.path.join(os.getcwd(), f"{output_name}2.png")
    try:
        img_32, (hot_x, hot_y) = get_cursor_hotspot(original_path)
        if img_32.size != (32, 32):
            raise ValueError("光标文件中未找到32x32尺寸")

        with Image.open(original_path) as img:

            # 第一步：按比例放大，限制最大尺寸255x255
            new_width = img.width * scale
            new_height = img.height * scale
            # 限制最大尺寸为255x255
            new_width = min(new_width, 255)
            new_height = min(new_height, 255)
            # 确保放大后尺寸至少为64x64（若原图太小，强制放大到64x64）
            new_width = max(new_width, 64)
            new_height = max(new_height, 64)
            # 转换为整数尺寸
            new_size = (int(new_width), int(new_height))

            # 执行放大
            resized_img = img.resize(new_size, Image.LANCZOS)  # 若报错用Image.ANTIALIAS

            # 第二步：从放大后的图片（至少64x64）中裁剪出32x32
            # 计算中心裁剪区域（left, upper, right, lower）
            crop_left = hot_x * scale
            crop_upper = hot_y * scale
            crop_right = crop_left + 32
            crop_lower = crop_upper + 32
            # 执行裁剪
            cropped_img = resized_img.crop((crop_left, crop_upper, crop_right, crop_lower))

            # 保存最终裁剪后的32x32图片
            cropped_img.save(png_path, format='PNG')
        print(f"放大PNG已保存：{png_path}")
        # 主体放大2倍，透明区域放大0.3倍（相对缩小）
    except Exception as e:
        print(f"放大PNG失败：{e}")
        return None

    # 2. 将PNG转换为标准CUR
    cur_path = os.path.join(os.getcwd(), f"{output_name}.cur")
    if png_to_cur(png_path, cur_path):
        print(f"标准CUR已生成：{cur_path}")
        return cur_path
    else:
        return None


def get_cursor_hotspot(cursor_path):
    """解析光标文件，返回32x32图像和热点坐标（兼容Python 3.6）"""
    # 加载光标文件（使用win32gui的LoadImage）
    hcursor = win32gui.LoadImage(
        None,
        cursor_path,
        win32con.IMAGE_CURSOR,  # 类型：光标
        32, 32,  # 目标尺寸32x32
        win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
    )
    if not hcursor:
        raise Exception(f"无法加载光标文件：{cursor_path}")

    # 通过GetIconInfo获取热点坐标（光标本质是特殊的图标）
    class ICONINFO(ctypes.Structure):
        _fields_ = [
            ("fIcon", wintypes.BOOL),
            ("xHotspot", ctypes.c_int),
            ("yHotspot", ctypes.c_int),
            ("hbmMask", wintypes.HBITMAP),
            ("hbmColor", wintypes.HBITMAP)
        ]

    icon_info = ICONINFO()
    # 调用Windows API获取图标信息（包含热点）
    success = ctypes.windll.user32.GetIconInfo(hcursor, ctypes.byref(icon_info))
    if not success:
        raise Exception("无法获取光标热点信息")
    hotspot_x, hotspot_y = icon_info.xHotspot, icon_info.yHotspot

    # 提取32x32光标图像
    # 将光标句柄转换为位图
    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    hbmp = win32ui.CreateBitmap()
    hbmp.CreateCompatibleBitmap(hdc, 32, 32)
    hdc_mem = hdc.CreateCompatibleDC()
    hdc_mem.SelectObject(hbmp)
    # 绘制光标到位图
    win32gui.DrawIconEx(
        hdc_mem.GetSafeHdc(),
        0, 0,  # 绘制位置
        hcursor,
        32, 32,  # 尺寸
        0, None,
        win32con.DI_NORMAL  # 绘制模式
    )
    # 转换为PIL图像
    bmp_info = hbmp.GetInfo()
    bmp_bits = hbmp.GetBitmapBits(True)
    img = Image.frombuffer(
        "RGBA",
        (bmp_info["bmWidth"], bmp_info["bmHeight"]),
        bmp_bits,
        "raw",
        "BGRA",
        0, 1
    )
    # 清理资源
    hdc_mem.DeleteDC()
    win32gui.DeleteObject(hbmp.GetHandle())

    return img, (hotspot_x, hotspot_y)


def set_arrow_cursor(cursor_path):
    """设置Arrow光标并刷新"""
    try:
        reg_path = r"Control Panel\Cursors"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "Arrow", 0, winreg.REG_SZ, cursor_path)
        winreg.CloseKey(key)

        # 刷新光标设置
        user32.SystemParametersInfoW(SPI_SETCURSORS, 0, None, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)
        return True
    except Exception as e:
        print(f"设置光标失败：{e}")
        return False

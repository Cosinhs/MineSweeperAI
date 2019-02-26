#https://zhuanlan.zhihu.com/p/39669361
import win32gui
import win32api
#import win32com.client
import win32con
from PIL import ImageGrab
import random
import sys
import time

lclick_t = 0.1

level = "n"

if level == "e":
    blocks_x_num, blocks_y_num = 8, 8
elif level == "n":
    blocks_x_num, blocks_y_num = 16, 16

block_width, block_height = 16, 16

img_unknown = ((54, (255, 255, 255)), (148, (192, 192, 192)), (54, (128, 128, 128)))
img_flag = ((54, (255, 255, 255)), (17, (255, 0, 0)),
            (109, (192, 192, 192)), (54, (128, 128, 128)), (22, (0, 0, 0)))
img_0 = ((225, (192, 192, 192)), (31, (128, 128, 128)))
img_1 = ((185, (192, 192, 192)), (31, (128, 128, 128)), (40, (0, 0, 255)))
img_2 = ((160, (192, 192, 192)), (31, (128, 128, 128)), (65, (0, 128, 0)))
img_3 = ((62, (255, 0, 0)), (163, (192, 192, 192)), (31, (128, 128, 128)))
img_4 = ((169, (192, 192, 192)), (31, (128, 128, 128)), (56, (0, 0, 128)))
img_5 = ((70, (128, 0, 0)), (155, (192, 192, 192)), (31, (128, 128, 128)))
img_6 = ((153, (192, 192, 192)), (31, (128, 128, 128)), (72, (0, 128, 128)))
img_7 = ((181, (192, 192, 192)), (31, (128, 128, 128)), (44, (0, 0, 0)))
img_8 = ((149, (192, 192, 192)), (107, (128, 128, 128)))
img_boom = ((4, (255, 255, 255)), (144, (192, 192, 192)),
            (31, (128, 128, 128)), (77, (0, 0, 0)))
img_boom_red = ((4, (255, 255, 255)), (144, (255, 0, 0)),
                (31, (128, 128, 128)), (77, (0, 0, 0)))

img_dict = {img_0: 0, img_1: 1, img_2: 2, img_3: 3, img_4: 4,
            img_5: 5, img_6: 6, img_7: 7, img_8: 8, img_unknown: -1,
            img_flag: -2, img_boom: -3, img_boom_red: -4}

    
# 扫描雷区图像
def scan_map(rect):
    num_img = []
    unknown_img = []
    game_img = ImageGrab.grab().crop(rect)
    for i in range(blocks_y_num):
        for j in range(blocks_x_num):
            this_image = tuple(game_img.crop((j*block_width, i*block_height,
                                              (j+1)*block_width, (i+1)*block_height)).getcolors())
            if this_image not in img_dict:
                print("无法识别图像")
                print("坐标: ", (i, j))
                sys.exit(1)
            game_map[(i, j)] = img_dict[this_image]
            if game_map[(i, j)]==-3 or game_map[(i, j)]==-4:
                print("game over")
                sys.exit(0)
            if game_map[(i, j)] == -1:
                unknown_img.append((i, j))
            elif game_map[(i, j)] in range(1, 9) and \
                 (i, j) not in num_blocks_not_consider:
                num_img.append((i, j))
    return (num_img, unknown_img)

def get_around(i, j):
    t = []
    for tt in [(i-1, j-1), (i-1, j), (i-1, j+1), (i, j-1), (i, j+1), (i+1, j-1), (i+1, j), (i+1, j+1)]:
        if 0<=tt[0]<=blocks_y_num-1 and 0<=tt[1]<=blocks_x_num-1:
            t.append(tt)
    return t

def lclick(i, j):
    win32api.SetCursorPos([left+j*block_width, top+i*block_height])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def rclick(i, j):
    win32api.SetCursorPos([left+j*block_width, top+i*block_height])
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)


# 窗口坐标
left = 0
top = 0
right = 0
bottom = 0

class_name = "TMain"
title_name = "Minesweeper Arbiter "
hwnd = win32gui.FindWindow(class_name, title_name)

if hwnd:
    print("找到窗口")
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    #shell = win32com.client.Dispatch("WScript.Shell")
    #shell.SendKeys('%')
    win32gui.SetForegroundWindow(hwnd)
    print("窗口坐标：")
    print(str(left)+', '+str(top)+'  '+str(right)+', '+str(bottom))
else:
    print("未找到窗口")
    sys.exit(1)
    
# 锁定雷区坐标
# 去除周围功能按钮以及多余的界面
# 具体的像素值是通过QQ的截图来判断的
left += 15
top += 101
right -= 15
bottom -= 42

# 抓取雷区图像
rect = (left, top, right, bottom)

game_map = dict()
num, unknown = scan_map(rect)

num_blocks_not_consider = set()

while True:
    if len(unknown) == 0:
        print("win")
        sys.exit(0)

    no_choice_for_naive_strategy = True
        
    for t in num[:]:
        around_unknown_img = []  
        flag = 0  
        for e in get_around(*t):
            if game_map[e] == -1:  
                around_unknown_img.append(e)
            elif game_map[e] == -2:  
                flag += 1
        
        if len(around_unknown_img) == 0:
            num_blocks_not_consider.add(t)
        elif flag == game_map[t]: # 周围安全
            num_blocks_not_consider.add(t)
            no_choice_for_naive_strategy = False
            for e in around_unknown_img:
                lclick(*e)
                time.sleep(lclick_t)
                num, unknown = scan_map(rect)
        elif len(around_unknown_img) == game_map[t] - flag: # 说明周围全是雷，右键点击所有格
            num_blocks_not_consider.add(t)
            no_choice_for_naive_strategy = False
            for e in around_unknown_img:
                rclick(*e)
                game_map[e] = -2
                unknown.remove(e)

    if no_choice_for_naive_strategy:
        r = random.choice(unknown)
        lclick(*r)
        time.sleep(lclick_t)
        num, unknown = scan_map(rect)


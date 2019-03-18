#https://zhuanlan.zhihu.com/p/39669361
import win32gui
import win32api
#import win32com.client
import win32con
from PIL import Image, ImageDraw, ImageGrab
import os
import shutil
import random
import sys
import time
from utils import gauss, get_connected_parts, get_x_sol

click_t = 0.1
debug = False
level = "h"

if level == "e":
    blocks_x_num, blocks_y_num, left_mines = 8, 8, 10
elif level == "n":
    blocks_x_num, blocks_y_num, left_mines = 16, 16, 40
elif level == "h":
    blocks_x_num, blocks_y_num, left_mines = 30, 16, 99
    
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
            img_5: 5, img_6: 6, img_7: 7, img_8: 8, img_unknown: "unknown",
            img_flag: "flag", img_boom: "dead", img_boom_red: "dead"}

def scan_map(game_img):
    game_map = {}
    num_blocks = []
    unknown_blocks = []
    for i in range(blocks_y_num):
        for j in range(blocks_x_num):
            this_image = tuple(game_img.crop((j*block_width, i*block_height,
                                              (j+1)*block_width, (i+1)*block_height)).getcolors())
            if this_image not in img_dict:
                print("无法识别图像")
                print("坐标: ", (i, j))
                sys.exit(1)
            if img_dict[this_image]=="dead":
                print("game over")
                sys.exit(0)
            game_map[(i, j)] = img_dict[this_image]
            if game_map[(i, j)] == "unknown":
                unknown_blocks.append((i, j))
            elif game_map[(i, j)] in range(1, 9) and \
                 (i, j) not in num_block_not_try:
                num_blocks.append((i, j))
    return game_map, num_blocks, unknown_blocks

def get_around(i, j):
    t = []
    for tt in [(i-1, j-1), (i-1, j), (i-1, j+1), (i, j-1), (i, j+1), (i+1, j-1), (i+1, j), (i+1, j+1)]:
        if tt[0] in range(blocks_y_num) and tt[1] in range(blocks_x_num):
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
    
left += 15
top += 101
right -= 15
bottom -= 42
rect = (left, top, right, bottom)

debug_folder = os.getcwd() + '\\debug'
if os.path.exists(debug_folder):
    shutil.rmtree(debug_folder)
os.mkdir(debug_folder)

outer_loop = -1
num_block_not_try = set()

while True:
    outer_loop += 1
    no_choice_for_naive_strategy = True

    game_img = ImageGrab.grab(rect)
    game_map, trial, unknown = scan_map(game_img)

    if debug:
        os.mkdir(f'{debug_folder}\\{outer_loop}')
        game_img.save(f'{debug_folder}\\{outer_loop}\\0.jpg')
        
    if len(unknown) == 0:
        print("win")
        sys.exit(0)

    if len(trial) == 0:
        r = random.choice(unknown)
        lclick(*r)
        time.sleep(click_t)
        if debug:
            game_img = ImageGrab.grab(rect)
            draw = ImageDraw.Draw(game_img)
            draw.rectangle((r[1]*block_width, r[0]*block_height,
                          (r[1]+1)*block_width, (r[0]+1)*block_height),
                          None, 'yellow', 2)
            game_img.save(f'{debug_folder}\\{outer_loop}\\random.jpg')
        continue
        
    inner_loop = 0
    unknown_blocks_info = []
    x_to_solve = []
    for t in trial:
        inner_loop += 1
        around_unknown_blocks = []  
        flag = 0
        
        for e in get_around(*t):
            if game_map[e] == "unknown":  
                around_unknown_blocks.append(e)
            elif game_map[e] == "flag":  
                flag += 1
        
        if len(around_unknown_blocks) == 0:
            num_block_not_try.add(t)
            
            if debug:
                game_img = ImageGrab.grab(rect)
                draw = ImageDraw.Draw(game_img)
                draw.rectangle((t[1]*block_width, t[0]*block_height,
                              (t[1]+1)*block_width, (t[0]+1)*block_height),
                              None, 'pink', 2)
                game_img.save(f'{debug_folder}\\{outer_loop}\\{inner_loop}.jpg')

        elif flag == game_map[t]: # 周围安全
            no_choice_for_naive_strategy = False
            num_block_not_try.add(t)
            
            for e in around_unknown_blocks:
                lclick(*e)
                time.sleep(click_t)
            game_img = ImageGrab.grab(rect)
            game_map = scan_map(game_img)[0]
            if debug:
                draw = ImageDraw.Draw(game_img)
                draw.rectangle((t[1]*block_width, t[0]*block_height,
                              (t[1]+1)*block_width, (t[0]+1)*block_height),
                              None, 'blue', 2)
                game_img.save(f'{debug_folder}\\{outer_loop}\\{inner_loop}.jpg')

        elif len(around_unknown_blocks) == game_map[t] - flag: 
            no_choice_for_naive_strategy = False
            num_block_not_try.add(t)
            
            for e in around_unknown_blocks:
                rclick(*e)
                time.sleep(click_t)
                game_map[e] = "flag"
                left_mines -= 1
            if debug:
                game_img = ImageGrab.grab(rect)
                draw = ImageDraw.Draw(game_img)
                draw.rectangle((t[1]*block_width, t[0]*block_height,
                              (t[1]+1)*block_width, (t[0]+1)*block_height),
                              None, 'red', 2)
                game_img.save(f'{debug_folder}\\{outer_loop}\\{inner_loop}.jpg')

        else:
            if debug:
                game_img = ImageGrab.grab(rect)
                draw = ImageDraw.Draw(game_img)
                draw.rectangle((t[1]*block_width, t[0]*block_height,
                              (t[1]+1)*block_width, (t[0]+1)*block_height),
                              None, 'black', 2)
                game_img.save(f'{debug_folder}\\{outer_loop}\\{inner_loop}.jpg')

            if no_choice_for_naive_strategy:
                for _ in around_unknown_blocks:
                    if _ not in x_to_solve:
                        x_to_solve.append(_)
                unknown_blocks_info.append([around_unknown_blocks, game_map[t]-flag])

    if no_choice_for_naive_strategy:
        mat_to_solve = [] # size is len(trial) * (len(x_to_solve)+1)
        if len(x_to_solve) == 0:
            continue
        for _ in unknown_blocks_info:
            mat_to_solve.append([])
            for x in x_to_solve:
                if x in _[0]: 
                    mat_to_solve[-1].append(1)
                else:
                    mat_to_solve[-1].append(0)
            mat_to_solve[-1].append(_[1])
            
        if len(x_to_solve) == len(unknown):
            mat_to_solve.append([1]*len(x_to_solve))
            mat_to_solve[-1].append(left_mines)

        mat_to_solve, x_to_solve, fix, free = gauss(mat_to_solve, x_to_solve)                    
        connected_parts = get_connected_parts(mat_to_solve, fix, free)
        x_with_sol, x_with_possibility_sol = get_x_sol(\
            mat_to_solve, x_to_solve, connected_parts, left_mines)
        
        if len(x_with_sol) != 0:
            for x in x_with_sol:
                inner_loop += 1
                if x_with_sol[x] == 1:
                    rclick(*x)
                    time.sleep(click_t)
                    game_map[x] = "flag"
                    left_mines -= 1
                    if debug:
                        game_img = ImageGrab.grab(rect)
                        draw = ImageDraw.Draw(game_img)
                        draw.rectangle((x[1]*block_width, x[0]*block_height,
                                      (x[1]+1)*block_width, (x[0]+1)*block_height),
                                      None, 'red', 2)
                        game_img.save(f'{debug_folder}\\{outer_loop}\\{inner_loop}_.jpg')
                else:
                    if game_map[x] == "unknown":
                        lclick(*x)
                        time.sleep(click_t)
                        game_img = ImageGrab.grab(rect)
                        game_map = scan_map(game_img)[0]
                        if debug:
                            draw = ImageDraw.Draw(game_img)
                            draw.rectangle((x[1]*block_width, x[0]*block_height,
                                          (x[1]+1)*block_width, (x[0]+1)*block_height),
                                          None, 'blue', 2)
                            game_img.save(f'{debug_folder}\\{outer_loop}\\{inner_loop}_.jpg')
        else:
            t = sorted(x_with_possibility_sol, key = lambda x: x_with_possibility_sol[x])
            if x_with_possibility_sol[t[0]] < 0.125:
                r = t[0]
                print(x_with_possibility_sol[t[0]])
            else:
                r = random.choice(unknown)
            lclick(*r)
            time.sleep(click_t)   
            if debug:
                game_img = ImageGrab.grab(rect)
                draw = ImageDraw.Draw(game_img)
                draw.rectangle((r[1]*block_width, r[0]*block_height,
                              (r[1]+1)*block_width, (r[0]+1)*block_height),
                              None, 'yellow', 2)
                game_img.save(f'{debug_folder}\\{outer_loop}\\random.jpg')

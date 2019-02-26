import win32gui
import win32api
#import win32com.client
import win32con
from PIL import ImageGrab
import time
import sys
import random
import itertools

# 窗口坐标
left = 0
top = 0
right = 0
bottom = 0

lclick_t = 0.1

block_width, block_height = 16, 16
blocks_x_num, blocks_y_num = 30, 16
mines_num = 99

# 数字1-8 周围雷数
# un 未被打开
# 0 被打开 空白
# hongqi 红旗
# boom 普通雷
# boom_red 踩中的雷
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

img_block = {img_0: 0, img_1: 1, img_2: 2, img_3: 3, img_4: 4, img_5: 5, img_6: 6,
       img_7: 7, img_8: 8, img_unknown: -1, img_flag: -2, img_boom: -3, img_boom_red: -4}

mine_map = dict()
for i in range(blocks_y_num):
    for j in range(blocks_x_num):
        mine_map[(i, j)] = None
    
# 扫描雷区图像
def show_map():
    num_img = []
    unsure_img = []
    game_img = ImageGrab.grab().crop(rect)
    #game_img.save("bug.png")
    for i in range(blocks_y_num):
        for j in range(blocks_x_num):
            this_image = tuple(game_img.crop((j*block_width, i*block_height,
                                              (j+1)*block_width, (i+1)*block_height)).getcolors())
            if this_image not in img_block:
                print("无法识别图像")
                print("坐标: ", (i, j))
                sys.exit(1)
            mine_map[(i, j)] = img_block[this_image]
            if mine_map[(i, j)]==-3 or mine_map[(i, j)]==-4:
                print("game over")
                sys.exit(0)
            if mine_map[(i, j)] == -1:
                unsure_img.append((i, j))
            elif 1<=mine_map[(i, j)]<=8 and mine_not_trial_map[(i, j)]==0:
                num_img.append((i, j))
    return [num_img, unsure_img]
    

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

def gauss():
    global label_with_x
    global mat_trial_xPlus1
    global fix
    global free
    
    row_index = 0
    col_index = 0

    while col_index != len(mat_trial_xPlus1[0])-1:
        row_index_tmp = row_index
        while row_index_tmp!=len(mat_trial_xPlus1) and mat_trial_xPlus1[row_index_tmp][col_index]==0:
            row_index_tmp += 1
        if row_index_tmp == len(mat_trial_xPlus1):
            free.append(col_index)
            col_index += 1
        else:
            mat_trial_xPlus1[row_index_tmp], mat_trial_xPlus1[row_index] = mat_trial_xPlus1[row_index], mat_trial_xPlus1[row_index_tmp]
            if mat_trial_xPlus1[row_index][col_index] != 1:
                c = mat_trial_xPlus1[row_index][col_index]
                for j in range(col_index, len(mat_trial_xPlus1[0])):
                    mat_trial_xPlus1[row_index][j] //= c
            
            for i in range(len(mat_trial_xPlus1)):
                c = mat_trial_xPlus1[i][col_index]  
                if mat_trial_xPlus1[i][col_index]!=0 and i!=row_index:
                    for j in range(col_index, len(mat_trial_xPlus1[0])):
                        mat_trial_xPlus1[i][j] -= mat_trial_xPlus1[row_index][j] * c
                                    
            fix.append(col_index)
            row_index += 1
            col_index += 1

    # row_index == len(fix)
    for i in range(len(mat_trial_xPlus1) - row_index):
        mat_trial_xPlus1.pop()
            
    # 交换列，同时改变label_with_x
    for i in range(len(fix)):
        label_with_x[fix[i]], label_with_x[i] = label_with_x[i], label_with_x[fix[i]]
        for ii in range(0, i+1):
            mat_trial_xPlus1[ii][fix[i]], mat_trial_xPlus1[ii][i] = mat_trial_xPlus1[ii][i], mat_trial_xPlus1[ii][fix[i]]
    # 更新fix 和 free
    for i in range(len(fix)):
        fix[i] = i
    for i in range(len(free)):
        free[i] = i + len(fix)
        

def logic():
    global win_prob
    
    def bfs(row_trial):
        rows = []
        cols = []
        connected = []
        search_queue = [row_trial]
        searched = []
        while search_queue:
            v = search_queue.pop(0)
            if v not in searched:
                connected.append(v)
                search_queue.extend(rows_cols_graph[v])
                searched.append(v)
        for i in connected:
            if i < len(fix):
                rows.append(i)
            else:
                cols.append(i)
        return [rows, cols]
    
    # 构建fix，free关系图
    rows_cols_graph = dict()

    for i in range(len(x_to_solve)):
        rows_cols_graph[i] = []
    for i in fix:
        for j in free:
            if mat_trial_xPlus1[i][j] != 0:
                rows_cols_graph[i].append(j)
                rows_cols_graph[j].append(i)
                
    connected_parts = []
    rows = fix[:]
    
    while len(rows) != 0:
        connected_parts.append(bfs(rows[0]))
        for r in connected_parts[-1][0]:
            rows.remove(r)
    
    x_with_sol = dict()
    x_with_possibility_sol = dict()
    mines_mean = 0
    
    for part in connected_parts:
        if len(part[1]) == 0:
            x_with_sol[label_with_x[part[0][0]]] = mat_trial_xPlus1[part[0][0]][-1]
        else:
            s = 0
            total = itertools.product(* [[0, 1]]*len(part[1]))
            all_possible = []
            for t in total:
                for i in range(len(part[0])): # 行标
                    tt = 0
                    for j in range(len(part[1])):
                        tt += mat_trial_xPlus1[part[0][i]][part[1][j]] * t[j]
                    if mat_trial_xPlus1[part[0][i]][-1] - tt not in [0, 1]:
                        break
                else:
                    all_possible.append(list(t))
                    for i in range(len(part[0])):
                        tt = 0
                        for j in range(len(part[1])):
                            tt += mat_trial_xPlus1[part[0][i]][part[1][j]] * t[j]
                        all_possible[-1].append(mat_trial_xPlus1[part[0][i]][-1] - tt)
                    s += sum(all_possible[-1])

            if len(all_possible) == 0:
                print("what the fuck")
                sys.exit()
            mines_mean += s/len(all_possible)

            # part[1], part[0]
            for j in range(len(part[1])):
                count_1 = 0
                for i in range(len(all_possible)):
                    if all_possible[i][j] == 1:
                        count_1 += 1
                if count_1 == len(all_possible):
                    x_with_sol[label_with_x[part[1][j]]] = 1
                elif count_1 == 0:
                    x_with_sol[label_with_x[part[1][j]]] = 0
                else:
                    x_with_possibility_sol[label_with_x[part[1][j]]] = count_1/len(all_possible)

            for j in range(len(part[1]), len(all_possible[0])):
                count_1 = 0
                for i in range(len(all_possible)):
                    if all_possible[i][j] == 1:
                        count_1 += 1
                if count_1 == len(all_possible):
                    x_with_sol[label_with_x[part[0][j-len(part[1])]]] = 1
                elif count_1 == 0:
                    x_with_sol[label_with_x[part[0][j-len(part[1])]]] = 0
                else:
                    x_with_possibility_sol[label_with_x[part[0][j-len(part[1])]]] = count_1/len(all_possible)

    '''
    if len(x_with_sol) == 0:
        # 为了完美，结合len(x_to_solve)<=left_mines的条件重新算一遍
        if len(x_to_solve) > left_mines:
            print("good luck: ", len(free))
            x_with_possibility_sol = {}
            mines_mean = 0
            s = 0
            total = itertools.product(* [[0, 1]]*len(free))
            all_possible = []
            for t in total:
                for i in range(len(fix)): 
                    tt = 0
                    for j in range(len(free)):
                        tt += mat_trial_xPlus1[fix[i]][free[j]] * t[j]
                    if mat_trial_xPlus1[fix[i]][-1] - tt not in [0, 1]:
                        break
                else:
                    all_possible.append(list(t))
                    for i in range(len(fix)):
                        tt = 0
                        for j in range(len(free)):
                            tt += mat_trial_xPlus1[fix[i]][free[j]] * t[j]
                        all_possible[-1].append(mat_trial_xPlus1[fix[i]][-1] - tt)
                    if sum(all_possible[-1]) > left_mines:
                        all_possible.pop()
                    else:
                        s += sum(all_possible[-1])

            if len(all_possible) == 0:
                print("fuck shit")
                sys.exit()
            mines_mean += s/len(all_possible)

            # free, fix 
            for j in range(len(free)):
                count_1 = 0
                for i in range(len(all_possible)):
                    if all_possible[i][j] == 1:
                        count_1 += 1
                x_with_possibility_sol[label_with_x[free[j]]] = count_1/len(all_possible)

            for j in range(len(fix), len(all_possible[0])):
                count_1 = 0
                for i in range(len(all_possible)):
                    if all_possible[i][j] == 1:
                        count_1 += 1
                x_with_possibility_sol[label_with_x[fix[j-len(free)]]] = count_1/len(all_possible)
    '''
    if len(x_with_sol) == 0:
        t = sorted(x_with_possibility_sol, key = lambda x: x_with_possibility_sol[x])
        out_unsure = unsure[:]
        for x in x_to_solve:
            out_unsure.remove(x)
            
        if len(x_with_possibility_sol)==0:
            print("fuck!")
            sys.exit()
            
        if len(out_unsure)!=0 and x_with_possibility_sol[t[0]]>(left_mines-mines_mean)/len(out_unsure):
            print("(left_mines-mines_mean)/len(out_unsure)", (left_mines-mines_mean)/len(out_unsure))
            win_prob *= 1 - (left_mines-mines_mean)/len(out_unsure)
            return {random.choice(out_unsure): 0}
        
        print("x_with_possibility_sol[t[0]]", x_with_possibility_sol[t[0]])
        win_prob *= 1 - x_with_possibility_sol[t[0]]
        return {t[0]: 0}
    
    return x_with_sol

class_name = "TMain"
title_name = "Minesweeper Arbiter"
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

mine_not_trial_map = dict()
# 1: not trial, 0: trial
for i in range(blocks_y_num):
    for j in range(blocks_x_num):
        mine_not_trial_map[(i, j)] = 0

lclick(0, 0)
time.sleep(lclick_t)
[trial, unsure] = show_map()
win_prob = 1
left_mines = mines_num
    
while 1:
    if len(unsure) == 0:
        print("win", win_prob)
        sys.exit(0)
        
    if len(trial) == 0:
        print("left_mines/len(unsure)", left_mines/len(unsure))
        win_prob *= 1 - left_mines/len(unsure)
        r = random.choice(unsure)
        lclick(*r)
        time.sleep(lclick_t)
        [trial, unsure] = show_map()
            
    no_choice_for_naive_strategy = 1
    x_blocks = []
    x_to_solve = []
    for t in trial[:]: # 总是忘了加[:], 大坑
        unknown = []  
        flag = 0  
        blocks_8_around = get_around(*t)
        for e in blocks_8_around:
            if mine_map[e] == -1:  
                unknown.append(e)
            elif mine_map[e] == -2:  
                flag += 1
        
        if len(unknown) == 0:
            mine_not_trial_map[t] = 1
            trial.remove(t)
        else:
            if flag == mine_map[t]: # 周围安全
                no_choice_for_naive_strategy = 0
                mine_not_trial_map[t] = 1
                #trial.remove(t) # 可有可无，反正show_map里已经更新了
                for e in unknown:
                    lclick(*e)
                    time.sleep(lclick_t)
                    [trial, unsure] = show_map()
            elif len(unknown) == mine_map[t] - flag: # 说明周围全是雷，右键点击所有格
                no_choice_for_naive_strategy = 0
                mine_not_trial_map[t] = 1
                trial.remove(t)
                for e in unknown:
                    rclick(*e)
                    mine_map[e] = -2
                    left_mines -= 1
                    unsure.remove(e)
            else:
                if no_choice_for_naive_strategy:
                    for _ in unknown:
                        if _ not in x_to_solve:
                            x_to_solve.append(_)
                    x_blocks.append([unknown, mine_map[t]-flag])
            
    if no_choice_for_naive_strategy:  
        mat_trial_xPlus1 = [] # size is len(trial) * len(x)+1
        if len(x_to_solve) == 0:
            continue

        # 构造矩阵
        for _ in x_blocks:
            mat_trial_xPlus1.append([])
            for xx in x_to_solve:
                if xx in _[0]:# unknown
                    mat_trial_xPlus1[-1].append(1)
                else:
                    mat_trial_xPlus1[-1].append(0)
            mat_trial_xPlus1[-1].append(_[1])
            
        # 特殊情况增加一个约束条件
        if len(x_to_solve) == len(unsure):
            mat_trial_xPlus1.append([1]*len(x_to_solve))
            mat_trial_xPlus1[-1].append(left_mines)

        # 联系变量下标(i)与实际坐标(x_to_solve[i])
        label_with_x = dict()
        for i in range(len(x_to_solve)):
            label_with_x[i] = x_to_solve[i]

        # 固定变元与自由变元
        fix = []
        free = []
        '''
        for _ in mat_trial_xPlus1:
            print(_)
        print()
        '''
        gauss()
        '''
        for _ in mat_trial_xPlus1:
            print(_)
        '''
        sol = logic()
        
        for k in sol:
            if sol[k] == 0:
                if mine_map[k] == -1:
                    lclick(*k)
                    time.sleep(lclick_t)
                    [trial, unsure] = show_map()
            elif sol[k] == 1:
                if mine_map[k] == -1:
                    rclick(*k)
                    mine_map[k] = -2
                    left_mines -= 1
                    unsure.remove(k)

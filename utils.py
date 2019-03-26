import itertools

def gauss(mat_to_solve, x_to_solve):
    fix, free = [], []
    row_index = 0
    col_index = 0

    while col_index != len(mat_to_solve[0])-1:
        row_index_tmp = row_index
        while row_index_tmp!=len(mat_to_solve) and mat_to_solve[row_index_tmp][col_index]==0:
            row_index_tmp += 1
        if row_index_tmp == len(mat_to_solve):
            free.append(col_index)
            col_index += 1
        else:
            mat_to_solve[row_index_tmp], mat_to_solve[row_index] = mat_to_solve[row_index], mat_to_solve[row_index_tmp]
            if mat_to_solve[row_index][col_index] != 1:
                c = mat_to_solve[row_index][col_index]
                for j in range(col_index, len(mat_to_solve[0])):
                    mat_to_solve[row_index][j] //= c
            
            for i in range(len(mat_to_solve)):
                c = mat_to_solve[i][col_index]
                if c!=0 and i!=row_index:
                    for j in range(col_index, len(mat_to_solve[0])):
                        mat_to_solve[i][j] -= mat_to_solve[row_index][j] * c
                                    
            fix.append(col_index)
            row_index += 1
            col_index += 1
            
    for i in range(len(mat_to_solve) - row_index):
        mat_to_solve.pop()
    for i in range(len(fix)):
        x_to_solve[fix[i]], x_to_solve[i] = x_to_solve[i], x_to_solve[fix[i]]
        for ii in range(0, i+1):
            mat_to_solve[ii][fix[i]], mat_to_solve[ii][i] = mat_to_solve[ii][i], mat_to_solve[ii][fix[i]]

    for i in range(len(fix)):
        fix[i] = i
    for i in range(len(free)):
        free[i] = i + len(fix)
        
    return mat_to_solve, x_to_solve, fix, free

def create_graph(mat_to_solve, fix, free):
    rows_cols_graph = {}
    for i in range(len(mat_to_solve[0])-1):
        rows_cols_graph[i] = []
    for i in fix:
        for j in free:
            if mat_to_solve[i][j] != 0:
                rows_cols_graph[i].append(j)
                rows_cols_graph[j].append(i)
    return rows_cols_graph

def bfs(row_trial, rows_cols_graph, fix):
    rows = []
    cols = []
    connected = []
    search_queue = [row_trial]
    searched = set()
    while search_queue:
        v = search_queue.pop(0)
        if v not in searched:
            connected.append(v)
            search_queue.extend(rows_cols_graph[v])
            searched.add(v)
    for i in connected:
        if i < len(fix):
            rows.append(i)
        else:
            cols.append(i)
    return rows, cols

def get_connected_parts(mat_to_solve, fix, free):
    rows_cols_graph = create_graph(mat_to_solve, fix, free)
    connected_parts = []
    rows = fix[:]
    
    while len(rows) != 0:
        connected_parts.append(bfs(rows[0], rows_cols_graph, fix))
        for r in connected_parts[-1][0]:
            rows.remove(r)
            
    return connected_parts

def get_x_sol(mat_to_solve, x_to_solve, connected_parts, left_mines):
    x_sol, x_possibility_sol = {}, {}
    
    for part in connected_parts:
        if len(part[1]) == 0:
            x_sol[x_to_solve[part[0][0]]] = mat_to_solve[part[0][0]][-1]
        else:
            total = itertools.product(*[[0, 1]]*len(part[1]))
            all_possible = []
            for t in total:
                tmp = []
                for i in range(len(part[0])): 
                    tt = 0
                    for j in range(len(part[1])):
                        tt += mat_to_solve[part[0][i]][part[1][j]] * t[j]
                    if mat_to_solve[part[0][i]][-1] - tt not in [0, 1]:
                        break
                    tmp.append(mat_to_solve[part[0][i]][-1] - tt)
                else:
                    tmp.extend(list(t))
                    if sum(tmp) > left_mines:
                        break
                    all_possible.append(tmp)

            p = part[0] + part[1]
            for j in range(len(p)):
                count_1 = 0
                for i in range(len(all_possible)):
                    if all_possible[i][j] == 1:
                        count_1 += 1
                if count_1 == len(all_possible):
                    x_sol[x_to_solve[p[j]]] = 1
                elif count_1 == 0:
                    x_sol[x_to_solve[p[j]]] = 0
                else:
                    x_possibility_sol[x_to_solve[p[j]]] = count_1/len(all_possible)

    return x_sol, x_possibility_sol

if __name__ == "__main__": 
    mat_to_solve = [[1,1,0,0,0,0,0,0,0,0,0,1],
                    [0,0,1,1,1,0,0,0,0,0,0,1],
                    [0,1,0,0,0,1,1,0,0,0,0,1],
                    [0,0,0,0,1,0,0,1,0,0,0,1],
                    [0,0,0,0,1,0,0,1,1,1,1,2],
                    [0,0,0,0,0,0,0,1,0,0,1,1],
                    [0,0,0,0,0,0,1,0,0,0,1,1]]
    x_to_solve = [(0,12),(1,12),(0,9),(1,9),(2,9),(2,12),(3,12),(3,9),(2,10),(2,11),(3,11)]
    left_mines = 6

    mat_to_solve, x_to_solve, fix, free = gauss(mat_to_solve, x_to_solve)                    
    connected_parts = get_connected_parts(mat_to_solve, fix, free)
    x_sol, x_possibility_sol = get_x_sol(\
                mat_to_solve, x_to_solve, connected_parts, left_mines)
    print(x_sol, '\n', x_possibility_sol)

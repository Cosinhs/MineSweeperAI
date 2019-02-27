# MineSweeperAI 
This is an AI for game MineSweeper, only works in Windows with ms_arbiter.exe downloaded from http://www.minesweeper.info/downloads/Arbiter.html   
  
MineSolverSimple.py implements a simple algorithm discussed in https://zhuanlan.zhihu.com/p/39669361, it's enough for easy and normal mode, but rather weak when dealing with the hard mode since it does not consider the logical relation.
In fact, we just need to solve a linear system of equations which may not has a unique solve. Furthermore, we can calculate the probability and choose a better position to click. 
It is implemented in MineSolver.py.
  
# Usage
Unpack Arbiter__0.52.3.zip, run ms_arbiter.exe, select a difficulty according to the code setting, then run the code and enjoy it!

# Demo
![image](https://github.com/Cosinhs/MineSweeperAI/blob/master/Video_2019-02-27_001928.gif)

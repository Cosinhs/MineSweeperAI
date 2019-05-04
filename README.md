# MineSweeperAI 
This is a bot for game MineSweeper, only works in Windows OS with ms_arbiter.exe downloaded from http://www.minesweeper.info/downloads/Arbiter.html.   
  
Here I improve the algorithm discussed in https://zhuanlan.zhihu.com/p/39669361. The original algorithm works well in easy (or beginner, i.e., 8×8, 10 mines) and normal (or intermediate, i.e., 16×16, 40 mines) difficulty modes, but too weak to solve the hard (or expert, i.e., 30×16, 99 mines) difficulty mode since it only uses naive strategy.
In fact, we just need to solve an indeterminate linear system of equations to find all the safe positions to left click and mine positions to right click. If none is found, we can calculate the probability of every possible positions further and choose the position with least mine probability to left click. 
  
# Usage
Unpack Arbiter__0.52.3.zip, start ms_arbiter.exe, select a difficulty according to the code, then run the script and enjoy.

# Demo  
![image](https://github.com/Cosinhs/MineSweeperAI/blob/master/demo_play/easy.gif)  
![image](https://github.com/Cosinhs/MineSweeperAI/blob/master/demo_play/normal.gif)  
![image](https://github.com/Cosinhs/MineSweeperAI/blob/master/demo_play/hard.gif)  

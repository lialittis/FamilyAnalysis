import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import defaultdict

import matplotlib.pyplot as plt
import matplotlib.image as mpimg
#from scipy.misc import imread
from matplotlib.animation import FuncAnimation, PillowWriter
import numpy as np

"""predefine"""

map_img = plt.imread("./floorplan02.png")
max_x = 4.55
min_x = -3.6
max_y = 11.4
min_y = -1.8

"""读取excel文件"""

input_path = "./目标/UncleOfSun_choice_data_all.xlsx"

df = pd.read_excel(open(input_path,'rb'))
df = df.iloc[: , 1:]

list_x = list(df["x"])
list_y = list(df["y"])

length = len(list_x)
length_each_gif = 100
number_of_gif = 0

fig,ax = plt.subplots(figsize=(max_x-min_x,max_y-min_y))

def animate(i):
    ax.clear()
    ax.imshow(map_img,extent=[min_x-0.1,max_x,min_y-0.1,max_y])
    ax.set_xlim(min_x,max_x)
    ax.set_ylim(min_y,max_y)
    init_index = (number_of_gif*length_each_gif)
    index = i + (number_of_gif*length_each_gif)
    line, = ax.plot(list_x[init_index:index], list_y[init_index:index], color = 'blue', lw=1)
    point, = ax.plot(list_x[index], list_y[index], marker='.', color='blue')
    return line, point

while(number_of_gif*length_each_gif<length):
    print(number_of_gif)
    number_of_gif += 1
    ani = FuncAnimation(fig, animate, interval=30, blit=True, repeat=True, frames=min(length_each_gif,length-number_of_gif*length_each_gif))
    ani.save("motion_"+str(number_of_gif)+".gif", dpi=300, writer=PillowWriter(fps=5))
#ani.save('motion.mp4', writer = 'ffmpeg', fps=30)

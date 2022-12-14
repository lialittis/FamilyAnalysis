import sys
colab_datadir = '/content/drive/MyDrive/sharedCodeBase/FamilyAnalysis/'
local_datadir = './'
datadir = local_datadir
sys.path.append(datadir)

import setConfigs as cf
from setConfigs import families
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

def saveDFtoWB(df_n,filename):
    # try to store by workbook
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df_n, index=False, header=True):
        ws.append(r)
    wb.save(filename)

# A special queue, which will record several biggest values
# This method has one bug : 
# it is a monotonic queue, which means once we get a big value,
# the rest will not be considered
# So I decide to just use internal sort function to replace it
class MQueue:
    def __init__(self,maxsize):
        self.queue = []
        self.maxsize = maxsize
	
    def push(self, value):
        # print("should add : ",value)
        self.queue.append(value)
        if len(self.queue)==self.maxsize : 
            self.queue.pop(-1)
	
    def pop(self):
        if self.queue:
            return self.queue.pop(0)

    def average(self):
        res = 0
        if len(self.queue) != 0:
            res = sum(self.queue)/len(self.queue)
        return  res

    def toString(self):
        string = ""
        for i in self.queue:
            string += str(i) + ","
        print("MQueue is : ",string)

def getKLargest(queue,k):
    queue.sort()
    d = k//2
    # delete the first k/2 largest ones, to avoid huge bias
    if len(queue) > k+d : 
        return queue[-k-d:-d]
    else:
        return queue

def getSlope(outLoc,points,zone):
    slope = 0
    all_slopes = []
    '''
    queue_max_slopes = MQueue(maxsize=20)
    for x,y in points:
        temp = (x-zone.ux)/(y-zone.ly) if outLoc == "right" else (zone.lx-x)/(y-zone.ly) 
        if temp > slope:
            slope = temp
            queue_max_slopes.push(slope)
    
    queue_max_slopes.toString()
    '''
    for x,y in points:
        temp = (x-zone.ux)/(y-zone.ly) if outLoc == "right" else (zone.lx-x)/(y-zone.ly)
        all_slopes.append(temp)

    # choose first k largest slope
    # k is the largest one between 20 and number of 0.1*points
    k = len(all_slopes) // 10
    #print("k is :",k)
    queue_max_slopes = getKLargest(all_slopes,k)
     
    average_slope = sum(queue_max_slopes)/len(queue_max_slopes) if len(queue_max_slopes) > 0 else 0

    #print("The adjusted slope is :", average_slope)
    return average_slope

def delete(x,y,zone):
    # print(x,y)
    x = zone.outdoor[0]
    y = zone.outdoor[1]
    # print(x,y)
    return x,y

def MoveToDown(x,y,zone):
    # print(x,y)
    y = zone.ly
    # print(x,y)
    return x,y

def MoveToUp(x,y,zone):
    # print(x,y)
    y = zone.uy
    # print(x,y)
    return x,y

def MoveToLeft(x,y,zone):
    # print(x,y)
    x = zone.lx
    # print(x,y)
    return x,y

def MoveToRight(x,y,zone):
    # print(x,y)
    x = zone.ux
    # print(x,y)
    return x,y

def MoveToLeftDown(x,y,zone):
    # print(x,y)
    x = zone.lx
    y = zone.ly
    # print(x,y)
    return x,y

def MoveToLeftUp(x,y,zone):
    # print(x,y)
    x = zone.lx
    y = zone.uy
    # print(x,y)
    return x,y

def MoveToRightDown(x,y,zone):
    # print(x,y)
    x = zone.ux
    y = zone.ly
    # print(x,y)
    return x,y

def MoveToRightUp(x,y,zone):
    # print(x,y)
    x = zone.ux
    y = zone.uy
    # print(x,y)
    return x,y

def MoveToLeftLinearly(points,zone):
    slope = getSlope("right",points,zone)
    for i in range(len(points)):
        x = points[i][0]
        y = points[i][1]
        # print("old:",x,y)
        nx = x - (y - zone.ly)*slope
        # print("slope",slope)
        if nx > zone.ux:
            points[i][0] = zone.ux
        else:
            points[i][0] = max(zone.lx,nx)
        # print("new:",points[i][0],points[i][1])
    return points

def MoveToRightLinearly(points,zone):
    slope = getSlope("left",points,zone)
    for i in range(len(points)):
        x = points[i][0]
        y = points[i][1]
        # print("old:",x,y)
        nx = x + (y - zone.ly)*slope
        # print("slope",slope)
        if nx < zone.lx:
            points[i][0] = zone.lx
        else:
            points[i][0] = min(zone.ux,nx)
        # print("new:",points[i][0],points[i][1])
    return points

def inZone(x,y,zone):
    # special scenario
    #if zone.place =="InKitchen":
    if zone.method == "MoveToLeftLinearly":
        if x<=zone.lx or y <= zone.ly or y >= zone.uy:
            return False
    elif zone.method == "MoveToRightLinearly":
        if x>=zone.ux or y <= zone.ly or y >= zone.uy:
            return False
    else:
        if x <= zone.lx or x >= zone.ux or y <= zone.ly or y >= zone.uy:
            return False
    #else:
    #    if x <= zone.lx or x >= zone.ux or y <= zone.ly or y >= zone.uy:
    #        return False
    return True

methods ={
    "Delete" :delete,
    "MoveToDown":MoveToDown,
    "MoveToUp":MoveToUp,
    "MoveToLeft":MoveToLeft,
    "MoveToRight":MoveToRight,
    "MoveToLeftDown":MoveToLeftDown,
    "MoveToLeftUp":MoveToLeftUp,
    "MoveToRightDown":MoveToRightDown,
    "MoveToRightUp":MoveToRightUp,
    "MoveToLeftLinearly":MoveToLeftLinearly,
    "MoveToRightLinearly":MoveToRightLinearly
}

def chooseFamily():
    print("==>Familiy choices :<==")
    for i in range(len(families)):
        print(i+1,":"+families[i].name)
    choice = input()
    index = int(choice)-1
    print("Family of "+ families[index].name+" is chosen!")
    cf.traceFamily([families[index]])
    return families[index]

if __name__ == "__main__":
    # cf.traceFamily(families)
    chooseFamily()
    f = methods["MoveToUp"]
    x,y = f(1,2,1)
    # print(x,y)


import xlrd,xlwt
import os
import matplotlib.pyplot as plt
from matplotlib import cm
import numpy as np
from mpl_toolkits.mplot3d import Axes3D

class ReadExcel():
    def __init__(self,path):
        self.path=path
    def readsheet1(self):
        book=xlrd.open_workbook(self.path)
        sheet=book.sheet_by_index(0)
        return sheet
    def getrow(self):
        row=self.readsheet1().cell_value(1,1)
        return row
    def getdata(self,minrow):
        dataset=[]
        for row in range(14,minrow+14):
            data=self.readsheet1().cell_value(row,1)
            dataset.append(data)
        return dataset
    def getrpm(self,minrow):
        rpmset=[]
        for row in range(14,minrow+14):
            rpm=self.readsheet1().cell_value(row,0)
            rpmset.append(rpm)
        return rpmset
    def getmark(self):
        markset=self.readsheet1().cell_value(4,1).split('_')
        return markset
    def getfiledirec(self):
        if 'CCW' in self.getmark():
            direction='CCW'
        elif 'CW' in self.getmark():
            direction='CW'
        else:
            print('文件与方向相关的命名错误！')
        return direction
    def getfileload(self,loads=['{num:.0%}'.format(num=load/100) for load in range(10,101,10)]):
        for load in loads:
            if load in self.getmark():
                return load

    def getfilevol(self):
        for vol in ['11V','12V','13V']:
            if vol in self.getmark():
                return vol

def valueset(voltage):
    CCW_markset=[]
    CW_markset=[]
    for rooz, dir, files in os.walk(root):
        for file in files:
            path = root + '\\' + file
            excel = ReadExcel(path)
            if excel.getfiledirec() == 'CCW' and excel.getfilevol() == voltage:
                CCW_markset.append([excel.getfileload(), excel.getdata(minrow)])
            if excel.getfiledirec() == 'CW' and excel.getfilevol() == vlotage:
                CW_markset.append([excel.getfileload(), excel.getdata(minrow)])
    dic_CCW_markset = dict(CCW_markset)
    dic_CW_markset = dict(CW_markset)
    markset = []
    for load in loadset:
        dic_CCW_markset[load].reverse()
        mark = dic_CCW_markset[load] + dic_CW_markset[load]
        markset.append(mark)
    return markset

root=input('Please input the root path: ')
# root='D:\GEM_CNB_Output_excel'
rowset=[]
loadset=['{num:.0%}'.format(num=load/100) for load in range(10,101,10)]

global files
for rooz,dir,files in os.walk(root):
    for file in files:
        path=root+'\\'+file
        excel=ReadExcel(path)
        rowset.append(excel.getrow())
minrow=int(min(rowset))
porpm=ReadExcel(root+'\\'+files[0]).getrpm(minrow)
nerpm=[-rpm for rpm in porpm]
nerpm.reverse()
rpmset=nerpm+porpm

CCW_12V_markset=[]
CW_12V_markset=[]
CCW_11V_markset=[]
CW_11V_markset=[]
CCW_13V_markset=[]
CW_13V_markset=[]

for file in files:
    path = root + '\\' + file
    excel = ReadExcel(path)
    filedirec=excel.getfiledirec()
    filevol=excel.getfilevol()
    fileload=excel.getfileload()
    filedata=excel.getdata(minrow)
    if filedirec=='CCW' and filevol=='12V':
        CCW_12V_markset.append([fileload, filedata])
    elif filedirec == 'CW' and filevol == '12V':
        CW_12V_markset.append([fileload, filedata])
    elif filedirec=='CCW' and filevol=='11V':
        CCW_11V_markset.append([fileload,filedata])
    elif filedirec == 'CW' and filevol == '11V':
        CW_11V_markset.append([fileload, filedata])
    elif filedirec=='CCW' and filevol=='13V':
        CCW_13V_markset.append([fileload,filedata])
    elif filedirec == 'CW' and filevol == '13V':
        CW_13V_markset.append([fileload, filedata])
    else:
        print('error')
dic_CCW_12v_markset=dict(CCW_12V_markset)
dic_CW_12v_markset=dict(CW_12V_markset)
dic_CCW_11v_markset=dict(CCW_11V_markset)
dic_CW_11v_markset=dict(CW_11V_markset)
dic_CCW_13v_markset=dict(CCW_13V_markset)
dic_CW_13v_markset=dict(CW_13V_markset)
markset_12V=[]
markset_11V=[]
markset_13V=[]
for load in loadset:
    a=dic_CCW_12v_markset[load]
    b=dic_CCW_11v_markset[load]
    c=dic_CCW_13v_markset[load]
    a.reverse()
    b.reverse()
    c.reverse()
    mark_12V=a+dic_CW_12v_markset[load]
    mark_11V=b+dic_CW_11v_markset[load]
    mark_13V = c + dic_CW_13v_markset[load]
    markset_12V.append(mark_12V)
    markset_11V.append(mark_11V)
    markset_13V.append(mark_13V)
loadpercent=[l for l in range(10,101,10)]

norm=cm.colors.Normalize(vmin=0,vmax=0.02)
fig = plt.figure(figsize=(12,6))
ax = fig.add_subplot(3, 2, 1, projection='3d')
X=rpmset
Y=loadpercent
X,Y=np.meshgrid(X,Y)
Z=np.array(markset_11V)
surf=ax.plot_surface(X, Y, Z, rstride=1, cstride=1, cmap=plt.get_cmap('rainbow'),linewidth=2,antialiased=True,norm=norm)
font2 = {'family' : 'Times New Roman','weight' : 'normal','size'  : 10}
ax.set_xlabel('Rotational Speed /rpm ',font2,labelpad=10)
ax.set_ylabel('load /%',font2,labelpad=10)
ax.set_zlabel('Torque 24th Order /Nm',font2,labelpad=2)
ax.set_title('Torque 24th Order Result with 11V',font2)
labels=ax.get_xticklabels()+ax.get_yticklabels()+ax.get_zticklabels()
[label.set_fontname('Times New Roman') for label in labels]
ax.tick_params(labelsize=8)
ax.set_zlim3d(0, 0.08)
position=fig.add_axes([0.45,0.7,0.01,0.2])
cb=plt.colorbar(surf,cax=position,shrink=1,aspect=10)#pad 参数 wh_space = 2 * pad / (1 - pad)
for l in cb.ax.yaxis.get_ticklabels():
    l.set_family('Times New Roman')

ax=fig.add_subplot(3,2,2)
cont=ax.contourf(X,Y,Z,60,cmap=plt.get_cmap('rainbow'),vmin=0,vmax=0.02)
ax.set_xlabel('Rotational Speed /rpm ',font2,labelpad=10)
ax.set_ylabel('load /%',font2,labelpad=10)
ax.set_title('Torque 24th Order Result with 11V',font2)
labels=ax.get_xticklabels()+ax.get_yticklabels()
[label.set_fontname('Times New Roman') for label in labels]
position=fig.add_axes([0.93,0.7,0.01,0.2])
cb=plt.colorbar(cont,cax=position,shrink=1,aspect=10)#pad 参数 wh_space = 2 * pad / (1 - pad)
for l in cb.ax.yaxis.get_ticklabels():
    l.set_family('Times New Roman')

#==================#
#=====second=======#
#==================#
ax = fig.add_subplot(3, 2, 3, projection='3d')
X=rpmset
Y=loadpercent
X,Y=np.meshgrid(X,Y)
Z=np.array(markset_12V)
surf=ax.plot_surface(X, Y, Z, rstride=1, cstride=1, cmap=plt.get_cmap('rainbow'),linewidth=2,antialiased=True,norm=norm)
font2 = {'family' : 'Times New Roman','weight' : 'normal','size'  : 10}
ax.set_xlabel('Rotational Speed /rpm ',font2,labelpad=10)
ax.set_ylabel('load /%',font2,labelpad=10)
ax.set_zlabel('Torque 24th Order /Nm',font2,labelpad=2)
ax.set_title('Torque 24th Order Result with 12V',font2)
labels=ax.get_xticklabels()+ax.get_yticklabels()+ax.get_zticklabels()
[label.set_fontname('Times New Roman') for label in labels]
ax.tick_params(labelsize=8)
ax.set_zlim3d(0, 0.08)
position=fig.add_axes([0.45,0.4,0.01,0.2])
cb=plt.colorbar(surf,cax=position,shrink=1,aspect=10)#pad 参数 wh_space = 2 * pad / (1 - pad)
for l in cb.ax.yaxis.get_ticklabels():
    l.set_family('Times New Roman')

ax=fig.add_subplot(3,2,4)
cont=ax.contourf(X,Y,Z,60,cmap=plt.get_cmap('rainbow'),vmin=0,vmax=0.02)
ax.set_xlabel('Rotational Speed /rpm ',font2,labelpad=10)
ax.set_ylabel('load /%',font2,labelpad=10)
ax.set_title('Torque 24th Order Result with 12V',font2)
labels=ax.get_xticklabels()+ax.get_yticklabels()
[label.set_fontname('Times New Roman') for label in labels]
position=fig.add_axes([0.93,0.4,0.01,0.2])
cb=plt.colorbar(cont,cax=position,shrink=1,aspect=10)#pad 参数 wh_space = 2 * pad / (1 - pad)
for l in cb.ax.yaxis.get_ticklabels():
    l.set_family('Times New Roman')

ax = fig.add_subplot(3, 2, 5, projection='3d')
X=rpmset
Y=loadpercent
X,Y=np.meshgrid(X,Y)
Z=np.array(markset_13V)
surf=ax.plot_surface(X, Y, Z, rstride=1, cstride=1, cmap=plt.get_cmap('rainbow'),linewidth=2,antialiased=True,norm=norm)
font2 = {'family' : 'Times New Roman','weight' : 'normal','size'  : 10}
ax.set_xlabel('Rotational Speed /rpm ',font2,labelpad=10)
ax.set_ylabel('load /%',font2,labelpad=10)
ax.set_zlabel('Torque 24th Order /Nm',font2,labelpad=2)
ax.set_title('Torque 24th Order Result with 13V',font2)
labels=ax.get_xticklabels()+ax.get_yticklabels()+ax.get_zticklabels()
[label.set_fontname('Times New Roman') for label in labels]
ax.tick_params(labelsize=8)
ax.set_zlim3d(0, 0.08)
position=fig.add_axes([0.45,0.1,0.01,0.2])
cb=plt.colorbar(surf,cax=position,shrink=1,aspect=10)#pad 参数 wh_space = 2 * pad / (1 - pad)
for l in cb.ax.yaxis.get_ticklabels():
    l.set_family('Times New Roman')

ax=fig.add_subplot(3,2,6)
cont=ax.contourf(X,Y,Z,60,cmap=plt.get_cmap('rainbow'),vmin=0.0,vmax=0.02)
ax.set_xlabel('Rotational Speed /rpm ',font2,labelpad=10)
ax.set_ylabel('load /%',font2,labelpad=10)
ax.set_title('Torque 24th Order Result with 13V',font2)
labels=ax.get_xticklabels()+ax.get_yticklabels()
[label.set_fontname('Times New Roman') for label in labels]
position=fig.add_axes([0.93,0.1,0.01,0.2])
cb=plt.colorbar(cont,cax=position,shrink=1,aspect=10) #pad 参数 wh_space = 2 * pad / (1 - pad)
for l in cb.ax.yaxis.get_ticklabels():
    l.set_family('Times New Roman')

plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=0.8, hspace=0.8)
isExists=os.path.exists('D:\\test_result')
if not isExists:
    os.makedirs('D:\\test_result')
plt.savefig('D:\\test_result\\surface3d.svg')
plt.show()
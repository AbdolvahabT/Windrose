# -*- coding: utf-8 -*-
"""
Created on Sat Aug 22 17:59:58 2020

@author: AA
"""


import matplotlib.font_manager as font_manager

from tkinter import ttk, Menu

import tkinter

import tkinter as tk

from tkinter import filedialog as fd 

from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
from matplotlib import pyplot as plt

import numpy as np

from matplotlib.cbook import get_sample_data

import matplotlib as mpl
import matplotlib.gridspec as gridspec


#pip install xlrd==1.2.0
import xlrd 

import win32com.client



#for persian texts
# import arabic_reshaper

# install: pip install python-bidi
#from bidi.algorithm import get_display



import tkvalidate #pip install tkvalidate

import os


#global ws, wd, awd_sp, aws_sp, awd_su, aws_su, awd_fa, aws_fa, awd_wi, aws_wi, image_path

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False

def excel_reader(fileName, colDirection=4, colSpeed=5, rowStart=1, rowEnd=100, Max=40):
    
    

    # name_reverse = fileName[::-1]
    
    # index1 = name_reverse.find('\\')
    
    # name_reverse1 = name_reverse[index1+1:]
    
    # path = name_reverse1[::-1]


    Number = 10000
    
    
    output = os.getcwd() + '\\' + 'Data' + str(Number) + '.xlsx'
    
    while os.path.isfile(output):
       Number +=  1
       output = os.getcwd() + '\\' + 'Data' + str(Number) + '.xlsx'
       
    
    
    try :
        wb = xlrd.open_workbook(fileName) 
        
        
    except:
        
        wbc = o.Workbooks.Open(fileName)
        wbc.ActiveSheet.SaveAs(output,51)
        o.Workbooks.Close()
        
        wb = xlrd.open_workbook(output) 
        

    
    sheet = wb.sheet_by_index(0) 
    
    row = sheet.nrows
    
    if rowEnd =='-':
        rowEnd = row
        
    
    wd = np.zeros((row, 1))
    
    ws = np.zeros((row, 1))
    
    j = 0
    
    #for date: xlrd.xldate_as_tuple(sheet.cell(3,1).value,wb.datemode)
    
    for i in range(rowStart, rowEnd):
         vel = sheet.cell_value(i, colSpeed)
         dr = sheet.cell_value(i, colDirection)
         
        # if(isinstance(vel, float) or isinstance(vel, int)) and 
        
         # if (not( (int(dr) == dr) and (vel == 0 or vel == 0.0))) and ((int(dr) == dr) and (isinstance(vel, float) or isinstance(vel, int))):
         try:
             
             wd[j, 0] = dr
             
             ws[j, 0] = vel
             
             if vel <= Max: 
                 j += 1
         
         except:
             pass
    
    wd = wd[0:j, 0]
    
    ws = ws[0:j, 0]   
    
    if bool(CheckVarT0.get()):
        
        
        seasons_determiner(wb, colDirection, colSpeed, rowStart, rowEnd, Max_speed=Max)
    
    return(wd, ws, j)






button_fontsize = 20

global  user_dpi

def image(wd, ws ,figname ='fig1', nd=8, ns=5, standard=True, color_number=0,
          title_font_size=30, angle_fontsize=10, dirction_fontsize=15,
          legend_fontsize=10, title1="WindRose",
          x1=1, y1=1.1, user_dpi=100, percent_angle=60,
          user_opening=0.6, figwidth=6, figheight=8,
          rectx = 1, Xlim=10, Linewidth=1,legx=0.5, legy=0.5,
          plot_length=20, image_length=5,
          title_font_color='black', angle_font_color='black',
          dirction_font_color='black', legend_font_color='black'
          ):
    
    
    
    
    # print(aa, (bb, cc))
    
    
    
    sizeOfLegend = legend_fontsize
   

    # nd = 8
    
    # ns = 5
    
    maxs = max(ws)
    
    speed_mean = (maxs - .5)/(ns - 1)
    
    speed_span = [0]
    
    speed_span1 = [0.5 + i*speed_mean for i in range(ns)]
    
    speed_span.extend(speed_span1)
    
    speed_span[-1] = speed_span[-1] + 0.01
    
    
    # standard = True
    
    standard_speed_span = [0, 0.5, 2.1, 3.6, 5.7, 8.8, 11.1, 1000]
    
    st = 0
    
    Calm = 1
    
    if standard:
            speed_span = standard_speed_span
            ns = 7
            st = 1
    
    direction_span = 360/nd
    
    
    ds = np.zeros((nd, ns))
    
    #m = 0
    
    
    
    for i in range(1, nd):
        
        for j in range(ns):
            k = 0
            for m in range(len(ws)):
            
                if (wd[m]>= direction_span*((2*i-1)/2) and wd[m]< direction_span*((2*i+1)/2)) and (ws[m]>=speed_span[j] and ws[m]< speed_span[j+1]):
                    
                    k += 1
                    
    #            m += 1
                
            ds[i, j] = k
    
    
    for j in range(ns):
        k = 0
        
        for m in range(len(ws)):
        
            if (wd[m]>= direction_span*((2*i+1)/2) or wd[m]< direction_span*(1/2)) and (ws[m]>=speed_span[j] and ws[m]< speed_span[j+1]):
                
                k += 1
                
    #            m += 1
            
        ds[0, j] = k
    
    
    
    rds = np.zeros((nd, ns+1))
    
    for i in range(nd):
        
        for j in range(ns):
            
            rds[i, j] = 100*ds[i, j]/len(ws)
            
        rds[i, j+1] = np.sum(rds[i, 0:j+1])
    
    
    
    
    
    maxrds = max(rds[:, -1])
    
    
    maxrds_mean = np.ceil(maxrds/5)
    
    r = [i*maxrds_mean for i in range(1,6)]
    
    
    
    
    
    new_rds = rds[:, 0:-1]
        
    total_calm = np.sum(new_rds[:,0])
    
    
    constant = 1000
    
    rr = r.copy()
    
    rrr = r.copy()
    
    
    for i in range(len(rrr)):
        
        rrr[i] = constant * rr[i]/rr[0]
        
    
    r = rrr.copy()
    
    
    
    new_rrds = new_rds.copy()
    
    new_rrrds = new_rds.copy()
    
    
    for i in range((np.shape(new_rds)[0])):
        
        for j in range((np.shape(new_rds)[1])):
            
            new_rrrds[i, j] = constant * new_rds[i, j]/rr[0]
    
    
    new_rds = new_rrrds
    
    
    
    ####
    
    
    
    
    
    
    
    # lw = 1
    
    color_strings = [ 'red', 'orange', 'yellow', 'green', 'cyan', 'blue', \
                     'brown', 'indigo', 'magenta', 'purple', 'black', \
                         'aquamarine', 'navy','olive', 'teal', 'violet',\
                             'tan', 'gold', 'lime', 'slateblue']
    
    
    colors = color_strings[color_number:]
    
    colors.extend(color_strings[0:color_number])
    
    
    vR = varR.get()
    
    if vR < 3:
        aa = 0
    else:
        aa = 1
        
        plot_length, image_length = image_length, plot_length
        
    bb = 1 - aa   
    
    cc = 1- (vR % 2)
    
     
    
    
    fig3 = plt.figure(figsize=(figwidth, figheight))#constrained_layout=True)
    gs = fig3.add_gridspec(2, 2, height_ratios=[plot_length, image_length], 
                           width_ratios=[plot_length, image_length])
    ax = fig3.add_subplot(gs[aa, :])
    
    
    
    
    
    adad = Xlim
    
    ax.set(xlim=(-(r[-1]+adad), (r[-1]+adad)), ylim = (-(r[-1]+adad), (r[-1]+adad)))
    
    plt.axis('off')
    
    
    
    
    for i in range(4):
        
    #    plt.axis('off')
        
        draw_circle = plt.Circle((0, 0), r[i], fill=False, color='gray', linewidth=Linewidth, linestyle='-')
        
        ax.set_aspect(1)
        ax.add_artist(draw_circle)
    
    
    draw_circle = plt.Circle((0, 0), r[-1], fill=False, color='black', linewidth=Linewidth, linestyle='-')
        
    ax.set_aspect(1)
    ax.add_artist(draw_circle)
    
    
    
    
    pi = np.pi
    
    
    x = lambda r, phi : r * np.cos(phi)
    
    y = lambda r, phi : r * np.sin(phi)
    
    
    
    # percent_angle = 60
    
    percent_angle_rad = percent_angle * pi / 180
    
    
    for i in range(len(r)):
        
        ax.text(x(r[i], percent_angle_rad), y(r[i], percent_angle_rad), str(int(rr[i]))+'%', fontsize=angle_fontsize, color=angle_font_color)
    
    
    
    x_values = np.array([0, 0, r[-1], -r[-1], x(r[-1], pi/4), x(r[-1], 5*pi/4), x(r[-1], 3*pi/4), x(r[-1], 7*pi/4)])
    
    y_values = np.array([r[-1], -r[-1], 0, 0, y(r[-1], pi/4), y(r[-1], 5*pi/4), y(r[-1], 3*pi/4), y(r[-1], 7*pi/4)])
    
    
    for i in range(4):
        
        draw_line = plt.plot(x_values[2*i: 2*i+2], y_values[2*i: 2*i+2], color ='gray', linewidth=Linewidth)
        
        ax.add_artist(draw_line[0])
    
    
    
    x_texts = np.array([0, 0, r[-1]+r[0]/2, -(r[-1]+r[0]), x(r[-1]+r[0]/2, pi/4), x(r[-1]+1.5*r[0], 5*pi/4), x(r[-1]+r[0], 3*pi/4), x(r[-1]+r[0]/2, 7*pi/4)])

    y_texts = np.array([r[-1]+r[0]/2, -(r[-1]+2*r[0]/3), 0, 0, y(r[-1]+r[0]/2, pi/4), y(r[-1]+1.5*r[0], 5*pi/4), y(r[-1]+r[0], 3*pi/4), y(r[-1]+r[0]/2, 7*pi/4)])

    texts = ['N', 'S', 'E', 'W', 'N-E', 'S-W', 'N-W', 'S-E']
    
    
    for i in range(len(texts)):
        
        ax.text(x_texts[i], y_texts[i], texts[i], fontsize=dirction_fontsize, color=dirction_font_color)#, rotation=135)
    
    
    
    
    
    
    opening = user_opening+pi/40
    
    theta = pi/2 #2*pi/nd
    
    #beta = pi/40 
    
    alpha = opening * (2*pi/nd)/2
    
    
    
    polygon = mpl.patches.Polygon
    
    Polygon = mpl.patches.Polygon
    
    ply = np.array([[polygon]*(ns-0)]*nd)
    
    
    
    for i in range(nd):
        
        k = 0
        
        r0 = new_rds[i, st]
        
        pts = np.array([ [x(0, 0), y(0, 0)], [x(r0, theta-alpha), y(r0, theta-alpha)],  [x(r0, theta+alpha), y(r0, theta+alpha)]])
    
        ply[i, st] = Polygon(pts, closed=True, color=colors[k], linewidth=.1)
    
        ax.add_artist(ply[i, st])
        
    #    del p
        
        for j in range(st, ns-1):
            
            
            
            r1 = r0 + new_rds[i, j+1]
            
            k += 1
            
            pts = np.array([ [x(r0, theta-alpha), y(r0, theta-alpha)], [x(r0, theta+alpha), y(r0, theta+alpha)], [x(r1, theta+alpha), y(r1, theta+alpha)], [x(r1, theta-alpha), y(r1, theta-alpha)]])
    
            ply[i, j+1] = Polygon(pts, closed=True, color=colors[k], linewidth=.1)
    
            ax.add_artist(ply[i, j+1])
            
    #        del p
            
            r0 = r1
            
        theta -= 2*pi/nd
        
    
    
    Labels = []
    
    if st != 1:
        Labels = ['Calm']
        
        Calm = 0
    
    for i in range(1, len(speed_span) - 1 ):
        
        str1 = str(np.round(speed_span[i], 1)) + '--' + str(np.round(speed_span[i+1], 1))
        
        
        
        Labels.append(str1)
        
    if standard:
            
        Labels = Labels[:-1]
        
        Labels.append('>= 11.1')
        
    
    ply1 = ply[0, Calm:]
    
    
    
    # reshaped_text = arabic_reshaper.reshape(title1)    # correct its shape
    # bidi_text = get_display(reshaped_text) 
    
    plt.title(title1,   x=x1, y=y1, fontsize=title_font_size, fontname='Times New Roman',
              color=title_font_color);
    
    
    ##@@@@%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    
    
    try:
    
        im = plt.imread(get_sample_data(image_path))# + '\\Irimo.jpg'))
        ax2 = fig3.add_subplot(gs[bb, cc])
        
        # ax2 = plt.subplot(gs[1, 0])
        img = ax2.imshow(im)
        ax2.axis('off')
        
    except:
        pass
    
    
    
    font = font_manager.FontProperties(family='Times New Roman',
                                       weight='bold',
                                       style='normal', size=sizeOfLegend)
    
    lgd1 = plt.legend(ply1, Labels, bbox_to_anchor=(legx, legy),##(0,0) = right-bottom
               bbox_transform=plt.gcf().transFigure, prop=font)
    # lgd1.set_color("red")
    
    
    lgd1_title = 'Wind Speeds \n    (m/s)'
    
    if st == 1:
        
        lgd1_title += '\n\nCalms: ' + str(np.round(total_calm, 2)) + '%'
    
    lgd1.set_title(lgd1_title, prop={'size':legend_fontsize}) 
    
    # print((fig3).top, 344)
    plt.setp(lgd1.get_texts(), color=legend_font_color)
    lgd1._legend_title_box._text.set_color(legend_font_color)
    
    
    
    
    # plt.tight_layout(rect=[0, 0, rectx, recty])#, h_pad=-4)
    
    
    btm = float(Ent22.get())
    
    lft = float(Ent23.get())
    
    rgt = float(Ent24.get())
    
    tp = float(Ent25.get())
    
    wspc = float(Ent26.get())
    
    hspc = float(Ent27.get())
    
    if rgt == 0:
        rgt = None
        
    if tp == 0:
        tp = None
    
    plt.subplots_adjust(left=lft, bottom=btm, right=rgt, top=tp, wspace=wspc, hspace=hspc)
    
    L = ax.figure
    
     


    Number = 10000
     
    
    output = os.getcwd() + '\\' + figname + str(Number) + '.jpg'
    
    while os.path.isfile(output):
       Number +=  1
       output = os.getcwd() + '\\' + figname + str(Number) + '.jpg'
    

    # mpl.rcParams["figure.figsize"] = [5, 5]
    
    if bool(CheckVarT01.get()):
        L.savefig(output,   dpi=user_dpi)
    #, bbox_extra_artists=(lgd1,), bbox_inches='tight' , pad_inches=u_pad_inches,
    
    
    
    
    
    return(L, lgd1)
    

   




class SeaofBTCapp(tk.Tk):

    def __init__(self, f, lgd1, *args, **kwargs):
        
        tk.Tk.__init__(self, *args, **kwargs)

        tk.Tk.iconbitmap(self)#, default='clienticon.ico')
        
        
        tk.Tk.wm_title(self, "Graph Page!")
        self.f = f
        
        self.lgd1 = lgd1
        
        
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand = True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
#    
            
        F = PageThree

        frame = F(f, lgd1,container, self)
        

        self.frames[F] = frame
        
        self.geometry("800x700+0+0")

        frame.grid(row=0, column=0, sticky="nsew")
    

        self.show_frame(PageThree)

    def show_frame(self, cont):

        frame = self.frames[cont]
        frame.tkraise()
        
    def close_frames(self):
        
        self.quit()
                        
        self.destroy()
            



class PageThree(tk.Frame):

    def __init__(self,f, lgd1, parent, controller):
        tk.Frame.__init__(self, parent)
        # self.geometry("500x500")
        # label = tk.Label(self, text="Graph Page!", font=LARGE_FONT)
        # label.pack(pady=1,padx=1)

#        button1 = ttk.Button(self, text="Back to Home",
#                            command=lambda: controller.show_frame(StartPage))
#        button1.pack()
        
        self.f = f
        
        self.lgd1 = lgd1
        
        frame1 = tk.Frame(self, borderwidth=2)
        frame1.pack(side=tk.TOP)
        button2=tk.Button(frame1, text='Quit',
                          font=button_fontsize,
                          height = 1, 
                          width = 4,
                          command= lambda: controller.close_frames() )
        
#        button2.winfo_pixels(1000)
        
        button2.pack(side=tk.LEFT)
        

        udpi = int(var1.get())#user_dpi
        
        def save(): 
            
            file = fd.asksaveasfile(filetypes=files, 
        defaultextension='.png', title="Window-2") 
            
            if file:
                if file.name:
                    canvas.figure.savefig(file.name, bbox_extra_artists=(lgd1,), pad_inches=.5, dpi=udpi)
                    # canvas.print_figure(file.name, bbox_extra_artists=(lgd1,), bbox_inches='tight',pad_inches=.5, dpi=100)
                    
            
            
        
        button3=tk.Button(frame1, text='Save',
                          font=button_fontsize,
                          height = 1, 
                          width = 4,
                          command= save )
        
        button3.pack(side=tk.LEFT)
        

        canvas = FigureCanvasTkAgg(f, self)
        # canvas.resize(self, w = 14, h=9)
#        canvas.show()
        canvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=0)

        toolbar = NavigationToolbar2Tk(canvas, frame1)
        toolbar.update()
        toolbar.pack(side="left", fill=tk.X, expand=True)
        
        canvas._tkcanvas.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        
        # print(toolbar.save_figure)
        

# app = SeaofBTCapp()
# app.mainloop()
    





def TK43():
    
    
    # image(wd, ws ,figname ='fig19', nd=8, ns=5, standard=True, color_number=0,
    #       title_font_size=30, angle_fontsize=10, dirction_fontsize=15,
    #       legend_fontsize=10, title1="WindRose",
    #       x1=1, y1=1.1, user_dpi=100, percent_angle=60,
    #       user_opening=0.6, figwidth=6, figheight=8,
    #       rectx = 1, recty=1.2, legx=0.5, legy=0.5,
    #       plot_length=20, image_length=5,
    #       title_font_color='black', angle_font_color='black',
    #       dirction_font_color='black', legend_font_color='black'
    #       )
    
    
    
    try:
        figname1 = 'Fig'
        nd1 = spinT1.get()
        ns1 = spinT2.get()
        standard1 = bool(CheckVarT1.get())
        ColorsN = var2.get()#color_number
        tiFS = var4.get()#title_font_size
        aFS = var7.get()#angle_fontsize
        dFS = var6.get()#dirction_fontsize
        lFS = var5.get()#legend_fontsize
        t1 = Ent1.get() #title1
        xt = float(Ent2.get())#x position of title
        yt = float(Ent3.get())#x position of title
        udpi = int(var1.get())#user_dpi
        pa = int(var3.get())#percent_angle
        uo = float(Ent7.get())#user_opening
        fgW = float(Ent8.get())#figwidth
        fgH = float(Ent9.get())#figheight
        # rtx = float(Ent10.get())#rectx
        # rty = float(Ent11.get())#recty
        xlm = float(Ent10.get())
        lwidth = float(Ent11.get())
        lgX = float(Ent4.get())#legx
        lgY = float(Ent5.get())#legy
        pL = float(Ent16.get())#plot_length
        iL = float(Ent17.get()) #image_length
        tFC = var8.get()#title_font_color
        aFC = var11.get()#angle_font_color
        dFC = var10.get()#dirction_font_color
        lFC = var9.get()#legend_font_color
        
        
        L, lgd1 = image(wd, ws, figname =figname1, nd=int(nd1), ns=int(ns1),
                        standard=standard1, color_number=int(ColorsN),
                        title_font_size=tiFS, angle_fontsize=aFS,
                        dirction_fontsize=dFS, legend_fontsize=lFS,
                        title1=t1, x1=xt, y1=yt, user_dpi=udpi,
                        percent_angle=pa, user_opening=uo, figwidth=fgW,
                        figheight=fgH, Xlim=xlm, Linewidth=lwidth, legx=lgX, legy=lgY,
                        plot_length=pL, image_length=iL, title_font_color=tFC,
                        angle_font_color=aFC, dirction_font_color=dFC,
                        legend_font_color=lFC)
        
            # TK4_3(L, lgd1)
        
        
        
        
            
            
            
        app = SeaofBTCapp(L, lgd1)
        app.mainloop()
        
        
            
        
            
            
        
    except:
        
        tk.messagebox.showinfo("Warning", 
                            'One or more of your input values is/are not valid. \n'+
                            'Or your data is not valid.')
        

        
def TK43_spring():
    
    
    # image(wd, ws ,figname ='fig19', nd=8, ns=5, standard=True, color_number=0,
    #       title_font_size=30, angle_fontsize=10, dirction_fontsize=15,
    #       legend_fontsize=10, title1="WindRose",
    #       x1=1, y1=1.1, user_dpi=100, percent_angle=60,
    #       user_opening=0.6, figwidth=6, figheight=8,
    #       rectx = 1, recty=1.2, legx=0.5, legy=0.5,
    #       plot_length=20, image_length=5,
    #       title_font_color='black', angle_font_color='black',
    #       dirction_font_color='black', legend_font_color='black'
    #       )
    
    
    
    try:
        figname1 = 'Spring'
        nd1 = spinT1.get()
        ns1 = spinT2.get()
        standard1 = bool(CheckVarT1.get())
        ColorsN = var2.get()#color_number
        tiFS = var4.get()#title_font_size
        aFS = var7.get()#angle_fontsize
        dFS = var6.get()#dirction_fontsize
        lFS = var5.get()#legend_fontsize
        t1 = Ent18.get() #title1
        xt = float(Ent2.get())#x position of title
        yt = float(Ent3.get())#x position of title
        udpi = int(var1.get())#user_dpi
        pa = int(var3.get())#percent_angle
        uo = float(Ent7.get())#user_opening
        fgW = float(Ent8.get())#figwidth
        fgH = float(Ent9.get())#figheight
        # rtx = float(Ent10.get())#rectx
        # rty = float(Ent11.get())#recty
        xlm = float(Ent10.get())
        lwidth = float(Ent11.get())
        lgX = float(Ent4.get())#legx
        lgY = float(Ent5.get())#legy
        pL = float(Ent16.get())#plot_length
        iL = float(Ent17.get()) #image_length
        tFC = var8.get()#title_font_color
        aFC = var11.get()#angle_font_color
        dFC = var10.get()#dirction_font_color
        lFC = var9.get()#legend_font_color
        
        
        L, lgd1 = image(awd_sp, aws_sp, figname =figname1, nd=int(nd1), ns=int(ns1),
                        standard=standard1, color_number=int(ColorsN),
                        title_font_size=tiFS, angle_fontsize=aFS,
                        dirction_fontsize=dFS, legend_fontsize=lFS,
                        title1=t1, x1=xt, y1=yt, user_dpi=udpi,
                        percent_angle=pa, user_opening=uo, figwidth=fgW,
                        figheight=fgH, Xlim=xlm, Linewidth=lwidth, legx=lgX, legy=lgY,
                        plot_length=pL, image_length=iL, title_font_color=tFC,
                        angle_font_color=aFC, dirction_font_color=dFC,
                        legend_font_color=lFC)
        
            
            
        app = SeaofBTCapp(L, lgd1)
        app.mainloop()
        
            
        
    except:
        
        tk.messagebox.showinfo("Warning", 
                            'One or more of your input values is/are not valid. \n'+
                            'Or your data is not valid.')




def TK43_summer():
    
    
    # image(wd, ws ,figname ='fig19', nd=8, ns=5, standard=True, color_number=0,
    #       title_font_size=30, angle_fontsize=10, dirction_fontsize=15,
    #       legend_fontsize=10, title1="WindRose",
    #       x1=1, y1=1.1, user_dpi=100, percent_angle=60,
    #       user_opening=0.6, figwidth=6, figheight=8,
    #       rectx = 1, recty=1.2, legx=0.5, legy=0.5,
    #       plot_length=20, image_length=5,
    #       title_font_color='black', angle_font_color='black',
    #       dirction_font_color='black', legend_font_color='black'
    #       )
    
    
    
    try:
        figname1 = 'Summer'
        nd1 = spinT1.get()
        ns1 = spinT2.get()
        standard1 = bool(CheckVarT1.get())
        ColorsN = var2.get()#color_number
        tiFS = var4.get()#title_font_size
        aFS = var7.get()#angle_fontsize
        dFS = var6.get()#dirction_fontsize
        lFS = var5.get()#legend_fontsize
        t1 = Ent19.get() #title1
        xt = float(Ent2.get())#x position of title
        yt = float(Ent3.get())#x position of title
        udpi = int(var1.get())#user_dpi
        pa = int(var3.get())#percent_angle
        uo = float(Ent7.get())#user_opening
        fgW = float(Ent8.get())#figwidth
        fgH = float(Ent9.get())#figheight
        # rtx = float(Ent10.get())#rectx
        # rty = float(Ent11.get())#recty
        xlm = float(Ent10.get())
        lwidth = float(Ent11.get())
        lgX = float(Ent4.get())#legx
        lgY = float(Ent5.get())#legy
        pL = float(Ent16.get())#plot_length
        iL = float(Ent17.get()) #image_length
        tFC = var8.get()#title_font_color
        aFC = var11.get()#angle_font_color
        dFC = var10.get()#dirction_font_color
        lFC = var9.get()#legend_font_color
        
        
        L, lgd1 = image(awd_su, aws_su, figname =figname1, nd=int(nd1), ns=int(ns1),
                        standard=standard1, color_number=int(ColorsN),
                        title_font_size=tiFS, angle_fontsize=aFS,
                        dirction_fontsize=dFS, legend_fontsize=lFS,
                        title1=t1, x1=xt, y1=yt, user_dpi=udpi,
                        percent_angle=pa, user_opening=uo, figwidth=fgW,
                        figheight=fgH, Xlim=xlm, Linewidth=lwidth, legx=lgX, legy=lgY,
                        plot_length=pL, image_length=iL, title_font_color=tFC,
                        angle_font_color=aFC, dirction_font_color=dFC,
                        legend_font_color=lFC)
        
        
            
            
        app = SeaofBTCapp(L, lgd1)
        app.mainloop()
        
        
          
        
    except:
        
        tk.messagebox.showinfo("Warning", 
                            'One or more of your input values is/are not valid. \n'+
                            'Or your data is not valid.')



def TK43_fall():
    
    
    # image(wd, ws ,figname ='fig19', nd=8, ns=5, standard=True, color_number=0,
    #       title_font_size=30, angle_fontsize=10, dirction_fontsize=15,
    #       legend_fontsize=10, title1="WindRose",
    #       x1=1, y1=1.1, user_dpi=100, percent_angle=60,
    #       user_opening=0.6, figwidth=6, figheight=8,
    #       rectx = 1, recty=1.2, legx=0.5, legy=0.5,
    #       plot_length=20, image_length=5,
    #       title_font_color='black', angle_font_color='black',
    #       dirction_font_color='black', legend_font_color='black'
    #       )
    
    
    
    try:
        figname1 = 'Fall'
        nd1 = spinT1.get()
        ns1 = spinT2.get()
        standard1 = bool(CheckVarT1.get())
        ColorsN = var2.get()#color_number
        tiFS = var4.get()#title_font_size
        aFS = var7.get()#angle_fontsize
        dFS = var6.get()#dirction_fontsize
        lFS = var5.get()#legend_fontsize
        t1 = Ent20.get() #title1
        xt = float(Ent2.get())#x position of title
        yt = float(Ent3.get())#x position of title
        udpi = int(var1.get())#user_dpi
        pa = int(var3.get())#percent_angle
        uo = float(Ent7.get())#user_opening
        fgW = float(Ent8.get())#figwidth
        fgH = float(Ent9.get())#figheight
        # rtx = float(Ent10.get())#rectx
        # rty = float(Ent11.get())#recty
        xlm = float(Ent10.get())
        lwidth = float(Ent11.get())
        lgX = float(Ent4.get())#legx
        lgY = float(Ent5.get())#legy
        pL = float(Ent16.get())#plot_length
        iL = float(Ent17.get()) #image_length
        tFC = var8.get()#title_font_color
        aFC = var11.get()#angle_font_color
        dFC = var10.get()#dirction_font_color
        lFC = var9.get()#legend_font_color
        
        
        L, lgd1 = image(awd_fa, aws_fa, figname =figname1, nd=int(nd1), ns=int(ns1),
                        standard=standard1, color_number=int(ColorsN),
                        title_font_size=tiFS, angle_fontsize=aFS,
                        dirction_fontsize=dFS, legend_fontsize=lFS,
                        title1=t1, x1=xt, y1=yt, user_dpi=udpi,
                        percent_angle=pa, user_opening=uo, figwidth=fgW,
                        figheight=fgH, Xlim=xlm, Linewidth=lwidth, legx=lgX, legy=lgY,
                        plot_length=pL, image_length=iL, title_font_color=tFC,
                        angle_font_color=aFC, dirction_font_color=dFC,
                        legend_font_color=lFC)
        
        
            
            
        app = SeaofBTCapp(L, lgd1)
        app.mainloop()
        
        
          
        
    except:
        
        tk.messagebox.showinfo("Warning", 
                            'One or more of your input values is/are not valid. \n'+
                            'Or your data is not valid.')




def TK43_winter():
    
    
    # image(wd, ws ,figname ='fig19', nd=8, ns=5, standard=True, color_number=0,
    #       title_font_size=30, angle_fontsize=10, dirction_fontsize=15,
    #       legend_fontsize=10, title1="WindRose",
    #       x1=1, y1=1.1, user_dpi=100, percent_angle=60,
    #       user_opening=0.6, figwidth=6, figheight=8,
    #       rectx = 1, recty=1.2, legx=0.5, legy=0.5,
    #       plot_length=20, image_length=5,
    #       title_font_color='black', angle_font_color='black',
    #       dirction_font_color='black', legend_font_color='black'
    #       )
    
    
    
    try:
        figname1 = 'Winter'
        nd1 = spinT1.get()
        ns1 = spinT2.get()
        standard1 = bool(CheckVarT1.get())
        ColorsN = var2.get()#color_number
        tiFS = var4.get()#title_font_size
        aFS = var7.get()#angle_fontsize
        dFS = var6.get()#dirction_fontsize
        lFS = var5.get()#legend_fontsize
        t1 = Ent21.get() #title1
        xt = float(Ent2.get())#x position of title
        yt = float(Ent3.get())#x position of title
        udpi = int(var1.get())#user_dpi
        pa = int(var3.get())#percent_angle
        uo = float(Ent7.get())#user_opening
        fgW = float(Ent8.get())#figwidth
        fgH = float(Ent9.get())#figheight
        # rtx = float(Ent10.get())#rectx
        # rty = float(Ent11.get())#recty
        xlm = float(Ent10.get())
        lwidth = float(Ent11.get())
        lgX = float(Ent4.get())#legx
        lgY = float(Ent5.get())#legy
        pL = float(Ent16.get())#plot_length
        iL = float(Ent17.get()) #image_length
        tFC = var8.get()#title_font_color
        aFC = var11.get()#angle_font_color
        dFC = var10.get()#dirction_font_color
        lFC = var9.get()#legend_font_color
        
        
        L, lgd1 = image(awd_wi, aws_wi, figname =figname1, nd=int(nd1), ns=int(ns1),
                        standard=standard1, color_number=int(ColorsN),
                        title_font_size=tiFS, angle_fontsize=aFS,
                        dirction_fontsize=dFS, legend_fontsize=lFS,
                        title1=t1, x1=xt, y1=yt, user_dpi=udpi,
                        percent_angle=pa, user_opening=uo, figwidth=fgW,
                        figheight=fgH, Xlim=xlm, Linewidth=lwidth, legx=lgX, legy=lgY,
                        plot_length=pL, image_length=iL, title_font_color=tFC,
                        angle_font_color=aFC, dirction_font_color=dFC,
                        legend_font_color=lFC)
        
        
            
            
        app = SeaofBTCapp(L, lgd1)
        app.mainloop()
        
        
          
        
    except:
        
        tk.messagebox.showinfo("Warning", 
                            'One or more of your input values is/are not valid. \n'+
                            'Or your data is not valid.')

# global awd_sp, aws_sp, awd_su, aws_su, awd_fa, aws_fa, awd_wi, aws_wi

    


def close_window():
    app1.quit()
    app1.destroy()
    

def enable():
    
    TAB_CONTROL.tab(TAB2, state="normal")



Pics_Formats = {'eps': 'Encapsulated Postscript',
 'jpg': 'Joint Photographic Experts Group',
 'jpeg': 'Joint Photographic Experts Group',
 'pdf': 'Portable Document Format',
 'pgf': 'PGF code for LaTeX',
 'png': 'Portable Network Graphics',
 'ps': 'Postscript',
 'raw': 'Raw RGBA bitmap',
 'rgba': 'Raw RGBA bitmap',
 'svg': 'Scalable Vector Graphics',
 'svgz': 'Scalable Vector Graphics',
 'tif': 'Tagged Image File Format',
 'tiff': 'Tagged Image File Format',
 #'bmp': 'Bitmap File Format'
 }

files = [('All Files', '*.*')]
    
for i in Pics_Formats:
    
    s = ( Pics_Formats[i], '*.' + i)
    
    files.append(s)


 

    


def callback1():
    ftypes = [('excel files', '*.xls'),('excel files', '*.xlsx'), ('All files', '*')]
    TAB_CONTROL.add(TAB2, text='Tab 2', state="disabled")
    global file_path, wd, ws
    name = fd.askopenfilename(filetypes = ftypes) 
    
    txt1.delete("1.0", "end")
    txt1.insert(tk.END, name)
    
    
    
    file_path = name
    try :
        
        colDi = int(Ent12.get()) - 1
        colSp = int(Ent13.get()) - 1
        rowSt = int(Ent14.get()) - 1
        
        mx = int(spinT3.get())
        
        try:
            rowEn = int(Ent15.get()) - 1
            
        except:
            rowEn = '-'

        
        wd, ws, j = excel_reader(name, colDirection=colDi, colSpeed=colSp, rowStart=rowSt, rowEnd=rowEn, Max=mx)
        
        if j >0 :
            
            enable()

        
    except:
        tk.messagebox.showinfo("Warning", 
                           'One or more of your input values is/are not valid. \n'+
                           'Or your data is not valid.')
        print('Exception in callback1')


    
    
    

def callback2():
    global image_path
    ftypes = [('Joint Photographic Experts Group', '*.jpg'),
              ('Joint Photographic Experts Group', '*.jpeg'),
              ('Portable Network Graphics', '*.png'),
              ('Tagged Image File Format', '*.tif'),
              ('Tagged Image File Format','*.tiff'),
              ('bitmap image file', '*.bmp'),
              ('All files', '*')]

    image_path = fd.askopenfilename(filetypes = ftypes) 
    
    
    txt2.insert(tk.END, image_path)
    
    # a = txt1.get(1.0, tk.END)
    
    # txt2.insert(tk.END, a)
    
    # print(a)
    

def seasons_determiner(wb, colDirection, colSpeed, rowStart, rowEnd, Max_speed=40):
    
    print(55)
    
    # wb = xlrd.open_workbook(f_name) 
    
    date_col = int(Ent11p.get()) - 1
    
    
    
    sheet = wb.sheet_by_index(0) 
    
    row = sheet.nrows
    
    wd = np.zeros((row, 1))
    
    print(np.shape(wd))
    
    ws = np.zeros((row, 1))
    
    j = 0
    
    wd_sp = [] ; ws_sp = []
    
    wd_su = [] ; ws_su = []
    
    wd_fa = [] ; ws_fa = []
    
    wd_wi = [] ; ws_wi = []
    
    
    
    for i in range(rowStart, rowEnd):
        
        try:
            
            vel = sheet.cell_value(i, colSpeed)
            dr = sheet.cell_value(i, colDirection)
         
            date = xlrd.xldate_as_tuple(sheet.cell(i, date_col).value,wb.datemode)
         
        # if(isinstance(vel, float) or isinstance(vel, int)) and 
        
         # if (not( (int(dr) == dr) and (vel == 0 or vel == 0.0))) and ((int(dr) == dr) and (isinstance(vel, float) or isinstance(vel, int))):
         # try:
            wd[j, 0] = dr
            ws[j, 0] = vel
            
            if vel > Max_speed :
                
                continue
   
            j += 1
        
        
        
            if date[1]==4 or date[1]==5 or (date[1]==3 and date[2]>20 ) or (date[1]==6 and date[2]<21 ):
                
                wd_sp.append(dr) ; ws_sp.append(vel)
        
            elif date[1]==7 or date[1]==8 or (date[1]==6 and date[2]>20 ) or (date[1]==9 and date[2]<23 ):
           
                wd_su.append(dr) ; ws_su.append(vel)
           
            elif date[1]==10 or date[1]==11 or (date[1]==9 and date[2]>22 ) or (date[1]==12 and date[2]<21 ):
           
                wd_fa.append(dr) ; ws_fa.append(vel)
           
            elif date[1]==1 or date[1]==2 or (date[1]==12 and date[2]>20 ) or (date[1]==3 and date[2]<21 ):
           
                wd_wi.append(dr) ; ws_wi.append(vel)
            
                
            
             
              
         
        except:
            pass
    
    wd = wd[0:j, 0]
    
    print(np.shape(wd))
    
    ws = ws[0:j, 0]   
    
    global awd_sp, aws_sp, awd_su, aws_su, awd_fa, aws_fa, awd_wi, aws_wi
    
    awd_sp = np.array(wd_sp) ; aws_sp = np.array(ws_sp)
    
    awd_su = np.array(wd_su) ; aws_su = np.array(ws_su)
    
    awd_fa = np.array(wd_fa) ; aws_fa = np.array(ws_fa)
    
    awd_wi = np.array(wd_wi) ; aws_wi = np.array(ws_wi)
    
    print(j, j)






app1 = tkinter.Tk()
app1.geometry("800x600")
app1.title('Plotting of Windrose Graphs')


def activate_speeds():
    
    if spinT2.config()['state'][4] == 'disabled':
        spinT2.config ( state = "readonly")
        
    else:
        
        spinT2.config ( state = "disabled")
    

    
def activate_buttons():
    
    buttonList = [button2, button3, button4, button5]
    
    for button in buttonList:
        
    
        if button.config()['state'][4] == 'disabled':
            button.config ( state = "normal")
            
        else:
            
            button.config ( state = "disabled")
    
    entList = [Ent18, Ent19, Ent20, Ent21]
    
    textList = ["Spring", "Summer", "Autumn", "Winter"]
    
    i = 0
    
    for ent in entList:
        
    
        if ent.config()['state'][4] == 'disabled':
            ent.config ( state = "normal")
            ent.delete(0, tk.END)
            ent.insert(0, textList[i])
            
            i += 1
            
        else:
            
            ent.config ( state = "disabled")
    
#    print(Ent18.config()['state'][4])
        
    



TAB_CONTROL = ttk.Notebook(app1)
#Tab1
TAB1 = ttk.Frame(TAB_CONTROL)
TAB_CONTROL.add(TAB1, text='Tab 1')

# button1=tk.Button(TAB1,text='Quit',command=close_window   )
# button1.grid(row=1,column=0,sticky='W',padx=50,pady=6)




frameT0 = tk.Frame(TAB1, height=1, borderwidth=2)#, bg='red')
frameT0.grid(row=0, column=0, sticky='W' ,padx=50,pady=4)#.pack(side=tk.TOP)

CheckVarT0 = tk.IntVar()
ChT0 = tk.Checkbutton(frameT0, text = "Seasons", variable = CheckVarT0, \
                 onvalue = 1, offvalue = 0, height=2, \
                 width = 7, bg='red' , command=activate_buttons)

ChT0.grid(row=0, column=0, sticky='W')


CheckVarT01 = tk.IntVar()
ChT01 = tk.Checkbutton(frameT0, text = "Autosave Plots", variable = CheckVarT01, \
                 onvalue = 1, offvalue = 0, height=2, \
                 width = 10, bg='lightgreen' )

ChT01.grid(row=0, column=1, sticky='W')


Lb11 = tk.Label(frameT0, text="Columns of Data:  ")
Lb11.grid(row=0,column=2,sticky='W')#,padx=5, pady=4)

Lb11p = tk.Label(frameT0, text="Date:  ")
Lb11p.grid(row=0,column=3,sticky='W')#,padx=5, pady=4)


Ent11p = tk.Entry(frameT0, width=4)
Ent11p.grid(row=0, column=4, sticky='w')#,padx=5, pady=4)
#Ent5.place(x=225, y=85)

Ent11p.delete(0, tk.END)
Ent11p.insert(0, "2")


Lb12 = tk.Label(frameT0, text="Directions:  ")
Lb12.grid(row=0,column=5,sticky='W')#,padx=5, pady=4)


Ent12 = tk.Entry(frameT0, width=4)
Ent12.grid(row=0, column=6, sticky='w')#,padx=5, pady=4)
#Ent5.place(x=225, y=85)

Ent12.delete(0, tk.END)
Ent12.insert(0, "4")


Lb13 = tk.Label(frameT0, text="Speeds:  ")
Lb13.grid(row=0,column=7,sticky='W')#,padx=5, pady=4)


Ent13 = tk.Entry(frameT0, width=4)
Ent13.grid(row=0, column=8, sticky='w')#,padx=5, pady=4)
#Ent5.place(x=225, y=85)

Ent13.delete(0, tk.END)
Ent13.insert(0, "3")


Lb36 = tk.Label(frameT0, text=" Maximum:  ")
Lb36.grid(row=0, column=9, sticky='W')



varT3 = tk.StringVar(frameT0)
varT3.set("40")

spinT3 = tk.Spinbox(frameT0, width=4, from_= 10, to = 40, textvariable=varT3, state='readonly')#, textvariable=text_variable)  

tkvalidate.int_validate(spinT3, from_=10, to=40)
  
spinT3.grid(row=0,column=10,sticky='W',padx=5,pady=4)




frameT00 = tk.Frame(TAB1, height=1, borderwidth=2)#, bg='red')
frameT00.grid(row=1, column=0, sticky='W' ,padx=50,pady=4)#.pack(side=tk.TOP)


Lb14 = tk.Label(frameT00, text="Rows of data:  ")
Lb14.grid(row=0,column=0,sticky='W')#,padx=5, pady=4)

Lb15 = tk.Label(frameT00, text="Start:  ")
Lb15.grid(row=0,column=1,sticky='W')#,padx=5, pady=4)


Ent14 = tk.Entry(frameT00, width=4)
Ent14.grid(row=0, column=2, sticky='w')#,padx=5, pady=4)
#Ent5.place(x=225, y=85)

Ent14.delete(0, tk.END)
Ent14.insert(0, "2")



Lb16 = tk.Label(frameT00, text="End:  ")
Lb16.grid(row=0,column=3,sticky='W')#,padx=5, pady=4)




Ent15 = tk.Entry(frameT00, width=4)
Ent15.grid(row=0, column=4, sticky='w')#,padx=5, pady=4)
#Ent5.place(x=225, y=85)

Ent15.delete(0, tk.END)
Ent15.insert(0, "-")


Lb17 = tk.Label(frameT00, text="Ratios:  ")
Lb17.grid(row=0,column=5,sticky='W',padx=1)#, pady=4)

Lb18 = tk.Label(frameT00, text="Length of Plot:  ")
Lb18.grid(row=0,column=6,sticky='W')#,padx=5, pady=4)


Ent16 = tk.Entry(frameT00, width=4)
Ent16.grid(row=0, column=7, sticky='w')#,padx=5, pady=4)
#Ent5.place(x=225, y=85)

Ent16.delete(0, tk.END)
Ent16.insert(0, "20")



Lb19 = tk.Label(frameT00, text="  Length of Image:  ")
Lb19.grid(row=0,column=8,sticky='W')#,padx=5, pady=4)


Ent17 = tk.Entry(frameT00, width=4)
Ent17.grid(row=0, column=9, sticky='w',padx=7)#, pady=4)
#Ent5.place(x=225, y=85)

Ent17.delete(0, tk.END)
Ent17.insert(0, "4")












frameT2 = tk.Frame(TAB1, height=1, borderwidth=2)#, bg='red')
frameT2.grid(row=2, column=0, sticky='W' ,padx=40,pady=4)#.pack(side=tk.TOP)




btn1 = tk.Button(frameT2, text='Open Data File', 
       command=callback1)

btn1.grid(row=0, column=0, padx=12,pady=4)

txt1 = tk.Text(frameT2, height=1, width=52)

txt1.grid(row=0, column=1, sticky='E')


frameT3 = tk.Frame(TAB1, height=1, borderwidth=2)#, bg='red')
frameT3.grid(row=3, column=0, sticky='W' ,padx=42,pady=4)#.pack(side=tk.TOP)


btn2 = tk.Button(frameT3, text='Open Image File', 
       command=callback2)

btn2.grid(row=0, column=0, padx=10,pady=4)

txt2 = tk.Text(frameT3, height=1, width=51)

txt2.grid(row=0, column=1)




Lbf4 = tk.LabelFrame(TAB1, text="Titles:")
Lbf4.grid(row=4, column=0, sticky='W',padx=50)


Lb1 = tk.Label(Lbf4, text="Title of Total Plot:")
Lb1.grid(row=0,column=0,sticky='W')#,padx=5, pady=4)

Ent1 = tk.Entry(Lbf4, width=67)
Ent1.grid(row=0,column=1,sticky='W')#,padx=5, pady=4)

Ent1.delete(0, tk.END)
Ent1.insert(0, "Windrose")



Lb26 = tk.Label(Lbf4, text="Title of Spring Plot:", bg='lawngreen', anchor="w")
Lb26.grid(row=1,column=0,sticky='W', ipadx=5)#,padx=5, pady=4)

Ent18 = tk.Entry(Lbf4, width=67, state="disabled")
Ent18.grid(row=1,column=1,sticky='W')#,padx=5, pady=4)



Lb27 = tk.Label(Lbf4, text="Title of Summer Plot:", bg='salmon')
Lb27.grid(row=2,column=0,sticky='W')#,padx=5, pady=4)

Ent19 = tk.Entry(Lbf4, width=67, state="disabled")
Ent19.grid(row=2,column=1,sticky='W')#,padx=5, pady=4)




Lb28 = tk.Label(Lbf4, text="Title of Autumn Plot:", bg='yellow')
Lb28.grid(row=3,column=0,sticky='W')#,padx=5, pady=4)

Ent20 = tk.Entry(Lbf4, width=67, state="disabled")
Ent20.grid(row=3,column=1,sticky='W')#,padx=5, pady=4)




Lb29 = tk.Label(Lbf4, text="Title of Winter Plot:", bg='skyblue', anchor="w")
Lb29.grid(row=4,column=0,sticky='W', ipadx=5)#,padx=5, pady=4)

Ent21 = tk.Entry(Lbf4, width=67, state="disabled")
Ent21.grid(row=4,column=1,sticky='W')#,padx=5, pady=4)



Lbf5 = tk.LabelFrame(TAB1, text="Location of Image:")
Lbf5.grid(row=5, column=0, sticky='W',padx=50)



varR = tk.IntVar()

varR.set(1)

R1 = tk.Radiobutton(Lbf5, text="Bottom Left", variable=varR, value=1)
R1.grid(row=0, column=0, sticky='W',padx=22)

R2 = tk.Radiobutton(Lbf5, text="Bottom Right", variable=varR, value=2)
R2.grid(row=0, column=1, sticky='W',padx=22)

R3 = tk.Radiobutton(Lbf5, text="Top Left", variable=varR, value=3)
R3.grid(row=0, column=2, sticky='W',padx=23)

R4 = tk.Radiobutton(Lbf5, text="Top Right", variable=varR, value=4)
R4.grid(row=0, column=3, sticky='W',padx=23)



# btn3=tkinter.Button(TAB1,text='Tab 2',width=10,
#                             height=4,
#                             command=enable   )

# btn3.grid(row=6,column=0,sticky='W',padx=50,pady=4)











####################Tab2####################
############################################


TAB2 = ttk.Frame(TAB_CONTROL)
TAB_CONTROL.add(TAB2, text='Tab 2', state="disabled")
TAB_CONTROL.pack(expand=1, fill="both")




frame1 = tk.Frame(TAB2, borderwidth=2)
frame1.grid(row=0, column=0, sticky='W',padx=50,pady=4)#.pack(side=tk.TOP)

#Lb0 = tk.Label(frame1, text="Name of Figure:  ")
#Lb0.grid(row=0,column=0,sticky='W')#,padx=5, pady=4)
#
#Ent0 = tk.Entry(frame1, width=50)
#Ent0.grid(row=0,column=1,sticky='W')#,padx=5, pady=4)
#
#Ent0.delete(0, tk.END)
#Ent0.insert(0, "Fig1")






frameT1 = tk.Frame(TAB2, height=1, borderwidth=2)#, bg='red')
frameT1.grid(row=1, column=0, sticky='w' ,padx=50,pady=4)#.pack(side=tk.TOP)


CheckVarT1 = tk.IntVar()
ChT1 = tk.Checkbutton(frameT1, text = "Standard", variable = CheckVarT1, \
                 onvalue = 1, offvalue = 0, height=2, \
                 width = 6, bg='red' , command=activate_speeds)

ChT1.grid(row=0, column=0, sticky='W')


LbT0 = tk.Label(frameT1, text="  Directions:  ")
LbT0.grid(row=0,column=1,sticky='W')#,padx=5, pady=4)

varT1 = tk.StringVar(frameT1)
varT1.set("4")
spinT1 = tk.Spinbox(frameT1, width=4, values=(4,8,16,32), textvariable=varT1,
                    state='readonly')#, textvariable=text_variable)  

tkvalidate.int_validate(spinT1, from_=4, to=32)
  
spinT1.grid(row=0,column=2,sticky='W',padx=5,pady=4) 




LbT1 = tk.Label(frameT1, text="  Speeds:  ")
LbT1.grid(row=0,column=3,sticky='W')#,padx=5, pady=4)

varT2 = tk.StringVar(frameT1)
varT2.set("4")
spinT2 = tk.Spinbox(frameT1, width=4, from_= 2, to = 20, textvariable=varT2, state='readonly')#, textvariable=text_variable)  

tkvalidate.int_validate(spinT2, from_=2, to=20)
  
spinT2.grid(row=0,column=4,sticky='W',padx=5,pady=4) 






Lb5 = tk.Label(frameT1, text="     dpi: ")
Lb5.grid(row=0, column=7, sticky='W')



var1 = tk.StringVar(TAB2)
var1.set("100")
spin1 = tk.Spinbox(frameT1, width=4, from_= 10, to = 1000, textvariable=var1)#, textvariable=text_variable)  

tkvalidate.int_validate(spin1, from_=0, to=1000)
  
spin1.grid(row=0,column=8,sticky='W',padx=5,pady=4) 








frame2 = tk.Frame(TAB2, borderwidth=2, height=10)
frame2.grid(row=2, column=0, sticky='W',padx=50,pady=4)#.pack(side=tk.TOP)



Lb2 = tk.Label(frame2, text="Position of Title: ")
Lb2.grid(row=0,column=0,sticky='W')#,padx=5, pady=4)

Lb3 = tk.Label(frame2, text="X =")
Lb3.grid(row=0,column=1,sticky='W')

Ent2 = tk.Entry(frame2, width=4)
Ent2.grid(row=0, column=2, sticky='W',padx=13, pady=4)

Ent2.delete(0, tk.END)
Ent2.insert(0, "0.5")

Lb4 = tk.Label(frame2, text="Y =")
Lb4.grid(row=0, column=3, sticky='W')


Ent3 = tk.Entry(frame2, width=4)
Ent3.grid(row=0, column=4, sticky='W',padx=13, pady=4)

Ent3.delete(0, tk.END)
Ent3.insert(0, "1.1")



Lb8 = tk.Label(frame2, text="  Legend Position: ")
Lb8.grid(row=0, column=5, sticky='W')



Lb9 = tk.Label(frame2, text="  X =  ")
Lb9.grid(row=0,column=6,sticky='W')

Ent4 = tk.Entry(frame2, width=4)
Ent4.grid(row=0, column=7, sticky='W',padx=12, pady=4)
#Ent4.place(x=135, y=85)

Ent4.delete(0, tk.END)
Ent4.insert(0, "1")

Lb10 = tk.Label(frame2, text="Y =  ")
Lb10.grid(row=0, column=8, sticky='W')


Ent5 = tk.Entry(frame2, width=4)
Ent5.grid(row=0, column=9, sticky='',padx=12, pady=4)
#Ent5.place(x=225, y=85)

Ent5.delete(0, tk.END)
Ent5.insert(0, "0.5")






frame3 = tk.Frame(TAB2, height=1, borderwidth=2)#, bg='red')
frame3.grid(row=4, column=0, sticky='N' ,padx=50,pady=4)#.pack(side=tk.TOP)




Lb12 = tk.Label(frame3, text=" Width of Sectors:  ")
Lb12.grid(row=0, column=0, sticky='W')



Ent7 = tk.Entry(frame3, width=4)
Ent7.grid(row=0, column=1, sticky='w',padx=20, pady=4)
#Ent5.place(x=225, y=85)

Ent7.delete(0, tk.END)
Ent7.insert(0, 0.3)

tkvalidate.float_validate(Ent7, from_=0, to=1)

Lb6 = tk.Label(frame3, text="Colors:")
Lb6.grid(row=0, column=2, sticky='W')



var2 = tk.StringVar(TAB2)
var2.set("0")
spin2 = tk.Spinbox(frame3, width =4,from_= 0, to = 19, textvariable=var2)#, textvariable=text_variable)  

tkvalidate.int_validate(spin2, from_=0, to=19)
  
spin2.grid(row=0,column=3,sticky='W',padx=20,pady=4) 



Lb7 = tk.Label(frame3, text="   Writing_Percent_Angles")
Lb7.grid(row=0, column=4, sticky='W')


var3 = tk.StringVar(TAB2)
var3.set("60")
spin3 = tk.Spinbox(frame3, width =4, from_= 0, to = 359, textvariable=var3)#, textvariable=text_variable)  

tkvalidate.int_validate(spin3, from_=0, to=359)
  
spin3.grid(row=0,column=5,sticky='W',padx=20,pady=4) 


Lbalaki1 = tk.Label(frame3, text="                            ")
Lbalaki1.grid(row=0, column=6, sticky='W')








# Lbalaki2 = tk.Label(frame4, text="                               ")
# Lbalaki2.grid(row=0, column=7, sticky='W')


Lbf0 = tk.LabelFrame(TAB2, text="Adjusting Size of Plot:", fg='orange')
Lbf0.grid(row=5, column=0, sticky='W',padx=50)



Lb13 = tk.Label(Lbf0, text="Width of Figure: ")
Lb13.grid(row=0, column=0, sticky='W')


Ent8 = tk.Entry(Lbf0, width=4)
Ent8.grid(row=0, column=1, sticky='w',padx=8, pady=4)
#Ent5.place(x=225, y=85)

Ent8.delete(0, tk.END)
Ent8.insert(0, "8")




Lb14 = tk.Label(Lbf0, text="Height of Figure: ")
Lb14.grid(row=0, column=2, sticky='W')


Ent9 = tk.Entry(Lbf0, width=4)
Ent9.grid(row=0, column=3, sticky='w',padx=8, pady=4)
#Ent5.place(x=225, y=85)

Ent9.delete(0, tk.END)
Ent9.insert(0, "6")


# Lb19 = tk.Label(Lbf0, text="  Bounding Box:  ")
# Lb19.grid(row=0, column=4, sticky='W')

Lb20 = tk.Label(Lbf0, text="XLimit ")
Lb20.grid(row=0, column=5, sticky='W')


Ent10 = tk.Entry(Lbf0, width=4)
Ent10.grid(row=0, column=6, sticky='w',padx=8, pady=4)
#Ent5.place(x=225, y=85)

Ent10.delete(0, tk.END)
Ent10.insert(0, "10")

Lb21 = tk.Label(Lbf0, text="Circles' Line-Width")
Lb21.grid(row=0, column=7, sticky='W')

Ent11 = tk.Entry(Lbf0, width=4)
Ent11.grid(row=0, column=8, sticky='w',padx=8, pady=4)
#Ent5.place(x=225, y=85)

Ent11.delete(0, tk.END)
Ent11.insert(0, "1")



Lbf00 = tk.LabelFrame(TAB2, text="Configure subplots:", fg='blue')
Lbf00.grid(row=6, column=0, sticky='W',padx=50)

#plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=2, hspace=.5)


Lb30 = tk.Label(Lbf00, text="Bottom:  ")
Lb30.grid(row=0, column=0, sticky='W')

Ent22 = tk.Entry(Lbf00, width=4)
Ent22.grid(row=0, column=1, sticky='w',padx=1, pady=4)
#Ent5.place(x=225, y=85)

Ent22.delete(0, tk.END)
Ent22.insert(0, "0")


Lb31 = tk.Label(Lbf00, text="Left:  ")
Lb31.grid(row=0, column=2, sticky='W')

Ent23 = tk.Entry(Lbf00, width=4)
Ent23.grid(row=0, column=3, sticky='w',padx=1, pady=4)
#Ent5.place(x=225, y=85)

Ent23.delete(0, tk.END)
Ent23.insert(0, "0")


Lb32 = tk.Label(Lbf00, text="Right:  ")
Lb32.grid(row=0, column=4, sticky='W')

Ent24 = tk.Entry(Lbf00, width=4)
Ent24.grid(row=0, column=5, sticky='w',padx=1, pady=4)
#Ent5.place(x=225, y=85)

Ent24.delete(0, tk.END)
Ent24.insert(0, "1")


Lb33 = tk.Label(Lbf00, text="Top:  ")
Lb33.grid(row=0, column=6, sticky='W')

Ent25 = tk.Entry(Lbf00, width=4)
Ent25.grid(row=0, column=7, sticky='w',padx=1, pady=4)
#Ent5.place(x=225, y=85)

Ent25.delete(0, tk.END)
Ent25.insert(0, ".9")


Lb34 = tk.Label(Lbf00, text="Width Space:  ")
Lb34.grid(row=0, column=8, sticky='W')

Ent26 = tk.Entry(Lbf00, width=4)
Ent26.grid(row=0, column=9, sticky='w',padx=1, pady=4)
#Ent5.place(x=225, y=85)

Ent26.delete(0, tk.END)
Ent26.insert(0, "1")



Lb35 = tk.Label(Lbf00, text="Height Space:  ")
Lb35.grid(row=0, column=10, sticky='W')

Ent27 = tk.Entry(Lbf00, width=4)
Ent27.grid(row=0, column=11, sticky='w',padx=1, pady=4)
#Ent5.place(x=225, y=85)

Ent27.delete(0, tk.END)
Ent27.insert(0, "0.5")
    




Lbf1 = tk.LabelFrame(TAB2, text="Fonts:", fg='green')
Lbf1.grid(row=7, column=0, sticky='W', padx=50)

Lb15 = tk.Label(Lbf1, text="Title: ")
Lb15.grid(row=0, column=0, sticky='N')


var4 = tk.StringVar(TAB2)
var4.set("20")
spin4 = tk.Spinbox(Lbf1, width=4, from_= 1, to = 100, textvariable=var4)#, textvariable=text_variable)  

tkvalidate.int_validate(spin4, from_=1, to=100)
  
spin4.grid(row=0,column=1,sticky='W',padx=17,pady=4) 




Lb16 = tk.Label(Lbf1, text="Legend: ")
Lb16.grid(row=0, column=2, sticky='N')


var5 = tk.StringVar(TAB2)
var5.set("12")
spin5 = tk.Spinbox(Lbf1, width=4, from_= 1, to = 100, textvariable=var5)#, textvariable=text_variable)  

tkvalidate.int_validate(spin5, from_=1, to=100)
  
spin5.grid(row=0, column=3, sticky='W', padx=17, pady=4) 




Lb17 = tk.Label(Lbf1, text="Directions: ")
Lb17.grid(row=0, column=4, sticky='N')


var6 = tk.StringVar(TAB2)
var6.set("12")
spin6 = tk.Spinbox(Lbf1, width=4, from_= 1, to = 100, textvariable=var6)#, textvariable=text_variable)  

tkvalidate.int_validate(spin6, from_=1, to=100)
  
spin6.grid(row=0, column=5, sticky='W', padx=17, pady=4) 






Lb18 = tk.Label(Lbf1, text="Percentage: ")
Lb18.grid(row=0, column=6, sticky='N')


var7 = tk.StringVar(TAB2)
var7.set("12")
spin7 = tk.Spinbox(Lbf1, width=4, from_= 1, to = 100, textvariable=var7)
#, textvariable=text_variable)  

tkvalidate.int_validate(spin7, from_=1, to=100)
  
spin7.grid(row=0, column=7, sticky='W', padx=17, pady=4) 






Lbf2 = tk.LabelFrame(TAB2, text="Colors of Fonts:")
Lbf2.grid(row=8, column=0, sticky='W',padx=50)




Lb22 = tk.Label(Lbf2, text="Title: ")
Lb22.grid(row=0, column=0, sticky='N')


fontColorRange = ['black', 'blue',  'cyan', 'green', 'indigo',  'magenta',  
                  'orange', 'purple', 'red',  'yellow']
var8 = tk.StringVar(TAB2)
var8.set("black")

spin8 = tk.Spinbox(Lbf2, width=max(len(i) for i in (fontColorRange))+1, 
                   values=fontColorRange, textvariable=var8, state='readonly')#, textvariable=text_variable)  

# tkvalidate.string_validate(spin8, values=fontColorRange)
  
spin8.grid(row=0,column=1,sticky='W',padx=5,pady=4) 


Lb23 = tk.Label(Lbf2, text="Legend: ")
Lb23.grid(row=0, column=2, sticky='N')

var9 = tk.StringVar(TAB2)
var9.set("black")

spin9 = tk.Spinbox(Lbf2, width=max(len(i) for i in (fontColorRange))+1, 
                   values=fontColorRange, textvariable=var9, state='readonly')#, textvariable=text_variable)  

# tkvalidate.string_validate(spin8, values=fontColorRange)
  
spin9.grid(row=0,column=3,sticky='W',padx=5,pady=4) 


Lb24 = tk.Label(Lbf2, text="Directions: ")
Lb24.grid(row=0, column=4, sticky='N')

var10 = tk.StringVar(TAB2)
var10.set("black")
spin10 = tk.Spinbox(Lbf2, width=max(len(i) for i in (fontColorRange))+1, 
                    values=fontColorRange, textvariable=var10, state='readonly')#, textvariable=text_variable)  

# tkvalidate.string_validate(spin8, values=fontColorRange)
  
spin10.grid(row=0,column=5,sticky='W',padx=5,pady=4) 


Lb25 = tk.Label(Lbf2, text="Percentage: ")
Lb25.grid(row=0, column=6, sticky='N')

var11 = tk.StringVar(TAB2)
var11.set("black")

spin11 = tk.Spinbox(Lbf2, width=max(len(i) for i in (fontColorRange))+1,
                    values=fontColorRange, textvariable=var11, state='readonly')#, textvariable=text_variable)  

# tkvalidate.string_validate(spin8, values=fontColorRange)
  
spin11.grid(row=0,column=7,sticky='W',padx=5,pady=4) 







Lbf3 = tk.LabelFrame(TAB2, text="Plot Buttons:")
Lbf3.grid(row=9, column=0, sticky='W',padx=50)


def func(TAB2, _event=None):
        TK43()

button1 = tkinter.Button(Lbf3,
                            text="Plot",
                            width=10,
                            height=4,  command = TK43)


button1.grid(row=0,column=0,sticky='W',padx=12,pady=4)




##############shortcut key in Tab2





TAB2.bind('<F5>', func)


######################



button2 = tkinter.Button(Lbf3,
                            text="Spring",
                            width=10,
                            height=4, bg='lawngreen',  command = TK43_spring,
                            state="disabled")


button2.grid(row=0,column=1,sticky='W',padx=12,pady=4)



button3 = tkinter.Button(Lbf3,
                            text="Summer",
                            width=10,
                            height=4, bg='salmon',  command = TK43_summer,
                            state="disabled")


button3.grid(row=0,column=2,sticky='W',padx=12,pady=4)




button4 = tkinter.Button(Lbf3,
                            text="Fall",
                            width=10,
                            height=4, bg='yellow',  command = TK43_fall,
                            state="disabled")


button4.grid(row=0,column=3,sticky='W',padx=11,pady=4)



button5 = tkinter.Button(Lbf3,
                            text="Winter",
                            width=10,
                            height=4, bg='skyblue',  command = TK43_winter,
                            state="disabled")


button5.grid(row=0,column=4,sticky='W',padx=11,pady=4)









button6=tkinter.Button(TAB2,text='Quit',width=10,
                            height=4,
                            command=close_window   )

button6.grid(row=10,column=0,sticky='W',padx=50,pady=4)



def About():
    tk.messagebox.showinfo("About", 
                           'This software is written by'+
                           '  "Seyed Abdolvahab Taghavi (s.av.taghavi@gmail.com)".'+'\nIt plots windrose'+
                           ' of data.')




def NewCommand():
    
    txt1.delete("1.0", "end")
    txt2.delete("1.0", "end")
    

    ChT0.deselect()
    
    buttonList = [button2, button3, button4, button5]
    
    for button in buttonList:
        
        button.config ( state = "disabled")
    
    entList = [Ent18, Ent19, Ent20, Ent21]
    
    
         
    for ent in entList:
        
    
        ent.config ( state = "disabled")
    
    TAB_CONTROL.add(TAB2, text='Tab 2', state="disabled")
    
#    TAB_CONTROL.focus_force()
    TAB_CONTROL.select(TAB1)
    
    global ws, wd, awd_sp, aws_sp, awd_su, aws_su, awd_fa, aws_fa, awd_wi, aws_wi, image_path
    
    
    try:
        
        
        del image_path, ws, wd, awd_sp, aws_sp, awd_su, aws_su, awd_fa, aws_fa, awd_wi, aws_wi
        
         
    except :
        pass


menuBar = Menu(app1) # 1
app1.config(menu=menuBar)
# Now we add a menu to the bar and also assign a menu item to the menu.
fileMenu = Menu(menuBar, tearoff=0) # 2


fileMenu.add_command(label="New", command=NewCommand)
fileMenu.add_separator()
fileMenu.add_command(label="Exit", command=close_window) # 3

menuBar.add_cascade(label="File", menu=fileMenu)


helpMenu = Menu(menuBar, tearoff=0) # 6
helpMenu.add_command(label="About", command=About)
menuBar.add_cascade(label="Help", menu=helpMenu)


# style = ttk.Style(app1)
# style.configure('TLabel', background='black', foreground='white')
# style.configure('TFrame', background='lightblue')
app1.resizable(0, 0)
app1.mainloop()



'''for date in xlrd: asp = xlrd.xldate_as_tuple(sheet.cell(3,1).value,wb.datemode)

start of spring : 21 March (3, 21)

end of spring : 20 June (6, 20)

start of summer : 21 June (6, 21)

end of summer : 22 September (9, 22)f

start of fall : 23 September (9, 23)

end of fall : 22 December (12, 20)

start of winter : 23 December (12, 21)

end of winter : 22 March (3, 20)

'''








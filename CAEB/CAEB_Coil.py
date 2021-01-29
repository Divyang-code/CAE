#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jan  6 14:52:23 2021

@author: divyang
"""
import pandas as pd
import numpy as np

#%% Generate data

class Coil:
    
    def __init__(self, Total_Coils, TurnsInCoil, I_in_1_turn, height, width, thickness, x_distance, width_leg1, width_leg2, width_leg3, width_leg4):
        
        self.Total_Coils = Total_Coils
        self.TurnsInCoil = TurnsInCoil
        self.I_in_1_turn = I_in_1_turn
        self.height      = height
        self.width       = width
        self.thickness   = thickness
        self.x_distance  = x_distance
        self.width_leg1  = width_leg1
        self.width_leg2  = width_leg2
        self.width_leg3  = width_leg3
        self.width_leg4  = width_leg4
        
        print('Number of components: ' + str(Total_Coils* 4))
        print('Copy this number and Builder.xlsx data into CAE_1.xlsx')
        
    def leg1(self):
        
        # centre
        xc = self.x_distance + self.width_leg1/2
        yc = 0
        zc = 0
        
        # dimensions
        l = self.thickness
        b = self.width_leg1
        h = self.height - self.width_leg2 - self.width_leg4
        
        # eular rotaion angles
        Rx = 0
        Ry = 0
        Rz = 0
        
        # current density
        J = self.TurnsInCoil* self.I_in_1_turn/ (l* b)
        
        return [xc, yc, zc, l, b, h, Rx, Ry, Rz, J]
    
    def leg2(self):
        
        # centre
        xc = self.x_distance + self.width/2
        yc = 0
        zc = self.height/2 - self.width_leg2/2
        
        # dimensions
        l = self.thickness
        b = self.width_leg2
        h = self.width
        
        # eular rotaion angles
        Rx = 0
        Ry = 90
        Rz = 0
        
        # current density
        J = self.TurnsInCoil* self.I_in_1_turn/ (l* b)
        
        return [xc, yc, zc, l, b, h, Rx, Ry, Rz, J]
    
    def leg3(self):
        
        # centre
        xc = self.x_distance + self.width - self.width_leg3/2
        yc = 0
        zc = 0
        
        # dimensions
        l = self.thickness
        b = self.width_leg3
        h = self.height - self.width_leg2 - self.width_leg4
        
        # eular rotaion angles
        Rx = 0
        Ry = 180
        Rz = 0
        
        # current density
        J = self.TurnsInCoil* self.I_in_1_turn/ (l* b)
        
        return [xc, yc, zc, l, b, h, Rx, Ry, Rz, J]

    def leg4(self):
        
        # centre
        xc = self.x_distance + self.width/2
        yc = 0
        zc = - self.height/2 + self.width_leg4/2
        
        # dimensions
        l = self.thickness
        b = self.width_leg4
        h = self.width
        
        # eular rotaion angles
        Rx = 0
        Ry = 270
        Rz = 0
        
        # current density
        J = self.TurnsInCoil* self.I_in_1_turn/ (l* b)
        
        return [xc, yc, zc, l, b, h, Rx, Ry, Rz, J]

    def Data_coil(self):
        
        l1 = self.leg1()
        l2 = self.leg2()
        l3 = self.leg3()
        l4 = self.leg4()

        xc = [ l1[0], l2[0], l3[0], l4[0] ]
        yc = [ l1[1], l2[1], l3[1], l4[1] ]
        zc = [ l1[2], l2[2], l3[2], l4[2] ]
        
        l  = [ l1[3], l2[3], l3[3], l4[3] ]
        b  = [ l1[4], l2[4], l3[4], l4[4] ]
        h  = [ l1[5], l2[5], l3[5], l4[5] ]
        
        Rx = [ l1[6], l2[6], l3[6], l4[6] ]
        Ry = [ l1[7], l2[7], l3[7], l4[7] ]
        Rz = [ l1[8], l2[8], l3[8], l4[8] ]
        
        J  = [ l1[9], l2[9], l3[9], l4[9] ]
        
        return xc, yc, zc, l, b, h, Rx, Ry, Rz, J
    
    def Coil_System(self):
        
        ang = np.arange(0, 360, 360/ self.Total_Coils)
        xc, yc, zc, l, b, h, Rx, Ry, Rz, J = self.Data_coil()
        
        # Unchanged variables
        zc = zc* self.Total_Coils
        l  = l*  self.Total_Coils
        b  = b*  self.Total_Coils
        h  = h*  self.Total_Coils
        Rx = Rx* self.Total_Coils
        Ry = Ry* self.Total_Coils
        J  = J*  self.Total_Coils
        
        Rz = []
        for i in ang:
            for j in range(0, 4):
                Rz.append(i)
        
        # x and y
        R = xc* self.Total_Coils
        xc, yc = [], []
        for r, phi in zip(R, Rz):
            xc.append(r* np.cos(phi* np.pi/180))
            yc.append(r* np.sin(phi* np.pi/180))
        
        return xc, yc, zc, l, b, h, Rx, Ry, Rz, J

    def Write_data(self):
        
        xc, yc, zc, l, b, h, Rx, Ry, Rz, J = self.Coil_System()
        
        Out = pd.DataFrame({'x': (['X center'] + xc),
                            'y': (['Y center'] + yc),
                            'z': (['Z center'] + zc),
                            'empty1': np.nan,
                            'l': (['Length']  + l),
                            'b': (['Breadth'] + b),
                            'h': (['Height']  + h),
                            'empty2': np.nan,
                            'pye': (['X rotation'] + Rx),
                            'tta': (['Y rotation'] + Ry),
                            'phi': (['Z rotation'] + Rz),
                            'empty3': np.nan,
                            'J': (['Current density'] + J)}, columns=['x','y','z','empty1','l','b','h','empty2','pye','tta','phi','empty3','J'] )
        
        Out = Out.rename(columns={'empty1':'', 'empty2':'','empty3':''})
        Out = Out.transpose()
        
        writer = pd.ExcelWriter('Builder.xlsx', engine='xlsxwriter')
        Out.to_excel(writer, index = False, sheet_name = 'Slabs')
        writer.save()

#%% Inputs

Total_Coils = 6
TurnsInCoil = 7
I_in_1_turn = 10
height      = 370e-3
width       = 130e-3
thickness   = 25e-3
x_distance  = 36.5e-3
width_leg1  = 28e-3
width_leg2  = 28e-3
width_leg3  = 28e-3
width_leg4  = 28e-3

Make_Coils = Coil(Total_Coils, TurnsInCoil, I_in_1_turn, height, width, thickness, x_distance, width_leg1, width_leg2, width_leg3, width_leg4)
Make_Coils.Write_data()


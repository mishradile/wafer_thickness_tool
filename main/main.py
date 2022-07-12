import openpyxl 
from scipy.interpolate import RBFInterpolator
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from copy import copy
from statistics import stdev
import os
import re
from matplotlib.patches import Circle

wb = openpyxl.load_workbook("./data.xlsx")
#Create image directory to store produced images
if not os.path.isdir("./images/"):
    os.makedirs("./images/") 
else:
    for f in os.listdir("./images/"):
        os.remove(os.path.join("./images/", f))   

for ws in wb.worksheets:
    print("Working on "+str(ws)+" ...")
    #Find the number of variables to be plotted - Number of columns minus the x, y coordinates columns
    var_num = int(ws.max_column-2)

    # Convert to pandas DataFrame
    df = pd.DataFrame(ws.values)
   
    points = df.iloc[:,-2:]
    points = points.to_numpy()
    print(points.shape)
    print(type(points))
    print(points)
    
    #Circle for cropping image, take radius to be furthest distance of any point from origin, to make sure circle fits for all data sets
    radius_list = (points[:,0]**2 + points[:,1]**2)**(0.5)
    circle = Circle((0, 0), max(radius_list)+5, facecolor='none',
             edgecolor=(0, 0, 0), linewidth=2, alpha=0.5)

    
    #Loop through each column
    for col_no in range(var_num):
        values = df.iloc[:,col_no]
        values = values.to_numpy()
        print(values.shape)
        print(type(values))
        
        fig, ax = plt.subplots()
        
        xgrid = np.mgrid[-max(radius_list)-10: max(radius_list)+10:200j, -max(radius_list)-10: max(radius_list)+10:200j]
        xflat = xgrid.reshape(2, -1).T
        #possible: gaussian with epsilon 0.01
        yflat = RBFInterpolator(points, values, kernel='linear')(xflat)
        ygrid = yflat.reshape(200, 200)
        
        im = ax.pcolormesh(*xgrid, ygrid, shading='gouraud')
        p = ax.scatter(*points.T, c=values, s=30, ec='k',)
        fig.colorbar(p)
        
        #Adding text to plot
        avg = "{:.2f}".format(np.mean(values, axis=0))
        rng = "{:.2f}".format(max(values)-min(values))
        onesig = "{:.2f}".format(100*stdev(values)/np.mean(values, axis=0))

        avg_string = 'Average: '+str(avg)
        rng_string = '±Range: '+str(rng)
        onesig_string = '1Sig: '+ str(onesig)+"%"
        ax.text(-150, -170, avg_string, fontsize=10, fontweight = 'bold')
        ax.text(-30, -170, rng_string, fontsize=10, fontweight = 'bold')
        ax.text(90, -170, onesig_string, fontsize=10, fontweight = 'bold')
        new_circle = copy(circle)
        ax.add_patch(new_circle)
        im.set_clip_path(new_circle)
        #Display height above each data point
        points_1 = df.iloc[:,-2:]
        for index, row in points_1.iterrows():
            ax.text(row.iloc[0], row.iloc[1], str( "{:.1f}".format(values[index])), fontsize=10)
        
        
        plt.title('')
        
        fig.tight_layout()
        plt.axis('off')
        image_path = "./images/"+ str(re.search(r'.*?\"(.*)".*' , str(ws)).group(1))+"_"+str(col_no+1)
        plt.show()
        plt.savefig(image_path, bbox_inches='tight', pad_inches=0)
        plt.close()

print("Images generated successfully. View results in /images/ folder.")
        


        
        
        
        
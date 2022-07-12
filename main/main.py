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
    # print(points.shape)
    # print(type(points))
    # print(points)
    
    #Circle for cropping image, take radius to be 150 (+4 for aesthetics so that data points won't be clipped)
    circle = Circle((0, 0), 154, facecolor='none',
             edgecolor=(0, 0, 0), linewidth=1, alpha=1)

    
    #Loop through each column
    for col_no in range(var_num):
        values = df.iloc[:,col_no]
        values = values.to_numpy()
        # print(values.shape)
        # print(type(values))
        
        fig, ax = plt.subplots()
        
        xgrid = np.mgrid[-155: 155:200j, -155: 155:200j]
        xflat = xgrid.reshape(2, -1).T
        #possible: gaussian with epsilon 0.01
        yflat = RBFInterpolator(points, values, kernel='linear')(xflat)
        ygrid = yflat.reshape(200, 200)
        
        im = ax.pcolormesh(*xgrid, ygrid, shading='gouraud', cmap='jet')
        # Use line below for point plot to see true value vs interpolated value
        #p = ax.scatter(*points.T, c=values, s=30, ec='k', cmap='jet')
        p = ax.scatter(*points.T, s=10, ec='k', c='black')
        fig.colorbar(im)
        
        #Adding text to plot
        avg = "{:.2f}".format(np.mean(values, axis=0))
        rng = "{:.2f}".format(max(values)-min(values))
        onesig = "{:.2f}".format(100*stdev(values)/np.mean(values, axis=0))

        avg_string = 'Average: '+str(avg)
        rng_string = '±Range: '+str(rng)
        onesig_string = '1Sig: '+ str(onesig)+"%"
        ax.text(-150, -180, avg_string, fontsize=10, fontweight = 'bold')
        ax.text(-30, -180, rng_string, fontsize=10, fontweight = 'bold')
        ax.text(90, -180, onesig_string, fontsize=10, fontweight = 'bold')
        new_circle = copy(circle)
        ax.add_patch(new_circle)
        im.set_clip_path(new_circle)
        #Display value above each data point
        points_1 = df.iloc[:,-2:]
        for index, row in points_1.iterrows():
            ax.text(row.iloc[0], row.iloc[1]+2, str( "{:.1f}".format(values[index])), fontsize=5)
        
        
        plt.title('')
        
        fig.tight_layout()
        plt.axis('off')
        image_path = "./images/"+ str(re.search(r'.*?\"(.*)".*' , str(ws)).group(1))+"_"+str(col_no+1)
        
        plt.savefig(image_path)
        #plt.show()
        plt.close()

print("Images generated successfully. View results in /images/ folder.")
        


        
        
        
        
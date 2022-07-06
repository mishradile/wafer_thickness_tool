import openpyxl 
from scipy.interpolate import griddata
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from copy import copy
from statistics import stdev
import os
import re
from matplotlib.patches import Circle
circle = Circle((0, 0), 150, facecolor='none',
             edgecolor=(0, 0, 0), linewidth=2, alpha=0.5)


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
    
    #Loop through each column
    for col_no in range(var_num):
        values = df.iloc[:,col_no]

        
        #https://docs.scipy.org/doc/scipy/tutorial/interpolate.html
        grid_x, grid_y = np.mgrid[0:1:100j, 0:1:100j]
        
        grid_data= griddata(points, values, (grid_x, grid_y), method='cubic')
        
        fig, ax = plt.subplots(figsize=(6, 6))
        
        #For some reasons need to copy circle object, if not will error
        new_circle = copy(circle)
        #Add patch to crop image into circle
        ax.add_patch(new_circle)
        #Plot data points
        plt.plot(points.iloc[:,0], points.iloc[:,1], 'k.', ms=5)
        
        
        im = plt.imshow(grid_data.T, extent=(min(points.iloc[:,0])-10,max(points.iloc[:,0])+10,min(points.iloc[:,1])-10,max(points.iloc[:,1])+10), origin='lower', cmap = 'jet')
        im.set_clip_path(new_circle)
        #Format colorbar to show values to 2 d.p
        plt.colorbar(format = '%1.1f', shrink =0.75)
        
        #Adding text to plot
        avg = "{:.2f}".format(np.mean(values, axis=0))
        rng = "{:.2f}".format(max(values)-min(values))
        onesig = "{:.2f}".format(100*stdev(values)/np.mean(values, axis=0))

        avg_string = 'Average: '+str(avg)
        rng_string = '±Range: '+str(rng)
        onesig_string = '1Sig: '+ str(onesig)+"%"
        ax.text(-150, -160, avg_string, fontsize=10, fontweight = 'bold')
        ax.text(-30, -160, rng_string, fontsize=10, fontweight = 'bold')
        ax.text(90, -160, onesig_string, fontsize=10, fontweight = 'bold')
        #Display height above each data point
        for index, row in points.iterrows():
            ax.text(row.iloc[0], row.iloc[1], str( "{:.1f}".format(values[index])), fontsize=5)
        
        
        plt.title('')
        
        fig.tight_layout()
        plt.axis('off')
        image_path = "./images/"+ str(re.search(r'.*?\"(.*)".*' , str(ws)).group(1))+"_"+str(col_no+1)
        plt.savefig(image_path, bbox_inches='tight', pad_inches=0)
        plt.close()

print("Images generated successfully. View results in /images/ folder.")
        


        
        
        
        
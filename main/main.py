import openpyxl 
from matplotlib.mlab import griddata
import scipy.interpolate as interp
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
    
    #Circle for cropping image, take radius to be furthest distance of any point from origin, to make sure circle fits for all data sets
    radius_list = (points.iloc[:,0]**2 + points.iloc[:,1]**2)**(0.5)
    circle = Circle((0, 0), max(radius_list)+5, facecolor='none',
             edgecolor=(0, 0, 0), linewidth=2, alpha=0.5)

    
    #Loop through each column
    for col_no in range(var_num):
        z = df.iloc[:,col_no]
        values = z

        x = points.iloc[:, 0]
        y = points.iloc[:, 1]
        # define grid.
        xi = np.linspace(-20, 20, 21)
        yi = np.linspace(-59, 59, 21)
        # grid the data.
        zi = griddata(x, y, z, xi, yi, interp='linear')

        #from http:// stackoverflow.com/questions/34489039/ retrieving-data-points-from-scipy-interpolate-griddata
        xi_coords = {value: index for index, value in enumerate(xi)}
        yi_coords = {value: index for index, value in enumerate(yi)}


        #iterate to find all the griddata z-height values
        zii = []

        for index, value in enumerate(xi):
            for index, value2 in enumerate(yi):
                zii.append(zi[xi_coords[value],yi_coords[value2]])

        #RBF Method
        xi,yi=np.meshgrid(xi, yi)
        RBFi = interp.Rbf(xi, yi, zii, function='quintic', smooth=0)
        # re-grid the data to fit the entire graph
        xi = np.linspace(-25, 25, 151)
        yi = np.linspace(-65, 65, 151)
        xi,yi=np.meshgrid(xi, yi)
        zi = RBFi(xi, yi)



        im = plt.figure(num=None, figsize=(9.95, 16.712), dpi=80, facecolor='w', edgecolor='k')
        # contour the gridded data, plotting dots at the nonuniform data points.
        CS = plt.contour(xi, yi, zi, 30, linewidths=0.5, colors='k')
        CS = plt.contourf(xi, yi, zi, 50, cmap=plt.cm.rainbow)
        
        fig, ax = plt.subplots(figsize=(6, 6))
        
        #For some reasons need to copy circle object, if not will error
        new_circle = copy(circle)
        #Add patch to crop image into circle
        ax.add_patch(new_circle)
        #Plot data points
        plt.plot(points.iloc[:,0], points.iloc[:,1], 'k.', ms=5)
        
        
        im.set_clip_path(new_circle)
        #Format colorbar to show values to 2 d.p
        plt.colorbar(format = '%1.0f', shrink =0.75)
        
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
        #Display height above each data point
        for index, row in points.iterrows():
            ax.text(row.iloc[0], row.iloc[1], str( "{:.1f}".format(values[index])), fontsize=5)
        
        
        plt.title('')
        
        fig.tight_layout()
        plt.axis('off')
        image_path = "./images/"+ str(re.search(r'.*?\"(.*)".*' , str(ws)).group(1))+"_"+str(col_no+1)
        plt.show()
        plt.savefig(image_path, bbox_inches='tight', pad_inches=0)
        plt.close()

print("Images generated successfully. View results in /images/ folder.")
        


        
        
        
        
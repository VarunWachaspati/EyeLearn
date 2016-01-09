import numpy as np
import pandas as pd
import os
import sys

root_loc = "/home/varunwachaspati/LVPEI/" 
csv_loc = "/home/varunwachaspati/LVPEI/csv/"
res_loc = "/home/varunwachaspati/LVPEI/result/"

def generate():
    os.chdir(csv_loc)
    for fil in os.listdir(csv_loc):
        dat = pd.read_csv(fil,low_memory=False)
        dat.sort_values(['uid','age'],ascending = True, kind='quicksort',inplace=True)
        dat.reset_index(drop=True,inplace=True)
        dat.to_csv(res_loc+fil,index=False,columns=['uid','visit_date','age','gender','od_sph','od_cyl','od_axis','od_ucva','od_bcva','os_sph','os_cyl','os_axis','os_ucva','os_bcva'])

def segregate():
    os.chdir(res_loc)
    for fil in os.listdir(res_loc):
        dat = pd.read_csv(fil,low_memory=False)
        d1 = dat.groupby(["uid"]).mean()
        lis1 = d1[d1["age"]>=10].index.tolist()
        lis2 = d1[d1["age"]<10].index.tolist()
        a1 = dat[dat["uid"].isin(lis1)]
        a2 = dat[dat["uid"].isin(lis2)]
        a1.to_csv(res_loc+fil[:-4]+"_above_10.csv",index=False,columns=['uid','visit_date','age','gender','od_sph','od_cyl','od_axis','od_ucva','od_bcva','os_sph','os_cyl','os_axis','os_ucva','os_bcva'])
        a2.to_csv(res_loc+fil[:-4]+"_below_10.csv",index=False,columns=['uid','visit_date','age','gender','od_sph','od_cyl','od_axis','od_ucva','od_bcva','os_sph','os_cyl','os_axis','os_ucva','os_bcva'])

def main():
    generate()
    segregate()
if __name__ == '__main__':
    main()
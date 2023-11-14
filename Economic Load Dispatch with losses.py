import sys as ss
import numpy as np
import pandas as pd
np.set_printoptions(precision=6, suppress=True)

#%% Accesing data provided from excel file
#NOTE: Update excel file path brfore executing the code

unitdata=pd.read_excel(r"C:\   FILEPATH   \ELDdata.xlsx", sheet_name="fuel")
#print(unitdata)
a = unitdata[['Alpha']].to_numpy() * unitdata[['Rate']].to_numpy()
b = unitdata[['Beta']].to_numpy() * unitdata[['Rate']].to_numpy()
c = unitdata[['Gamma']].to_numpy() * unitdata[['Rate']].to_numpy()
Pmax = unitdata[['Max Limit']].to_numpy()
Pmin = unitdata[['Min Limit']].to_numpy()

lossdata=pd.read_excel(r"C:\   FILEPATH   \ELDdata.xlsx", sheet_name="loss")
B=lossdata.to_numpy()
#print(B)

Lambda=max(b) #Initialising lambda
Load=int(input("Enter the load(in MW):")) #Taking load as input
if(Load>=sum(Pmax)):
    print("Demand exceeds generation capacity")
    ss.exit()
elif(Load<=sum(Pmin)):
    print("Load below minimum allowed limits")
    ss.exit()    

delP=1 #Initialising delP

#%% Obtaining starting values without consideing loss

while abs(delP) > 0.0001:

    P = (Lambda-b)/(2*c)
    P = np.maximum(P, Pmin)
    P = np.minimum(P, Pmax)
    delP = Load - np.sum(P)
    Lambda = Lambda + 2 * delP/np.sum(1/c) #Lambda correction

    
#%% Considering loss and limits
iteration=0
maxiter=20

while iteration<maxiter:
    iteration+=1
    Ploss=np.dot(np.dot(np.transpose(P),B),P)
    Pdemand=Load+Ploss
    Penaltyfactor=1/(np.ones((len(a), 1))-2*np.dot(B,P))
    L=Lambda
    delP=1
    while abs(delP) > 0.0001:        
        Pnew = ((L/Penaltyfactor)-b)/(2*c)
        Pnew = np.maximum(Pnew, Pmin)
        Pnew = np.minimum(Pnew, Pmax)
        delP = Pdemand - np.sum(Pnew)
        L = L + 2 * delP/np.sum((1/c)) #Lambda correction

    if max(abs(Pnew-P))<0.0001:
        break
    else:
        P=Pnew.copy()
        
if iteration==maxiter:
    print("Not converged")
else:
    print("Converged in",iteration,"iterations")

#%% Final results

cost = a + b * P + c * P * P 
totalcost = np.sum(cost) 

for i in range(len(a)):
    print(f"The unit {i + 1} generates power {P[i, 0]:.4f} MW, costing ${cost[i, 0]:.4f}/h")
    
print(f"\nTotal cost ${totalcost:.4f}/h")        
print("Total generation =",np.round(sum(Pnew)[0],4),"MW")
print("Total Losses =",np.round(sum(Pnew)[0]-Load,4),"MW")

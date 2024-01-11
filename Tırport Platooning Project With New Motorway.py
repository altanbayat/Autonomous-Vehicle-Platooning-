#!/usr/bin/env python
# coding: utf-8

# In[4]:


import numpy as np
import pandas as pd
import gurobipy as gp
from gurobipy import GRB, quicksum

def modeltry(file):
    # Read data from the Excel file
    cost = pd.read_excel(file, sheet_name=0)  # Reading cost data
    flow = pd.read_excel(file, sheet_name=4)  # Reading flow data
    flow2 = pd.read_excel(file, sheet_name=4)  # Reading another copy of flow data for direct routes
    cost_interhub = pd.read_excel(file, sheet_name=6)  # Reading inter-hub cost data

    # Initialize sets for locations and hubs
    I = np.arange(1, 82).tolist()  # Set of origin locations
    J = np.arange(1, 82).tolist()  # Set of destination locations
    K = [1, 4, 5, 6, 9, 10, 14, 16, 17, 18, 19, 20, 22, 24, 25, 27, 33, 34, 35, 39, 40, 41, 45, 51, 54, 58, 59, 60, 68, 77, 80, 81]  # Potential hub locations
    L = K.copy()  # Duplicate set for hub-to-hub routes
    Q = 10  # Desired number of hubs
    alpha = 0.744  # Cost factor for inter-hub routes

    # Define possible routes and assignments
    Routes = [(i, j, k, l) for i in I for j in J for k in K for l in L]  # All possible routes using hubs
    DirectAssignment = [(i, j) for i in I for j in J]  # All possible direct routes

    # Create dictionaries for flow and cost data
    flow = dict(((i, j, k, l), flow[j][i]) for i in I for j in J for k in K for l in L)  # Flow data for hub-used routes
    fuelcost = dict(((i, j), cost[j][i]) for i in I for j in J)  # Fuel cost for direct routes
    flowdirect = dict(((i, j), flow2[j][i]) for i in I for j in J)  # Flow data for direct routes
    fuelcostinterhub = dict(((k, l), cost[l][k]) for k in K for l in L)  # Fuel cost for inter-hub routes

    # Calculate costs for routes
    HubRouteCost = dict(((i, j, k, l), (flow[i, j, k, l]) * (fuelcost[i, k] + alpha * fuelcostinterhub[k, l] + fuelcost[l, j]))
                        for i in I for j in J for k in K for l in L)  # Cost for hub-used routes
    DirectRouteCost = dict(((i, j), (flowdirect[i, j]) * (fuelcost[i, j]))
                           for i in I for j in J)  # Cost for direct routes

    # Initialize and configure the optimization model
    m = gp.Model("cfl")
    m.ModelSense = GRB.MINIMIZE

    # Add variables to the model
    xk = m.addVars(K, name="City Nodes", vtype=GRB.BINARY, obj=0)  # Binary variables for selecting city nodes as hubs
    yijkl = m.addVars(Routes, name="Hub Used Routes", obj=HubRouteCost)  # Variables for routes using hubs
    yij = m.addVars(DirectAssignment, name="Direct Routes", obj=DirectRouteCost)  # Variables for direct routes

    # Add constraints to the model
    desiredhub = m.addConstr((quicksum(xk[k] for k in K) <= Q), "Desired Hub")  # Limit the number of hubs
    directorhub = m.addConstrs((yij[i, j] + quicksum(yijkl[i, j, k, l] for k in K for l in L) == 1
                               for i in I for j in J), "Direct or Hub Decision Constraint")  # Ensure each route is either direct or uses a hub
    hub1 = m.addConstrs((quicksum(yijkl[i, j, k, l] for l in L) <= xk[k] for i in I for j in J for k in K), "Hub1")  # Constraint for hub 1
    hub2 = m.addConstrs((quicksum(yijkl[i, j, k, l] for k in K) <= xk[l] for i in I for j in J for l in L), "Hub2")  # Constraint for hub 2

    # Run the optimization
    m.optimize()

    # Print the optimal solution's details
    for v in m.getVars():
        if v.x > 0:
            print(f"{v.varName}: {v.x}")

    # Return the objective value of the model
    return m.objVal

# Call the function with the path to the Excel file
modeltry("temporary data (1).xls")
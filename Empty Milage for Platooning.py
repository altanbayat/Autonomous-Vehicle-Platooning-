#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import numpy as np
import gurobipy as gp
from gurobipy import GRB, quicksum

def modeltry(file, network_file, output_file="output_flow_info.xlsx", new_output_file="new_output_file.xlsx"):
    # Read input data from Excel files: cost, flow, duplicate flow, and inter-hub cost data
    cost = pd.read_excel(file, sheet_name=0)  # Reading cost data
    flow = pd.read_excel(file, sheet_name=4)  # Reading flow data
    flow2 = pd.read_excel(file, sheet_name=4)  # Reading another copy of flow data
    cost_interhub = pd.read_excel(file, sheet_name=2)  # Reading inter-hub cost data

    # Define sets and parameters for the optimization model
    I = np.arange(1, 82).tolist()  # Range of indices, possibly representing locations
    J = np.arange(1, 82).tolist()  # Similar range as I
    K = [1, 6, 9, 10, 14, 16, 17, 20, 22, 27, 33, 34, 35, 39, 40, 41, 45, 51, 54, 59, 68, 77, 80, 81]  # Specific set of indices
    L = [1, 6, 9, 10, 14, 16, 17, 20, 22, 27, 33, 34, 35, 39, 40, 41, 45, 51, 54, 59, 68, 77, 80, 81]  # Similar set as K
    Q = 10  # Desired number of hubs
    alpha = 0.610  # A parameter, possibly a cost factor

    # Define Routes and DirectAssignment sets for the model
    Routes = [(i, j, k, l) for i in I for j in J for k in K for l in L]  # Possible routes using hubs
    DirectAssignment = [(i, j) for i in I for j in J]  # Direct routes between locations

    # Create dictionaries for flows, costs, and direct flows
    flow = dict(((i, j, k, l), flow[j][i]) for i in I for j in J for k in K for l in L)  # Flow data structured for model
    fuelcost = dict(((i, j), cost[j][i]) for i in I for j in J)  # Fuel cost for direct routes
    flowdirect = dict(((i, j), flow2[j][i]) for i in I for j in J)  # Duplicate flow data for direct routes
    fuelcostinterhub = dict(((k, l), cost[l][k]) for k in K for l in L)  # Fuel cost between hubs

    # Define cost dictionaries for hub routes and direct routes
    HubRouteCost = dict(((i, j, k, l), (flow[i, j, k, l]) * (fuelcost[i, k] + alpha * fuelcostinterhub[k, l] + fuelcost[l, j]))
                        for i in I for j in J for k in K for l in L)  # Cost calculation for routes using hubs
    DirectRouteCost = dict(((i, j), (flowdirect[i, j]) * (fuelcost[i, j]))
                          for i in I for j in J)  # Cost calculation for direct routes

    # Create Gurobi model for optimization
    m = gp.Model("cfl")
    m.ModelSense = GRB.MINIMIZE  # Setting the model to minimize the objective function

    # Add decision variables to the model
    xk = m.addVars(K, name="City_Nodes", vtype=GRB.BINARY, obj=0)  # Binary variable indicating if a city is chosen as a hub
    yijkl = m.addVars(Routes, name="Hub_Used_Routes", obj=HubRouteCost)  # Variable for routes using hubs
    yij = m.addVars(DirectAssignment, name="Direct_Routes", obj=DirectRouteCost)  # Variable for direct routes

    # Add constraints to the model
    desiredhub = m.addConstr((quicksum(xk[k] for k in K) <= Q), "Desired_Hub")  # Constraint for the number of hubs
    directorhub = m.addConstrs((yij[i, j] + quicksum(yijkl[i, j, k, l] for k in K for l in L) == 1
                                for i in I for j in J), "Direct_or_Hub_Decision_Constraint")  # Ensure each route is either direct or uses a hub
    hub1 = m.addConstrs((quicksum(yijkl[i, j, k, l] for l in L) <= xk[k] for i in I for j in J for k in K), "Hub1")  # Hub usage constraint
    hub2 = m.addConstrs((quicksum(yijkl[i, j, k, l] for k in K) <= xk[l] for i in I for j in J for l in L), "Hub2")  # Another hub usage constraint

    m.optimize()  # Optimize the model

    # Create a list to store flow information for further analysis
    flow_info_list = []

    # Iterate through all variables in the optimized model to extract relevant data
    for v in m.getVars():
        if v.x > 0:  # Consider only variables that are part of the optimal solution
            # Specifically check for 'Hub_Used_Routes' variables with a value of 1
            if v.varName.startswith("Hub_Used_Routes") and v.x == 1:
                # Extract indices (i, j, k, l) from the variable name
                i, j, k, l = map(int, v.varName.split("[")[1].split("]")[0].split(","))
                # Append the route information along with its flow value to the list
                flow_info_list.append([i, j, k, l, flow[i, j, k, l]])

    # Convert the list into a DataFrame for better data handling and visualization
    flow_info_df = pd.DataFrame(flow_info_list, columns=['i', 'j', 'k', 'l', 'flow_k1_l6'])

    # Read additional network data from another Excel file
    network_data = pd.read_excel(network_file, sheet_name=3, index_col=0)  # Adjust the sheet number as required

    # Create a dictionary from the network data for quick access
    network_dict = {(i, k): network_data.loc[i, k] for i in I for k in K}

    # Add a new column to the DataFrame by mapping the network data to the flow information
    flow_info_df['Turkish network (6)'] = flow_info_df.apply(lambda row: network_dict[(row['i'], row['k'])], axis=1)

    # Round the values in the 'flow_k1_l6' column and convert them to integers
    flow_info_df['flow_k1_l6'] = np.ceil(flow_info_df['flow_k1_l6']).astype(int)

    # Calculate a new column 'result_column' as the product of 'flow_k1_l6' and 'Turkish network (6)'
    flow_info_df['result_column'] = flow_info_df['flow_k1_l6'] * flow_info_df['Turkish network (6)']
    
    # Sum the values in the 'result_column' to get a total
    total_sum = flow_info_df['result_column'].sum()
    print("Toplam:", total_sum)  # Print the total sum for verification

    # Optionally, add the total sum to the DataFrame as a new column
    flow_info_df['total_sum'] = total_sum

    # Save the processed and enriched data to a new Excel file
    flow_info_df.to_excel(output_file, index=False)
    
    return m.objVal  # Return the objective value of the optimization model

# Call the function with specified file paths
result = modeltry("temporary data (1).xls", "Turkish network (6).xls", output_file="output_flow_info.xlsx", new_output_file="new_output_file.xlsx")
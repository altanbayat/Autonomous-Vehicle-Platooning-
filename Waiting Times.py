#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import gurobipy as gp
from gurobipy import GRB, quicksum

def modeltry(file, network_file, output_file="output_flow_info.xlsx", new_output_file="new_output_file.xlsx"):
    # Read data from Excel files
    cost = pd.read_excel(file, sheet_name=0)  # Cost data
    flow = pd.read_excel(file, sheet_name=4)  # Flow data
    flow2 = pd.read_excel(file, sheet_name=4)  # Duplicate flow data for direct routes
    cost_interhub = pd.read_excel(file, sheet_name=2)  # Inter-hub cost data

    # Define sets for locations and hubs
    I = np.arange(1, 82).tolist()  # Set of origin locations
    J = np.arange(1, 82).tolist()  # Set of destination locations
    K = [1, 6, 9, 10, 14, 16, 17, 20, 22, 27, 33, 34, 35, 39, 40, 41, 45, 51, 54, 59, 68, 77, 80, 81]  # Potential hub locations
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
    xk = m.addVars(K, name="City_Nodes", vtype=GRB.BINARY, obj=0)  # Binary variables for selecting city nodes as hubs
    yijkl = m.addVars(Routes, name="Hub_Used_Routes", obj=HubRouteCost)  # Variables for routes using hubs
    yij = m.addVars(DirectAssignment, name="Direct_Routes", obj=DirectRouteCost)  # Variables for direct routes

    # Add constraints to the model
    desiredhub = m.addConstr((quicksum(xk[k] for k in K) <= Q), "Desired_Hub")  # Limit the number of hubs
    directorhub = m.addConstrs((yij[i, j] + quicksum(yijkl[i, j, k, l] for k in K for l in L) == 1
                                for i in I for j in J), "Direct_or_Hub_Decision_Constraint")  # Ensure each route is either direct or uses a hub
    hub1 = m.addConstrs((quicksum(yijkl[i, j, k, l] for l in L) <= xk[k] for i in I for j in J for k in K), "Hub1")  # Constraint for hub 1
    hub2 = m.addConstrs((quicksum(yijkl[i, j, k, l] for k in K) <= xk[l] for i in I for j in J for l in L), "Hub2")  # Constraint for hub 2

    m.optimize()  # Execute optimization

    # Extract and process data from the optimized model
    flow_info_list = []
    for v in m.getVars():
        if v.x > 0:
            if v.varName.startswith("Hub_Used_Routes") and v.x == 1:
                i, j, k, l = map(int, v.varName.split("[")[1].split("]")[0].split(","))
                if k == 14 and l == 40:  # Specific condition for certain routes
                    flow_info_list.append([i, j, k, l, flow[i, j, k, l]])

    flow_info_df = pd.DataFrame(flow_info_list, columns=['i', 'j', 'k', 'l', 'flow_k1_l6'])  # Convert list to DataFrame
    
    # Read additional network data from an Excel file
    network_data = pd.read_excel(network_file, sheet_name=3, index_col=0)  # Adjust sheet number if needed

    # Create a dictionary mapping for network data
    network_dict = {(i, k): network_data.loc[i, k] for i in I for k in K}

    # Add a new column to the DataFrame using the network data
    flow_info_df['Turkish network (6)'] = flow_info_df.apply(lambda row: network_dict[(row['i'], row['k'])], axis=1)

    # Round the values in the DataFrame
    flow_info_df['flow_k1_l6'] = np.ceil(flow_info_df['flow_k1_l6']).astype(int)
    flow_info_df['Turkish network (6)'] = np.round(flow_info_df['Turkish network (6)']).astype(int)

    # Process the data to create a new column (7th column)
    column7_df = pd.DataFrame(columns=['column7'])
    for _, row in flow_info_df.iterrows():
        column7_df = pd.concat([column7_df, pd.DataFrame({'column7': [row['Turkish network (6)']] * row['flow_k1_l6']})], ignore_index=True)

    # Merge the new column with the original DataFrame
    flow_info_df = pd.concat([flow_info_df, column7_df], axis=1)

    # Sort and group the values in the 7th column
    flow_info_df.sort_values(by='column7', inplace=True)
    grouped_values = [list(g['column7']) for _, g in flow_info_df.groupby(np.arange(len(flow_info_df)) // 4)]

    # Calculate differences in each group and find the average difference
    differences = [max(group) - min(group) for group in grouped_values if len(group) == 4]
    avg_difference = np.mean(differences)

    # Add the average difference to the DataFrame
    flow_info_df['avg_difference'] = avg_difference

    # Export the DataFrame to an Excel file
    flow_info_df.to_excel(output_file, index=False)

    # Generate random integers with triangular distribution and fill the 2nd column
    np.random.seed(42)  # For reproducibility, you can remove this line if not needed
    new_output_df['2nd_column'] = np.random.triangular(left=240, mode=360, right=600, size=len(new_output_df))

    # Update the 2nd column to 0 if the value in the 1st column is 0
    new_output_df.loc[new_output_df['new_column7'] == 0, '2nd_column'] = 0

    # Sum the two columns and add the result to a new column
    new_output_df['sum_of_columns'] = new_output_df['new_column7'] + new_output_df['2nd_column']
    
    # Check and update sum_of_columns values until all values are between 600 and 960  
    new_output_df.sort_values(by='sum_of_columns', inplace=True)

    # Save the modified DataFrame to the new Excel file
    new_output_df.to_excel(new_output_file, index=False)
    grouped_df = new_output_df.groupby(np.arange(len(new_output_df)) // 4)

    # Calculate the differences for each group
    differences = []
    for _, group in grouped_df:
        min_value = group['sum_of_columns'].min()
        max_value = group['sum_of_columns'].max()
        difference = max_value - min_value
        differences.append(difference)

    # Calculate the average, minimum, and maximum differences
    avg_difference = np.mean(differences)
    min_difference = min(differences)
    max_difference = max(differences)

    # Print or display the calculated values
    print("Average Difference:", avg_difference)
    print("Minimum Difference:", min_difference)
    print("Maximum Difference:", max_difference)    
    return m.objVal

# Call the function with your file path
result = modeltry("temporary data (1).xls", "Turkish network (6).xls", output_file="output_flow_info.xlsx", new_output_file="new_output_file.xlsx")

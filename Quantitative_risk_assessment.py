# -*- coding: utf-8 -*-
"""
Created on Sun May 28 00:07:01 2023

@author: riadh
"""

import math
import xlsxwriter

# Read asset values from a text file
with open("asset_values.txt", "r") as f:
    asset_values = {}
    for line in f:
        parts = line.strip().split(",")
        asset = parts[0]
        value = int(parts[1])
        asset_values[asset] = value

# Read threat levels from a text file
with open("threat_levels.txt", "r") as f:
    threat_levels = {}
    for line in f:
        parts = line.strip().split(",")
        threat = parts[0]
        level = float(parts[1])
        threat_levels[threat] = level

# Define a function to calculate annualized loss expectancy (ALE)
def calculate_ale(asset_value, threat_level, vulnerability):
    s = 1 - vulnerability
    ale = asset_value * threat_level * s
    return round(ale, 2)

# Define a function to interpret the ALE
def interpret_ale(ale):
    ale_level = ""
    if ale >= 100000:
        ale_level = "very high"
    elif ale >= 50000:
        ale_level = "high"
    elif ale >= 10000:
        ale_level = "moderate"
    elif ale >= 1000:
        ale_level = "low"
    else:
        ale_level = "very low"
    return ale_level

# Define a matrix of ALE values for each asset-threat pair
ale_matrix = {}
for asset in asset_values:
    ale_matrix[asset] = {}
    asset_value = asset_values[asset]
    for threat in threat_levels:
        threat_level = threat_levels[threat]
        vulnerability = 0.25 # You can replace this with your own value
        ale = calculate_ale(asset_value, threat_level, vulnerability)
        ale_level = interpret_ale(ale)
        ale_matrix[asset][threat] = ale_level

# Save the ALE matrix to an Excel file
workbook = xlsxwriter.Workbook('ale_matrix.xlsx')
worksheet = workbook.add_worksheet()

# Write the asset names as column headers
col = 0
for asset in ale_matrix:
    worksheet.write(0, col + 1, asset)
    col += 1

# Write the threat names as row headers, and the ALE values as cells
row = 1
for threat in ale_matrix:
    worksheet.write(row, 0, threat)
    col = 1
    for asset in ale_matrix[threat]:
        worksheet.write(row, col, ale_matrix[threat][asset])
        col += 1
    row += 1

workbook.close()
print("The ALE matrix has been saved to ale_matrix.xlsx.")
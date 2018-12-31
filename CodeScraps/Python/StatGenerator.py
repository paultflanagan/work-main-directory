#Generates a list of 6 stat values using the 4d6-drop-lowest method

import random

StatResults = []
CurrentRolls = []

for stat in range(6):
    CurrentStatSum = 0
    for roll in range(4):
        CurrentRolls.append(random.randint(1,6))
    CurrentRolls.sort()
    CurrentRolls[0] = 0
    for value in range(len(CurrentRolls)):
        CurrentStatSum += CurrentRolls[value-1]
    StatResults.append(CurrentStatSum)
    CurrentRolls = []

print(StatResults)


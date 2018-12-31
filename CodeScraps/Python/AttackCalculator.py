import random

advantage = True
disadvantage = False
AC = 15
toHitBonus = 5
damageDice = [1,8]
damageBonus = 2
resultString = []

rollResult = random.randint(1,20)
if advantage != disadvantage:
    rollResult2 = random.randint(1,20)
    if advantage:
        rollResult = max(rollResult, rollResult2)
        resultString.append("Attack has advantage!")
    if disadvantage:
        rollResult = min(rollResult, rollResult2)
        resultString.append("Attack has disadvantage!")

resultString.append("{rollResult} was rolled!".format(**locals()))

if rollResult == 20:
    resultString.append("Critical strike!")
    damageDice[0] = damageDice[0] * 2
if rollResult >= (AC - toHitBonus) or rollResult == 20:
    resultString.append(str(rollResult + toHitBonus) + " to hit!")
    resultString.append("Attack hit!")
    resultString.append("Rolling {damageDice[0]}d{damageDice[1]} + {damageBonus}...".format(**locals()))
    damage = 0
    for x in range(damageDice[0]):
        damage += random.randint(1, damageDice[1])
    damage += damageBonus
    resultString.append(str(damage) + " damage dealt!")
else:
    resultString.append(str(rollResult + toHitBonus) + " to hit!")
    resultString.append("Attack miss!")

print(resultString)

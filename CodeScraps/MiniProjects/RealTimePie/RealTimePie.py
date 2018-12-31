import matplotlib.pyplot as plt
import subprocess
import TimeCalc as tc


plt.ion()
target = "X:\Automation\Duplicate Check\Results.txt"
labels = 'Pass', 'Fail'
colors = ['green', 'red']
plt.axis('equal')

continuePoll = True

while continuePoll:
    passCount = 0
    failCount = 0
    try:
        Results_FileObject = open(target, "r")
    except FileNotFoundError:
        continue
    EntryList = Results_FileObject.read().split('\n')
    Results_FileObject.close()
    for entry in EntryList:
        if entry != '':
            entrySplit = entry.split('#')
            successState = entrySplit[2]
            if successState == "True":
                passCount += 1
            elif successState == "False":
                failCount += 1
    sizes = [passCount, failCount]
    testStart = (EntryList[0].split('#'))[0].split(' ')
    testStartTime = testStart[1] + ' ' + testStart[2]
    elapsedTest = tc.formatVal(tc.elapsed(tc.process(testStartTime), tc.process(tc.now())))
    lastResult = (EntryList[len(EntryList)-2].split('#'))[1].split(' ')
    lastResultEnd = lastResult[1] + ' ' + lastResult[2]
    elapsedLast = tc.formatVal(tc.elapsed(tc.process(lastResultEnd), tc.process(tc.now())))
    totalResults = passCount + failCount

    plt.clf()
    plt.pie(sizes, labels=labels, colors=colors)
    plt.annotate("Pass: {}".format(passCount), (-1.2, 1))
    plt.annotate("Fail: {}".format(failCount), (1, 1))
    plt.annotate("Total Results: {}".format(totalResults), (-.75, .5))
    plt.annotate("Test start:    {}".format(testStartTime), (-.75, .25))
    plt.annotate("Time since start: {}".format(elapsedTest), (-.75, 0))
    plt.annotate("Last result:   {}".format(lastResultEnd), (-.75, -.25))
    plt.annotate("Time since last result: {}".format(elapsedLast), (-.75, -.5))
    plt.pause(1)
    plt.show()
    #continuePoll = False
        

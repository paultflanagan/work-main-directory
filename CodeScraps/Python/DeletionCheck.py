#Testing functionality of os.mkdir/.rmdir, specifically interactions with nonexistent or previously existent paths
#Prereq: empty folder exists at string targetFolder
import os
import time
import shutil

#Directory to be interacted with
pathname = "C:/Users/paul.flanagan/MainDirectory/CodeScraps/Python/TestEnv"
targetFolder = pathname + "/DeletionTest"

#Deleting currently existent directory
os.rmdir(targetFolder)
time.sleep(3)

#Attempting to delete currently non existent directory
#os.rmdir(targetFolder)
#time.sleep(3)
#Conclusion: this generates an error.

#Creating currently non existent directory
os.mkdir(targetFolder)
time.sleep(3)

#Attempting to create currently existent directory
#os.mkdir(targetFolder)
#time.sleep(3)
#Conclusion: this also generates an error.


#Testing removal of a populated directory

#Populating directory
open(targetFolder + "/ToDelete.txt", mode='x')
time.sleep(3)
#note: this fails if file already exists

#Attempting to deleted populated directory
#os.rmdir(targetFolder)
#time.sleep(3)
#Conclusion: unable to os.rmdir a populated directory.

#Looking around I found a module and command shutil.rmtree, which removes a directory and all contents
#Retroactively importing shutil
shutil.rmtree(targetFolder)
time.sleep(3)

#Attempting to delete non existent tree
#shutil.rmtree(targetFolder)
#time.sleep(3)
#Conclusion: throws an error if attempting to delete non existent tree

#Resetting environment
os.mkdir(targetFolder)

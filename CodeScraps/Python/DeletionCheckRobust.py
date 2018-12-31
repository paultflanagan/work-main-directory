#Creating more robust file creation / deletion code
#Prereq: none
import os
import time
import shutil

#Spent some time trying to find a "does it exist?" function, but then I found try/except/finally

def mk_dir_robust(target):
    try:
        os.mkdir(target)
    except FileExistsError:
        print("Directory to create already exists")
    else:
        print("Directory created")
    finally:
        print("Creation Attempt completed")
    time.sleep(2)

def rm_dir_robust(target):
    try:
        shutil.rmtree(target)
    except FileNotFoundError:
        print("Directory to be deleted does not exist")
    else:
        print("Directory deleted")
    finally:
        print("Deletion Attempt completed")
    time.sleep(2)

pathname = "C:/Users/paul.flanagan/MainDirectory/CodeScraps/Python/TestEnv"
targetFolder = pathname + "/DeletionTest"

#starting with existing folder
mk_dir_robust(targetFolder)

#creating with existing folder
mk_dir_robust(targetFolder)

#deleting with existing folder
rm_dir_robust(targetFolder)

#deleting with non existing folder
rm_dir_robust(targetFolder)

#creating with non existing folder
mk_dir_robust(targetFolder)

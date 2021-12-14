import xlsxwriter
import time
import os
import shutil
import platform
from colorama import init
from colorama import Fore  # Color in Windows and Linux
import signal
import sys
import subprocess
from subprocess import run
from decimal import Decimal

print("Hi")

init()  # Makes sure colour is displayed on Windows. --KEEP AT TOP--


# Check Platform, set file path slashes
if platform.system() == "Linux":
    fileSlashes = "/"
    currentOS = "Linux"
elif platform.system() == "Windows":
    fileSlashes = "\\"
    currentOS = "Windows"
else:
    print(Fore.RED + "Unsupported operating system :(" + Fore.RESET)
    input("Press Enter to exit...")
    sys.exit()
# Check that av1an, ffmpeg, aomenc is in the path
try:
    run("av1an", stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    run("ffmpeg", stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    run("aomenc", stdout=subprocess.PIPE, stderr=subprocess.PIPE)
except:
    print("Missing prerequisite programs in Path")
    #sys.exit()

# WORKING DIRS
workingFileDir = "C:\\Users\\admin\\Videos\\AV1\\Target_Quality\\"  # Where to dump excel files
inputSampleDIR = "C:\\Users\\admin\\Videos\\AV1\\Samples\\"
inputSampleNames = []
inputSampleShortNames = []
outputDir = "C:\\Users\\admin\\Videos\\AV1\\Output\\"
av1anWorkingDIR = "C:\\Users\\admin\\Videos\\AV1\\av1an_working\\"

# Recursively look for all the .mp4 files in the sample folder and dump the paths and names into a list
index1 = 0
for root, dirs, files in os.walk(inputSampleDIR):
    for inputFilename in files:
        if inputFilename.endswith(".mp4"):
            inputSampleNames.append(str(root) + fileSlashes + str(inputFilename))
            inputSampleShortNames.append(str(inputFilename))
            print(str(root) + fileSlashes + str(inputFilename))
            if index1 == 2:
                break  # TO USE ONLY ONE SAMPLE
            index1 += 1


# Setup val's
#xOffset = 5
#yOffset = 1
currentIteration = -1

# what are we testing?
crfValues = [20]
cpuUsedValues = [4, 3, 2]
targetQuality = [72, 74, 76, 80, 84, 88, 92, 94, 96]

# Pre Heat
#print("Running pre heat")
#run("ffmpeg -y -i \"D:\\Footage Dump\\4.mp4\" -c:v libx264 -preset veryslow -crf 0 -an -sn \"D:\\Footage Dump\\tmp.mp4\"", stdout=subprocess.PIPE, stderr=subprocess.PIPE)

for currentInputSampleName in inputSampleNames:  # Iterate through fps + samples

    currentSampleShortName = inputSampleShortNames[currentIteration]

    print("")
    print("Starting: " + str(currentInputSampleName))
    print("")

    currentIteration += 1
    # creates a new excel file for each sample
    workbook = xlsxwriter.Workbook(workingFileDir + str(currentSampleShortName)[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 13)
    worksheet.set_column('B:B', 18)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 10)
    bold = workbook.add_format({'bold': True})

    xOffset = 5
    yOffset = 1
    worksheet.write(xOffset - 1, yOffset + 1, str(currentSampleShortName)[:-4], bold)
    for currentCpuUsed in cpuUsedValues:  # Iterate through CPU-USED

        xOffset += 2
        worksheet.write(xOffset - 1, yOffset - 1, "-CPU-Used " + str(currentCpuUsed), bold)
        worksheet.write(xOffset - 1, yOffset + 1, "Encode Time")
        worksheet.write(xOffset - 1, yOffset + 3, "File Size")
        for currentCRF in crfValues:

            for currentTargetQuality in targetQuality:

                for currentDeviation in range(0, 4):

                    aomencArg = str
                    av1anArg = str
                    extra = str

                    if currentDeviation == 0:
                        aomencArg = "\" -v \" --end-usage=q --cq-level=" + str(currentCRF) + " --cpu-used=" + str(currentCpuUsed) + " --threads=16\""
                        av1anArg = " -w 20 --target-quality " + str(currentTargetQuality)
                        extra = ""
                    elif currentDeviation == 1:
                        aomencArg = "\" -v \" --end-usage=q --cq-level=" + str(currentCRF) + " --lag-in-frames=5 --tile-rows=2 --tile-columns=1 --cpu-used=" + str(currentCpuUsed) + " --threads=16\""
                        av1anArg = " -w 20 --target-quality " + str(currentTargetQuality)
                        extra = "_lag-in-frames"
                    elif currentDeviation == 2:
                        aomencArg = "\" -v \" --end-usage=q --cq-level=30 --cpu-used=" + str(currentCpuUsed) + " --threads=16\""
                        av1anArg = " -w 20 --target-quality " + str(currentTargetQuality)
                        extra = "_CRF30"
                    elif currentDeviation == 3:
                        aomencArg = "\" -v \" --end-usage=q --cq-level=" + str(currentCRF) + " --cpu-used=" + str(currentCpuUsed) + " --threads=8\""
                        av1anArg = " -w 10 --target-quality " + str(currentTargetQuality)
                        extra = "_t8w10"
                    elif currentDeviation == 4:
                        aomencArg = "\" -v \" --end-usage=q --cq-level=" + str(currentCRF) + " --cpu-used=" + str(currentCpuUsed) + " --threads=16\""
                        av1anArg = " -w 20"
                        extra = "_noTarget"


                    outputFileName = str(outputDir) + str(currentSampleShortName)[:-4] + "_" + str(currentCpuUsed) + "_" + str(currentCRF) + "_" + str(currentTargetQuality) + extra + ".mkv"  # [:-4] removed the .mkv extension so i can pu tit on the end
                    outputName = str(currentSampleShortName)[:-4] + "_" + str(currentCpuUsed) + "_" + str(currentCRF) + "_" + str(currentTargetQuality) + extra
                    start_time = time.time()

                    # Printing for debugging
                    print("cd " + av1anWorkingDIR + " && av1an -i \"" + currentInputSampleName + aomencArg + av1anArg + " -o " + outputFileName)
                    # Why do i cd and then run av1an? because evil things happen.
                    #os.system("cd " + av1anWorkingDIR + " && av1an -i \"" + currentInputSampleName + aomencArg + av1anArg + " -o " + outputFileName)



                    # TOFO: Av1an output log and parse for encode duration.

                    processTime = round((time.time() - start_time) / 60, 2)  # Return in Minutes

                    worksheet.write(xOffset, yOffset - 1, " --cq-level=" + str(currentCRF))
                    worksheet.write(xOffset, yOffset, " --target-quality " + str(currentTargetQuality))

                    # Time Taken to excel
                    worksheet.write(xOffset, yOffset + 1, processTime)

                    # File Size
                    try:
                        size = round(os.path.getsize(outputFileName) / 1048576), 2  # IN MEBIBYTES (MiB) with 2 decimal places
                    except:
                        size = -1

                    worksheet.write_number(xOffset, yOffset + 3, size, 2)

                    # txt Backup

                    f = open(workingFileDir + outputName + "_Sizes.txt", "a")
                    f.write(str(time.time()) + " :" + outputName + ":" + str(size) + "\n")
                    f.close()

                    f = open(workingFileDir + outputName + "_Times.txt", "a")
                    f.write(str(time.time()) + " :" + outputName + ":" + str(processTime) + "\n")
                    f.close()

                    xOffset += 1

    workbook.close()

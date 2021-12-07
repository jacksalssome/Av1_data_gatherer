import xlsxwriter
import time
import os
import shutil
import platform
from colorama import init
from colorama import Fore  # Color in Windows and Linux
import sys
import subprocess
from subprocess import run

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
    sys.exit()

# WORKING DIRS
workingFileDir = "C:\\Users\\admin\Videos\\AV1\\Target_Quality\\"  # Where to dump excel files
inputSampleDIR = "C:\\Users\\admin\Videos\\AV1\\Samples\\"
inputSampleNames = []
inputSampleShortNames = []
outputDir = "C:\\Users\\admin\\Videos\\AV1\\Output\\"
av1anWorkingDIR = "C:\\Users\\admin\\Videos\\AV1\\av1an_working\\"
vmafPath = "C:\\Program Files\\ffmpeg\\vmaf_v0.6.1.json"  # Path to vmaf model HAS TO BE .json


# Recursively look for all the .mp4 files in the sample folder and dump the paths and names into a list
for root, dirs, files in os.walk(inputSampleDIR):
    for inputFilename in files:
        if inputFilename.endswith(".mp4"):
            inputSampleNames.append(str(root) + fileSlashes + str(inputFilename))
            inputSampleShortNames.append(str(inputFilename))
            print(str(root) + fileSlashes + str(inputFilename))
            continue  # TO USE ONLY ONE SAMPLE


# Setup val's
#xOffset = 5
#yOffset = 1
currentIteration = -1

# what are we testing?
crfValues = [20]
cpuUsedValues = [4, 3, 2]
targetQuality = [75, 80, 82, 83, 84, 85, 86, 87, 88, 89, 90, 93, 95]

# Pre Heat
print("Running pre heat")
run("ffmpeg -y -i \"D:\\Footage Dump\\4.mp4\" -c:v libx264 -preset veryslow -crf 0 -an -sn \"D:\\Footage Dump\\tmp.mp4\"", stdout=subprocess.PIPE, stderr=subprocess.PIPE)

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

                outputFileName = str(outputDir) + str(currentSampleShortName)[:-4] + "_" + str(currentCpuUsed) + "_" + str(currentCRF) + "_" + str(currentTargetQuality) + ".mkv"  # [:-4] removed the .mkv extension so i can pu tit on the end
                start_time = time.time()

                # Printing for debugging
                print("cd " + av1anWorkingDIR + " && av1an -i \"" + currentInputSampleName + "\" -v \" --end-usage=q --cq-level=" + str(currentCRF) + " --cpu-used=" + str(currentCpuUsed) + " -t 16\" -w 20 --target-quality " + str(currentTargetQuality) + " --vmaf-path \"" + vmafPath + "\" -o " + outputFileName)
                # Why do i cd and then run av1an? because evil things happen.
                os.system("cd " + av1anWorkingDIR + " && av1an -i \"" + currentInputSampleName + "\" -v \" --end-usage=q --cq-level=" + str(currentCRF) + " --cpu-used=" + str(currentCpuUsed) + " -t 16\" -w 20 --target-quality " + str(currentTargetQuality) + " --vmaf-path \"" + vmafPath + "\" -o " + outputFileName)

                processTime = int(time.time() - start_time)

                worksheet.write(xOffset, yOffset - 1, " --cq-level=" + str(currentCRF))
                worksheet.write(xOffset, yOffset, " --target-quality " + str(currentTargetQuality))

                # Time Taken to excel
                worksheet.write(xOffset, yOffset + 1, processTime)

                # File Size
                try:
                    size = "{:.2f}".format(os.path.getsize(outputFileName) / 1049000)  # IN MEBIBYTES (MiB) with 2 decimal places
                except:
                    size = -1
                worksheet.write(xOffset, yOffset + 3, str(size))

                # txt Backup

                f = open(workingFileDir + currentSampleShortName + "_Sizes.txt", "a")
                f.write("-crf " + str(currentCpuUsed) + " " + str(currentCRF) + " " + str(currentTargetQuality) + ":" + str(size) + "\n")
                f.close()

                f = open(workingFileDir + currentSampleShortName + "_Times.txt", "a")
                f.write("-crf " + str(currentCpuUsed) + " " + str(currentCRF) + " " + str(currentTargetQuality) + ":" + str(processTime) + "\n")
                f.close()

                xOffset += 1

    workbook.close()

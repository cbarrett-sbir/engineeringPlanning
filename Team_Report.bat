@echo off

rem Activate Conda environment
call conda activate excel

rem Change directory to the location of the Python script
cd /d "Q:\EngineeringPlanning\ReportTools\stable builds\Team Report Generator"

rem Start a new command prompt window and run the Python script in it
start cmd /k "python main.py"

rem Deactivate Conda environment (optional)
call conda deactivate
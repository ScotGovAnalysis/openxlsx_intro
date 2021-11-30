# This script calls 'Run Excel Macro.vbs' which opens VBA within the 
# 'VBA - Convert XLSX to ODS.xlsm' file, which creates .ods files from all 
# compatible .xlsx files in the /output folder.

# Requires to have both the .vbs and the .xlsm files in the Rproject folder.

system_command <- paste("WScript",
                        '"Run Excel Macro.vbs"',
                        sep = " ")
system(command = system_command,
       wait = TRUE)

print("ods. files have been saved to /output")

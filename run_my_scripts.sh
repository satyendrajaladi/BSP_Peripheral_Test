 #!/bin/bash

# Read Python and Excel Files as Inputs,Write Result to Excel Workbook.

helpFunction()

{

   echo ""

   echo "Usage: $0 -a Main.py -b Automation.xlsx "

   echo -e "\t-a Python Main Script File eg: main.py"

   echo -e "\t-b Excel Workbook eg: /home/user/PycharmProjects/Automation.xlsx"

 

   exit 1 # Exit script after printing help

}



while getopts "a:b:c:" opt

do

   case "$opt" in

      a ) parameterA="$OPTARG" ;;

      b ) parameterB="$OPTARG" ;;

     

      ? ) helpFunction ;; # Print helpFunction in case parameter is non-existent

   esac

done



# Print helpFunction in case parameters are empty

if [ -z "$parameterA" ] || [ -z "$parameterB" ]  

then

   echo "Some or all of the parameters are empty";

   helpFunction

fi



# Begin script in case all parameters are correct

echo "Given Python File: $parameterA"

echo "Given ExcelWorkbook File: $parameterB"



./venv/bin/python $parameterA $parameterB    


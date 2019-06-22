# How to use it


## Before we get started...

Please install nodejs on your computer:
https://nodejs.org/en/download/


## Getting started and understanding a few things.

1. Click on the button `Clone or Download` on the right side of the page.

2. After downloading the zip file, extract the archive.

3. Ok, now we have a couple files inside the folder, let me describe what each one means

* ```blacklist.json```

    Have all the words in a json format that is **NOT** allowed. the `filtrar_clientes.js` script will read that file and remove from the cell all these values. That's how the script filter the client name.

* ```config```

    All the project settings are inside that file, it means that you can controll all the code and make it work for any spreadsheet just changing the values in this file.
    
* ```filtrar_clientes.js```

    This is the script that will read the `blacklist.json` file and then filter the labels in the spreadsheet to retrive the client name.

* ```filtrar_severidades.js```

    This script will define what number of severity is inside the cell value. I mean, classify as `Sev1`=`1`,`Sev2`=`2`,`Sev3`=`3` and `Sev4`=`4`. And set a random value for the severity that is empty.

* ```package.json```

    Have all the dependencies and libraries that the project needs to run. So never delete this file.


## Config File

As I said, the project have a main file called `config` that have all the parameters for each action in each script. Here's the structure
```
LABELS_COLUMN=P
DESCRIPTION_COLUMN=K
STORE_CLIENT_COLUMN=AF
CLIENTS_COLUMN=L
SOURCE_FILE=Metricas Maio.xlsx
OUTPUT_FILE=new.xlsx
WORKSHEET=Dados
STORE_SEVERITY_COLUNM=AH

SEV_SUMMARY_LABELS=AI
SEV_SUMMARY_VALUES=AJ

SEV_SUMMARY_CLIENT_NAME=AK
SEV_SUMMARY_CLIENT_SEV1=AL
SEV_SUMMARY_CLIENT_SEV2=AM
SEV_SUMMARY_CLIENT_SEV3=AN
SEV_SUMMARY_CLIENT_SEV4=AO
```

* LABELS_COLUMN
    
    Say to the script where to look for the `labels`, in this case, in the column `P`.

* DESCRIPTION_COLUMN
    
    Say to the script where to look for the `description`, in this case, in the column `K`.
    
* STORE_CLIENT_COLUMN
    
    ![](https://github.com/BrunoEleodoro/GabrielPlanilha/blob/master/doc/Screen%20Shot%202019-06-22%20at%2000.21.32.png?raw=true)

    Say to the script where to **save** the client name, where to **store** the values. in this case `AF`

* CLIENTS_COLUMN
    
    If your spreadsheet already have a column with the client names already filled, so put the column name here.

* SOURCE_FILE
    
    The input filename, in this case `Metricas Maio.xlsx`

* OUTPUT_FILE
    
    The output filename, in this case `new.xlsx`
    
    *It's recommended that you save in a different spreadsheet to prevent data loss. But if you want to proceed anyway, you can put the same name of the SOURCE_FILE, it means that when the script finish to run, the spreadsheet will be replaced with the new values*

* WORKSHEET
    
    the name of the worksheet that the script will read all the cells and columns, in this case `Dados`.

* STORE_SEVERITY_COLUNM
       
    Where the script will **save**, **store** the severity numbers, in this case `AH`

The script have a functionallity that counts the amount of `sev1`, `sev2`, `sev3` and `sev4`. To control where this data will be **saved**, **stored** you have to change this two parameters.

*   SEV_SUMMARY_LABELS (`sev1`, `sev2`, `sev3` and `sev4`)
*   SEV_SUMMARY_VALUES (`189`, `200`,`20` and `230`)

And lastely, here is the parameters to say where the script will store the data for each client acording to each severity.
The script generate a summary view of all the clients and the amount of severities for each one. And you can say where to store these values here:

*   EV_SUMMARY_CLIENT_NAME=AK
*   SEV_SUMMARY_CLIENT_SEV1=AL
*   SEV_SUMMARY_CLIENT_SEV2=AM
*   SEV_SUMMARY_CLIENT_SEV3=AN
*   SEV_SUMMARY_CLIENT_SEV4=AO

![](https://github.com/BrunoEleodoro/GabrielPlanilha/blob/master/doc/Screen%20Shot%202019-06-22%20at%2000.24.10.png?raw=true)


## Finally running the script.

1. Open the terminal (`cmd`) and then navigate to the folder where you extracted the files in the `Step 1`.

![](https://github.com/BrunoEleodoro/GabrielPlanilha/blob/master/doc/Screen%20Shot%202019-06-22%20at%2000.30.31.png?raw=true)

2. Now just type `node filtrar_clientes.js` or `node filtrar_severidades.js`

![](https://github.com/BrunoEleodoro/GabrielPlanilha/blob/master/doc/Screen%20Shot%202019-06-22%20at%2000.28.54.png?raw=true)

3. Now you're done, you should see a file in the folder named `new.xlsx` or with the name that you've set in `OUTPUT_FILE` in the `config`.

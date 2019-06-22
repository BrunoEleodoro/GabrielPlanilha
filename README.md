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


## Running the project






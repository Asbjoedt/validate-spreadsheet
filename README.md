# validate-spreadsheet
The program validates any .xlsx or .ods spreadsheet filepath according to the their file format standards and according to [archival data quality specifications](https://github.com/Asbjoedt/CLISC/wiki/Archival-Data-Quality)[^1]. It can be used in simple archival workflows.

* For more information, see repository **[CLISC](https://github.com/Asbjoedt/CLISC)**

## Dependencies
:warning: **[ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html)**
* ODF Validator is used for validating OpenDocument Spreadsheets file format (.ods).
* You need to install program in "C:\Program Files\ODF Validator" and name program "odfvalidator-0.10.0-jar-with-dependencies.jar", or create environment variable "ODFValidator" with path to your installation
* ODF Validator needs latest version of Java Development Kit installed

## How to use
Download the executable version [here](https://github.com/Asbjoedt/validate-spreadsheet/releases). There's no need to install. In your terminal change directory to the folder where validate-spreadsheet.exe is. Then, to execute the program input:
```
.\validate-spreadsheet.exe inputfilepath="[filepath]"
```

**Optional parameters**

```
--standard //If you want to validate the file format standard
--archivalrequirements //If you want to validate archival data quality specifications
```

**Exit codes**

The program writes information to the terminal and it returns an exit code to integrate in your workflows.
```
0 = spreadsheet is invalid
1 = spreadsheet is valid
2 = program error occured (e.g. unsupported file format or ODF Validator was not found)
```

## Packages and software
The following packages and software are used under license.
* [ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html), Apache License, [copyright info](https://github.com/tdf/odftoolkit/blob/master/NOTICE)
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation

[^1]: The program supports validation of .xlsx file format standard and archival data quality specifications. For ods. the program currently only supports validation of file format standard.

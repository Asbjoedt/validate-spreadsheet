# validate-spreadsheets
The program validates any .xlsx or .ods spreadsheet filepaths according to the their file format standards and according to [archival data quality specifications](https://github.com/Asbjoedt/CLISC/wiki/Archival-Data-Quality). It can be used in simple archival workflows.

* For more information, see repository **[CLISC](https://github.com/Asbjoedt/CLISC)**

## Dependencies
:warning: **[ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html)**
* ODF Validator is used for validating OpenDocument Spreadsheets file format (.ods).
* You need to install program in "C:\Program Files\ODF Validator" and name program "odfvalidator-0.10.0-jar-with-dependencies.jar", or create environment variable "ODFValidator" with path to your installation
* ODF Validator needs latest version of Java Development Kit installed

## How to use
Download the executable version [here](https://github.com/Asbjoedt/validate-spreadsheets/releases). There's no need to install. In your terminal change directory to the folder where validate-spreadsheets.exe is. Then, to execute the program input:
```
.\validate-spreadsheets.exe "[filepath]"
```

## Packages and software
The following packages and software are used under license.
* [ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html), Apache License, [copyright info](https://github.com/tdf/odftoolkit/blob/master/NOTICE)
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation

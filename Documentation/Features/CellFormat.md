# Supported cell formats

_**Note to future maintainers:**
Don't change the enum values of OpenXml CellFormats._

| ID   | Enum Name            | Description               |  Format Example  |
|:----:| -------------------- | ------------------------- | :--------------: |
|  0   | Text                 | Default format            |     wibble       |
|  1   | Integer              | Integer format            |       0          |
|  2   | Decimal2Dp           | 2 decimal places          |      0.00        |
|  3   | IntegerWithCommas    | Adds commas for 1000+     |     10,000       |
|  4   | Decimal2DpWithCommas | Adds commas for 1000+     |     10,000.00    |
|  9   | Percentage           | Decimals from 0.00 - 1.00 |       20%        |
|  10  | Percentage2Dp        | Percentage to 2DP         |      20.00%      |
|  11  | Scientific           | Standard Form             |      2.00E+01    |
|  12  | FractionSmall        | fraction                  |       1/2        |
|  13  | FractionLarge        | fraction                  |      23/50       |
|  14  | Date                 | Date format               |    19/04/1991    |
|  20  | TimeSpanHours        | Hours, minutes            |      01:30       |
|  21  | TimeSpanFull         | Hours, minutes, seconds   |     01:30:00     |
|  22  | DateTime             | dd/mm/yyyy H:mm           | 10/02/2019 21:04 |
|  30  | DateVariant          | dd/mm/yyyy                |    10/02/2019    |
|  44  | Currency             | Deal with it              |      £10.00      |
|  45  | TimeSpanMinutes      | Minutes, seconds          |       30:00      |
|  164 | AccountingGBP        | Pounds Sterling           |      £10.00      |
|  165 | AccountingUSD        | US Dollar                 |      $20.00      |
|  166 | AccountingEUR        | Euros                     |      €30.00      |
| 1000 | BooleanOneZero       | 1 or 0                    |        1         |
| 1100 | BooleanTrueFalse     | TRUE or FALSE             |      TRUE        |
| 1200 | BooleanYesNo         | Yes or No                 |       Yes        |
| 1300 | BooleanYN            | Y or N                    |       Y          |

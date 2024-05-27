# XLeratorOptions

The `XLeratorOptions` are used to configure the interaction with the spreadsheet. 
 
## FilePath

The path to the file you want to use.

Type: `string` <br>
Required

## SheetName

The name for the Sheet in the spreadsheet you want to use.

Type: `string` <br>
Default: "Sheet1"

## HeaderLength

If the Spreadsheet has a header, this property must be used to set the amount of rows (**1 Based**) that are consider as a row.
If the value is zero, this means, there is no header and the first row already contains the value.

Type: `int` <br>
Default: 0
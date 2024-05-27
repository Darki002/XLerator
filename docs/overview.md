# Overview

## Attributes

## Factory

The `XLeratorFactory` is used to create a new Instance of any `ExcelUtility`.

Use the `XLeratorFactory.CreateFactory` method to create a new Factory. Also see [XLeratorOptions](./XLeratorOptions.md)
Here is an example;

```csharp
var options = new XLeratorOptions 
{
    FilePath = 'path/to/file.xsls'
}
var factory = XLeratorFactory<YourClass>.CreateFactory(options);
```
The Generic Type is the class that you want to Serialize or Deserialize.

- `CreateExcelCreator`: returns a new `IExcelCreator`
- `CreateExcelEditor`: returns a new `IExcelEditor`
- `CreateExcelIterator`: returns a new `IExcelIterator`
- `CreateExcelReader`: returns a new `IExcelReader`

## Excel Utilities

The following functionalities are possible to use. Also see [Factory](#factory).

- [IExcelCreator](./creater.md): used to create a new Spreadsheet.
- [IExcelEditor](./editor.md): used to edit an existing Spreadsheet.
- [IExcelIterator](./iterator.md): used to iterator over the Spreadsheet.
- [IExcelReader](./reader.md): used to read a Spreadsheet.

## Plans for the future

- option for custom Mapper
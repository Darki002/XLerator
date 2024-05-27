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

### Creator

### Editor

### Iterator

### Reader

## Excel Utilities

The following functionalities are possible to use. Also see [Factory](#factory).

- [Creator](./creater.md): used to create a new Spreadsheet.
- [Editor](./editor.md): used to edit an existing Spreadsheet.
- [Iterator](./iterator.md): used to iterator over the Spreadsheet.
- [Reader](./reader.md): used to read a Spreadsheet.

## Plans for the future

- option for custom Mapper
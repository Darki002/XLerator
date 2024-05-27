# Overview

## Attributes

## Factory

The `XLeratorFactory` is used to create a new Instance of any `ExcelUtility`.

- `CreateExcelCreator`: [CreatorFactory](./Factory.md#creator) 
- `CreateExcelEditor`: [EditorFactory](./Factory.md#editor)
- `CreateExcelIterator`: [IteratorFactory](./Factory.md#iterator)
- `CreateExcelReader`: [ReaderFactory](./Factory.md#reader)

## Excel Utilities

The following functionalities are possible to use. Also see [Factory](#factory).

- [Creator](./creater.md): used to create a new Spreadsheet.
- [Editor](./editor.md): used to edit an existing Spreadsheet.
- [Iterator](./iterator.md): used to iterator over the Spreadsheet.
- [Reader](./reader.md): used to read a Spreadsheet.

## Plans for the future

- option for custom Mapper
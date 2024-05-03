﻿using XLerator.ExcelUtility.Editor;

namespace XLerator.ExcelUtility.Creator;

/// <summary>
/// Creates a new spreadsheet. It will be structure based on <typeparamref name="T"/>.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet</typeparam>
public interface IExcelCreator<in T> where T : class
{
    /// <summary>
    /// Creates a new Excel file and returns a new Instance of a <see cref="IExcelEditor{T}"/>.
    /// </summary>
    /// <param name="addHeader">If True adds a header line on creation.</param>
    /// <returns>The new Instance of a <see cref="IExcelEditor{T}"/>.</returns>
    IExcelEditor<T> CreateExcel(bool addHeader);
}
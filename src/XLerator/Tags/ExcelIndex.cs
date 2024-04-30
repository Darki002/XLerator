﻿namespace XLerator.Tags;

[AttributeUsage(AttributeTargets.Property)]
public abstract class ExcelIndex(int index) : Attribute
{
    public int Index { get; } = index;
}
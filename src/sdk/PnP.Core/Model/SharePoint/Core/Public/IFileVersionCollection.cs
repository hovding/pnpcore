﻿namespace PnP.Core.Model.SharePoint
{
    /// <summary>
    /// Collection of file versions
    /// </summary>
    [ConcreteType(typeof(FileVersionCollection))]
    public interface IFileVersionCollection : IDataModelCollection<IFileVersion>, ISupportQuery<IFileVersion>
    {

    }
}

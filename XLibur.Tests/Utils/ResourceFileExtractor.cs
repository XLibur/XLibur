using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace XLibur.Tests.Utils;

/// <summary>
/// Summary description for ResourceFileExtractor.
/// </summary>
public sealed class ResourceFileExtractor
{
    #region Static

    #region Private fields

    private static readonly IDictionary<string, ResourceFileExtractor> extractors =
        new ConcurrentDictionary<string, ResourceFileExtractor>();

    #endregion Private fields

    #region Public properties

    /// <summary>Instance of resource extractor for executing assembly </summary>
    public static ResourceFileExtractor Instance
    {
        get
        {
            var assembly = Assembly.GetCallingAssembly();
            var key = assembly.GetName().FullName;
            if (extractors.TryGetValue(key, out var extractor)
                || extractors.TryGetValue(key, out extractor)) return extractor;
            extractor = new ResourceFileExtractor(assembly, true, null);
            extractors.Add(key, extractor);

            return extractor;
        }
    }

    #endregion Public properties

    #endregion Static

    #region Private fields

    private readonly ResourceFileExtractor _mBaseExtractor;

    #endregion Private fields

    #region Constructors

    /// <summary>
    /// Create instance
    /// </summary>
    /// <param name="resourceFilePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
    /// <param name="baseExtractor"></param>
    public ResourceFileExtractor(string resourceFilePath, ResourceFileExtractor baseExtractor)
        : this(Assembly.GetCallingAssembly(), baseExtractor)
    {
        ResourceFilePath = resourceFilePath;
    }

    /// <summary>
    /// Create instance
    /// </summary>
    /// <param name="baseExtractor"></param>
    public ResourceFileExtractor(ResourceFileExtractor baseExtractor)
        : this(Assembly.GetCallingAssembly(), baseExtractor)
    {
    }

    /// <summary>
    /// Create instance
    /// </summary>
    /// <param name="resourcePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
    public ResourceFileExtractor(string resourcePath)
        : this(Assembly.GetCallingAssembly(), resourcePath)
    {
    }

    /// <summary>
    /// Instance constructor
    /// </summary>
    /// <param name="assembly"></param>
    /// <param name="resourcePath"></param>
    public ResourceFileExtractor(Assembly assembly, string resourcePath)
        : this(assembly ?? Assembly.GetCallingAssembly())
    {
        ResourceFilePath = resourcePath;
    }

    /// <summary>
    /// Instance constructor
    /// </summary>
    public ResourceFileExtractor()
        : this(Assembly.GetCallingAssembly())
    {
    }

    /// <summary>
    /// Instance constructor
    /// </summary>
    /// <param name="assembly"></param>
    public ResourceFileExtractor(Assembly assembly)
        : this(assembly ?? Assembly.GetCallingAssembly(), (ResourceFileExtractor)null)
    {
    }

    /// <summary>
    /// Instance constructor
    /// </summary>
    /// <param name="assembly"></param>
    /// <param name="baseExtractor"></param>
    public ResourceFileExtractor(Assembly assembly, ResourceFileExtractor baseExtractor)
        : this(assembly ?? Assembly.GetCallingAssembly(), false, baseExtractor)
    {
    }

    /// <summary>
    /// Instance constructor
    /// </summary>
    /// <param name="assembly"></param>
    /// <param name="isStatic"></param>
    /// <param name="baseExtractor"></param>
    /// <exception cref="ArgumentNullException">Argument is null.</exception>
    private ResourceFileExtractor(Assembly assembly, bool isStatic, ResourceFileExtractor baseExtractor)
    {
        Assembly = assembly ?? throw new ArgumentNullException(nameof(assembly));
        _mBaseExtractor = baseExtractor;
        AssemblyName = Assembly.GetName().Name;
        IsStatic = isStatic;
        ResourceFilePath = ".Resources.";
    }

    #endregion Constructors

    #region Public properties

    /// <summary> Work assembly </summary>
    public Assembly Assembly { get; }

    /// <summary> Work assembly name </summary>
    public string AssemblyName { get; }

    /// <summary>
    /// Path to read resource files. Example: .Resources.Upgrades.
    /// </summary>
    public string ResourceFilePath { get; }

    public bool IsStatic { get; set; }

    public IEnumerable<string> GetFileNames(Func<String, Boolean> predicate = null)
    {
        predicate ??= (s => true);

        var path = AssemblyName + ResourceFilePath;
        foreach (var resourceName in Assembly.GetManifestResourceNames())
        {
            if (resourceName.StartsWith(path) && predicate(resourceName))
            {
                yield return resourceName.Replace(path, string.Empty);
            }
        }
    }

    #endregion Public properties

    #region Public methods

    public string ReadFileFromResource(string fileName)
    {
        var stream = ReadFileFromResourceToStream(fileName);
        string result;
        var sr = new StreamReader(stream);
        try
        {
            result = sr.ReadToEnd();
        }
        finally
        {
            sr.Close();
        }

        return result;
    }

    public string ReadFileFromResourceFormat(string fileName, params object[] formatArgs)
    {
        return string.Format(ReadFileFromResource(fileName), formatArgs);
    }

    /// <summary>
    /// Read file in the current assembly by a specific path
    /// </summary>
    /// <param name="specificPath">Specific path</param>
    /// <param name="fileName">Read the file name</param>
    public string ReadSpecificFileFromResource(string specificPath, string fileName)
    {
        var ext = new ResourceFileExtractor(Assembly, specificPath);
        return ext.ReadFileFromResource(fileName);
    }

    /// <summary>
    /// Read a file in the current assembly by a specific file name
    /// </summary>
    /// <param name="fileName"></param>
    /// <exception cref="ApplicationException"><c>ApplicationException</c>.</exception>
    public Stream ReadFileFromResourceToStream(string fileName)
    {
        var nameResFile = AssemblyName + ResourceFilePath + fileName;
        var stream = Assembly.GetManifestResourceStream(nameResFile);

        #region Not found

        if (stream is null)
        {
            #region Get from base extractor

            return _mBaseExtractor is not null
                ? _mBaseExtractor.ReadFileFromResourceToStream(fileName)
                : throw new ArgumentException("Can't find resource file " + nameResFile, nameof(fileName));

            #endregion Get from base extractor
        }

        #endregion Not found

        return stream;
    }

    #endregion Public methods
}

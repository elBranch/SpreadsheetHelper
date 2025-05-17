using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using System.Resources;

namespace SpreadsheetHelper.Internal;

/// <summary>
///     Provides utility methods for managing and retrieving localized resources.
/// </summary>
internal static class ResMan
{
    /// <summary>
    ///     A cache for storing <see cref="ResourceManager" /> instances, keyed by their associated assemblies.
    /// </summary>
    private static readonly ConcurrentDictionary<Assembly, ResourceManager> ResourceManagerCache = new();

    /// <summary>
    ///     Retrieves the <see cref="ResourceManager" /> for the specified assembly, creating and caching it if necessary.
    /// </summary>
    /// <param name="assembly">The assembly for which to retrieve the resource manager.</param>
    /// <returns>A <see cref="ResourceManager" /> instance for the specified assembly.</returns>
    private static ResourceManager GetResourceManager(Assembly assembly)
    {
        return ResourceManagerCache.GetOrAdd(assembly, asm => new ResourceManager(asm.GetName().FullName, asm));
    }

    /// <summary>
    ///     Retrieves a localized string resource by its key from the specified assembly.
    /// </summary>
    /// <param name="assembly">The assembly containing the resource.</param>
    /// <param name="resourceKey">The key of the resource to retrieve.</param>
    /// <returns>The localized string if found; otherwise, a message indicating the key was not found.</returns>
    public static string GetString(Assembly assembly, string resourceKey)
    {
        return GetResourceManager(assembly).GetString(resourceKey) ?? $"Key {resourceKey} not found.";
    }

    /// <summary>
    ///     Retrieves and formats a localized string resource by its key from the specified assembly.
    /// </summary>
    /// <param name="assembly">The assembly containing the resource.</param>
    /// <param name="resourceKey">The key of the resource to retrieve.</param>
    /// <param name="args">An array of objects to format the string with.</param>
    /// <returns>The formatted localized string.</returns>
    public static string Format(Assembly assembly, string resourceKey, params object[] args)
    {
        return string.Format(CultureInfo.CurrentUICulture, GetString(assembly, resourceKey), args);
    }

    /// <summary>
    ///     Retrieves a localized string resource by its key from the calling assembly.
    /// </summary>
    /// <param name="resourceKey">The key of the resource to retrieve.</param>
    /// <returns>The localized string if found; otherwise, a message indicating the key was not found.</returns>
    public static string GetString(string resourceKey)
    {
        return GetString(Assembly.GetCallingAssembly(), resourceKey);
    }

    /// <summary>
    ///     Retrieves and formats a localized string resource by its key from the calling assembly.
    /// </summary>
    /// <param name="resourceKey">The key of the resource to retrieve.</param>
    /// <param name="args">An array of objects to format the string with.</param>
    /// <returns>The formatted localized string.</returns>
    public static string Format(string resourceKey, params object[] args)
    {
        return Format(Assembly.GetCallingAssembly(), resourceKey, args);
    }
}
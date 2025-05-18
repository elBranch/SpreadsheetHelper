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
    private static readonly ResourceManager ResourceManager = new("SpreadsheetHelper.Resources.Strings", typeof(ResMan).Assembly);

    /// <summary>
    ///     Retrieves a localized string resource by its key from the calling assembly.
    /// </summary>
    /// <param name="resourceKey">The key of the resource to retrieve.</param>
    /// <returns>The localized string if found; otherwise, a message indicating the key was not found.</returns>
    public static string GetString(string resourceKey)
    {
        return ResourceManager.GetString(resourceKey) ?? $"Key {resourceKey} not found.";
    }

    /// <summary>
    ///     Retrieves and formats a localized string resource by its key from the calling assembly.
    /// </summary>
    /// <param name="resourceKey">The key of the resource to retrieve.</param>
    /// <param name="args">An array of objects to format the string with.</param>
    /// <returns>The formatted localized string.</returns>
    public static string Format(string resourceKey, params object[] args)
    {
        return string.Format(GetString(resourceKey), args);
    }
}
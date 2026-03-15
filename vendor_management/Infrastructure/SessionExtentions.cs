using System.Text.Json;
using Microsoft.AspNetCore.Http;

namespace SYS_VENDOR_MGMT.Infrastructure
{
    /// <summary>
    /// Extension helpers to store / retrieve typed objects in ISession.
    /// Replaces using Session["key"] as SomeType in code-behind.
    /// </summary>
    public static class SessionExtensions
    {
        public static void Set<T>(this ISession session, string key, T value)
        {
            session.SetString(key, JsonSerializer.Serialize(value));
        }

        public static T? Get<T>(this ISession session, string key)
        {
            var value = session.GetString(key);
            return value == null ? default : JsonSerializer.Deserialize<T>(value);
        }
    }
}

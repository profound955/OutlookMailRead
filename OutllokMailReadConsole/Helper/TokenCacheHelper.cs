using Microsoft.Identity.Client;
using System.IO;

namespace OutllokMailReadConsole
{
    public static class TokenCacheHelper
    {
        public static void EnableSerialization(ITokenCache tokenCache)
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
        }

        private static readonly string _fileName = "msalcache.bin3";

        private static readonly object _fileLock = new object();


        private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (_fileLock)
            {
                byte[] data = null;
                if (File.Exists(_fileName))
                    data = File.ReadAllBytes(_fileName);
                args.TokenCache.DeserializeMsalV3(data);
            }
        }

        private static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            if (args.HasStateChanged)
            {
                lock (_fileLock)
                {
                    byte[] data = args.TokenCache.SerializeMsalV3();
                    File.WriteAllBytes(_fileName, data);
                }
            }
        }
    }
}

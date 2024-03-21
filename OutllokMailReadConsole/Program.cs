using Limilabs.Client.IMAP;
using Limilabs.Mail;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Configuration;

namespace OutllokMailReadConsole
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Starting for reading mails...");
            Op3();
            Console.ReadLine();
        }

        private async static Task Op3()
        {
            try
            {
                string userEmail = WebConfigurationManager.AppSettings["SenderEmailid"];
                string userPassword = WebConfigurationManager.AppSettings["Password"];
                string redirectUri = WebConfigurationManager.AppSettings["redirectUri"];
                string clientId = WebConfigurationManager.AppSettings["applicationId"];
                string tenantId = WebConfigurationManager.AppSettings["tenantId"];

                IPublicClientApplication app = PublicClientApplicationBuilder
                    .Create(clientId)
                    .WithTenantId(tenantId)
                    .WithRedirectUri(redirectUri)
                    .Build();

                // This allows saving access/refresh tokens to some storage
                TokenCacheHelper.EnableSerialization(app.UserTokenCache);

                var scopes = new string[]
                {
                     "User.Read",
                     "offline_access",
                     "email",
                     "https://outlook.office.com/IMAP.AccessAsUser.All",
                     //"https://outlook.office.com/POP.AccessAsUser.All",
                     //"https://outlook.office.com/SMTP.Send",
                     "Mail.Read"
                };
                
                string userName;
                string accessToken;

                Console.WriteLine("Signing in...");

                var account = (await app.GetAccountsAsync()).FirstOrDefault();

                try
                {
                    AuthenticationResult refresh = await app
                        .AcquireTokenSilent(scopes, account)
                        .ExecuteAsync();

                    userName = refresh.Account.Username;
                    accessToken = refresh.AccessToken;
                }
                catch (MsalUiRequiredException)
                {
                    Console.WriteLine("Signing in interactively at first time...");
                    var result = await app.AcquireTokenInteractive(scopes).ExecuteAsync();

                    userName = result.Account.Username;
                    accessToken = result.AccessToken;
                }

                Console.WriteLine("Signed in as user: " + userName);

                using (Imap client = new Imap())
                {
                    client.ConnectSSL("outlook.office365.com");
                    client.LoginOAUTH2(userName, accessToken);

                    Console.WriteLine("Accessed outlook.com with user: " + userName);

                    client.SelectInbox();

                    List<long> uids = client.Search(Flag.Unseen);
                    Console.WriteLine(string.Format("-----Count Of Unseen Mails {0}-----", uids.Count()));
                    Console.WriteLine();

                    int curIndex = 1;
                    foreach (long uid in uids)
                    {
                        IMail email = new MailBuilder()
                                .CreateFromEml(client.GetMessageByUID(uid));

                        string subject = email.Subject;
                        string text = email.Text;

                        Console.WriteLine(string.Format("---Mail {0}---", curIndex));
                        Console.WriteLine(string.Format("Subject: {0}", subject));
                        Console.WriteLine(string.Format("Message: {0}", text));
                        Console.WriteLine(string.Format("---Mail {0} End---", curIndex));
                        Console.WriteLine();
                        curIndex++;
                    }
                    client.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while reading mails: " + ex.Message);
            }
        }
    }
}

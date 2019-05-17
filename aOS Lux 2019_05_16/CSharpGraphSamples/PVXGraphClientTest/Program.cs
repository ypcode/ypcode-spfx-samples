using Microsoft.Graph;
using Microsoft.Identity.Client;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PVXGraphClientTest
{
    class Program
    {
        static void Main(string[] args)
        {
            RunAppOnly();
        }

        static void RunAppOnly()
        {
            GraphServiceClient client = AuthenticationHelper.GetAuthenticatedClientForApp();

            SeeAllUsers(client);
            Pause();

            CheckUserEmails(client);
            Pause("Press ENTER to continue...");

            AddNoteFileToAllUsers(client);
            Pause("Press ENTER to exit.");
        }

        static void SeeAllUsers(GraphServiceClient client)
        {
            try
            {
                Console.WriteLine("Getting information about all users...");
                var allUsers = client.Users.Request().GetAsync().Result;
                foreach (var user in allUsers)
                {
                    Console.WriteLine("========= USER ==========");
                    Console.WriteLine($"Principal Name  : {user.UserPrincipalName}");
                    Console.WriteLine($"Display Name    : {user.DisplayName}");
                    Console.WriteLine($"E-mail          : {user.Mail}");
                    Console.WriteLine("=========================");
                }

                Console.WriteLine();
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Error("Cannot get information about all users...", ex);
                Pause("Press ENTER to exit...");
                Environment.Exit(-1);
            }
        }
        static void CheckUserEmails(GraphServiceClient client)
        {
            try
            {
                Console.WriteLine("See all mails of a specific user.");
                Console.Write("Enter the principal name to get the e-mails of: ");
                string userLogin = Console.ReadLine();
                if (string.IsNullOrEmpty(userLogin))
                {
                    return;
                }
                string tenant = ConfigurationManager.AppSettings["Tenant"];
                string userPrincipalName = $"{userLogin}@{tenant}";
                var userAllMails = client.Users[userPrincipalName].Messages.Request().GetAsync().Result;

                if (userAllMails.Count == 0)
                {
                    Console.WriteLine("No e-mails found...");
                }

                foreach (var email in userAllMails)
                {
                    Console.WriteLine("========== E-mail ===========");
                    Console.WriteLine($"Subject         : {email.Subject}");
                    Console.WriteLine($"Received        : {email.ReceivedDateTime}");
                    Console.WriteLine($"Body preview    : {email.BodyPreview}");
                    Console.WriteLine("=============================");
                }

                Console.WriteLine();
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Error("Cannot get information about user emails...", ex);
                Environment.Exit(-1);
            }

        }

        static void AddNoteFileToAllUsers(GraphServiceClient client)
        {
            try
            {
                Console.WriteLine("Creating note txt files in all users OneDrive.");
                Console.Write("Choose the file name [Note]: ");
                string fileNameInput = Console.ReadLine();
                if (string.IsNullOrEmpty(fileNameInput))
                {
                    fileNameInput = "Note";
                }
                Console.Write("Enter the note content [Hello]: ");
                string noteContent = Console.ReadLine();
                if (string.IsNullOrEmpty(fileNameInput))
                {
                    noteContent = "Hello";
                }
                var allUsers = client.Users.Request().GetAsync().Result;
                foreach (var user in allUsers)
                {
                    Console.WriteLine($"Adding file to {user.DisplayName} OneDrive root folder...");
                    string fileName = $"{fileNameInput} {DateTime.Now.ToString("yyyyMMddhhmmss")}.txt";
                    string contentTxt = $"Note for {user.DisplayName}!\n------------------------------\n{noteContent}";
                    MemoryStream contentStream = new MemoryStream(Encoding.UTF8.GetBytes(contentTxt));
                    client.Users[user.Id].Drive.Root.ItemWithPath(fileName).Content.Request().PutAsync<DriveItem>(contentStream).Wait();
                    Console.WriteLine("File added.");
                }

                Console.WriteLine();
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Error("Cannot get information about user emails...", ex);
                Environment.Exit(-1);
            }

        }

        static void Error(string message, Exception ex)
        {
            ConsoleColor initialForeColor = Console.ForegroundColor;
            ConsoleColor initialBackColor = Console.BackgroundColor;

            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Red;

            Console.WriteLine(message);
            if (ex != null)
            {
                Console.WriteLine(ex.Message);

                if (ex.InnerException != null)
                {
                    var msalException = ex.InnerException as MsalException;
                    if (msalException != null && msalException.ErrorCode == "unauthorized_client")
                    {
                        Console.WriteLine("The application is not authorized to access the resource");
                    }
                    else
                    {
                        Console.WriteLine(ex.InnerException.Message);
                    }
                }
            }


            Console.ForegroundColor = initialForeColor;
            Console.BackgroundColor = initialBackColor;

            Pause();
        }

        static void Pause(string message = null)
        {
            Console.WriteLine(message ?? "Press ENTER to continue...");
            Console.ReadLine();
        }
    }
}

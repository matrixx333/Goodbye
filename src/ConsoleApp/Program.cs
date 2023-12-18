using AdClient.Models;
using AdClient.Models.Requests;
using AdClient.Services;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Goodbye
{
    class Program
    {
        static readonly string _rootDomain = ConfigurationManager.AppSettings["RootDomain"];
        static readonly string _rootOu = ConfigurationManager.AppSettings["RootOu"];
        static readonly string _serviceUser = ConfigurationManager.AppSettings["ServiceUser"];
        static readonly string _servicePassword = ConfigurationManager.AppSettings["ServicePassword"];
        static readonly string _canSendEmail = ConfigurationManager.AppSettings["CanSendEmail"];

        static readonly IUsersService _userSvc = new UsersService(_rootDomain, _rootOu, _serviceUser, _servicePassword);
        static readonly IGroupsService _groupSvc = new GroupsService(_rootDomain, _rootOu, _serviceUser, _servicePassword);

        static void Main(string[] args)
        {
            var credentials = new Credentials();

            Login(ref credentials);
            var samAccountName = GetTermUsersActiveDirectoryUserName();
            ValidateAuthorization(credentials);

            var termUser = GetUserAccount(samAccountName);
            var request = CreateMessageRequest(credentials.Username, termUser);

            var message = string.Format("Terminating {0}'s account: ", termUser.DisplayName);

            if (TerminateUserAccount(termUser))
            {
                SuccessfullResponse(message);

                if (Convert.ToBoolean(_canSendEmail))
                {
                    SendEmailMessage(request);
                }
                                
                ExitApplication();
                Environment.Exit(0);
            }
            else
            {
                FailureResponse(message);
                ExitApplication();
                Environment.Exit(1);
            }
        }

        private static void ExitApplication()
        {
            Console.WriteLine();
            Console.WriteLine("Press any key to exit the application.");
            Console.ReadLine();
        }

        private static string GetTermUsersActiveDirectoryUserName()
        {
            Console.WriteLine("Enter the peron's Active Directory username for whom you'd like to terminate: ");
            var samAccountName = Console.ReadLine();

            return samAccountName;
        }

        static void Login(ref Credentials credentials)
        {
            Console.Write("username: ");
            credentials.Username = Console.ReadLine();
            credentials.Password = GetPassword();
            Console.WriteLine();
        }

        static string GetPassword()
        {
            string passord = "";
            Console.Write("password: ");
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    passord += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && passord.Length > 0)
                    {
                        passord = passord.Substring(0, (passord.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            }
            // Stops Receving Keys Once Enter is Pressed
            while (key.Key != ConsoleKey.Enter);

            return passord;
        }

        static void ValidateAuthorization(Credentials credentials)
        {
            var isValid = _userSvc.ValidateCredentials(credentials.Username, credentials.Password);

            if (isValid)
            {
                var userGroups = _groupSvc.GetUserGroups(credentials.Username).ToList();
                var isAuthorized = (userGroups.Any(g => g.Name.Contains("HR")) ||
                                    userGroups.Any(g => g.Name.Contains("Domain Admins")));

                if (!isAuthorized)
                {
                    Console.WriteLine("You must be a member of the HR group in order to use this application.");
                    Environment.Exit(1);
                }
            }
            else
            {
                Console.WriteLine("Incorrect username or password. Please check your credentials and try again.");
                Environment.Exit(1);
            }
        }

        private static User GetUserAccount(string samAccountName)
        {
            User user = null;
            var response = "N";

            while (response != "Y")
            {
                user = _userSvc.GetUser(samAccountName);

                if (user == null)
                {
                    Console.Write("\nError: Unable to locate ");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write("{0} ", samAccountName);
                    Console.ResetColor();
                    Console.WriteLine("in Active Directory. Please check the spelling of the user name and try again.");
                    Environment.Exit(1);
                }

                Console.WriteLine();
                Console.Write("Are you sure ");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("{0} ", user.DisplayName);
                Console.ResetColor();
                Console.WriteLine("is the person you want to disable?");
                Console.Write("Enter 'Y' for yes, or 'N' for no: ");
                response = Console.ReadLine();
                Console.WriteLine();

                if (response == "N")
                {
                    Console.WriteLine();
                    Console.Write("Enter the person's username you want to disable: ");
                    samAccountName = Console.ReadLine();
                }
            }

            return user;
        }

        private static MessageRequest CreateMessageRequest(string serviceUser, User termUser)
        {
            var svcUser = _userSvc.GetUser(serviceUser);

            var request = new MessageRequest
            {
                ServiceUser = svcUser,
                TermUser = termUser,
                TermDate = DateTime.Now,
                IpPhone = termUser.IpPhone,
                AltRecipient = termUser.AltRecipient,
                UserGroups = _groupSvc.GetUserGroups(termUser.SamAccountName).ToList()
            };

            return request;
        }

        private static bool TerminateUserAccount(User termUser)
        {
            var today = DateTime.Today;
            var dateOnly = today.Date;
            var termDate = "TERM " + dateOnly.ToString("d");            
            var newContainer = $"OU=Disabled Accounts,{_rootOu}";

            ValidateUserGroupMemberships(termUser);

            _groupSvc.RemoveAllUserGroupMemberships(termUser.SamAccountName);
            _userSvc.MoveUser(termUser.DistinguishedName, newContainer);

            var response = _userSvc.UpdateUser(termDate, false, null, null, true, termUser.SamAccountName);

            return response;
        }

        private static void SendEmailMessage(MessageRequest request)
        {
            string toEmailAddress;
            var hrEmail = ConfigurationManager.AppSettings["HrEmailAddress"];
            var mailHost = ConfigurationManager.AppSettings["SmtpServer"];

            if (request.TermUser.DistinguishedName.Contains("SouthCarolina"))
            {
                toEmailAddress = ConfigurationManager.AppSettings["ScHelpDeskEmailAddress"];
            }
            else
            {
                toEmailAddress = ConfigurationManager.AppSettings["FlHelpDeskEmailAddress"];
            }

            if (request.ServiceUser.EmailAddress == null)
            {
                request.ServiceUser.EmailAddress = "terminationreport@fabrikam.com";
            }

            var message = new MailMessage(request.ServiceUser.EmailAddress, toEmailAddress);
            message.CC.Add(hrEmail);

            var smtpServer = new SmtpClient
            {
                Host = mailHost,
                Port = 25,
                EnableSsl = false,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(request.ServiceUser.SamAccountName, request.ServicePassword),
            };

            message.Subject = String.Format("Termination Report: {0}", request.TermUser.DisplayName);
            var mailBody = new StringBuilder();
            mailBody.AppendFormat("Term Information");
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("Service User: {0}", request.ServiceUser.DisplayName);
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("Termed User: {0}", request.TermUser.DisplayName);
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("Termed User DN: {0}", request.TermUser.DistinguishedName);
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("Term Date: {0}", request.TermDate);
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("Term User Telephone Extension: {0}", request.IpPhone);
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("Term User Alternate Recipient: {0}", request.AltRecipient);
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("Group Memberships");
            mailBody.AppendFormat("\n");
            mailBody.AppendFormat("\n");

            foreach (var group in request.UserGroups)
            {
                mailBody.AppendFormat(@group.Name);
                mailBody.AppendFormat("\n");
            }

            message.Body = mailBody.ToString();

            smtpServer.Send(message);
        }

        static void FailureResponse(string message)
        {
            const int padding = 52;

            Console.Write(message);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("FAILED".PadLeft(padding - message.Length));
            Console.ResetColor();
        }

        static void SuccessfullResponse(string message)
        {
            const int padding = 50;

            Console.Write(message);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("DONE".PadLeft(padding - message.Length));
            Console.ResetColor();
        }

        static void ValidateUserGroupMemberships(User termUser)
        {
            var groups = _groupSvc.GetUserGroups(termUser.SamAccountName);

            if (groups.Any(g => g.Name == "Owners" || g.Name == "HR" || g.Name == "Domain Admins"))
            {
                Console.WriteLine("Cannot terminate {0}'s account. The user account is a member of 'Owners', 'HR', or 'Domain Admins'.", termUser.DisplayName);
                Environment.Exit(1);
            }
        }
    }
}

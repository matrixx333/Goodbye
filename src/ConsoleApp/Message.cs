using AdClient.Models;
using System;
using System.Collections.Generic;

namespace Goodbye
{
    public class Message
    {
        public string InformationMessage { get; set; }
        public string FailureMessage { get; set; }

        public Message(string infomessage, string failmessage)
        {
            InformationMessage = infomessage;
            FailureMessage = failmessage;
        }

        public static IDictionary<string, Message> InitializeMessageDictionary(UserPrincipalEx userPrincipal)
        {
            //Create a list of Message objects
            var messagelist = new Dictionary<string, Message>();

            var infomessage = "Updating description with date of TERM:";
            var failmessage = String.Format("Unable to update '" + userPrincipal.DisplayName + "'s description.");

            messagelist.Add("description", new Message(infomessage, failmessage));

            infomessage = "Disabling messaging features:";
            failmessage = String.Format("Unable to disable '" + userPrincipal.DisplayName + "'s messaging features.");

            messagelist.Add("messaging", new Message(infomessage, failmessage));

            infomessage = "Removing ipPhone extension:";
            failmessage = String.Format("Unable to remove '" + userPrincipal.DisplayName + "'s phone extension.");

            messagelist.Add("ipphone", new Message(infomessage, failmessage));

            infomessage = "Managing user's 'memberOf' groups:";
            failmessage = String.Format("Unable to manage '" + userPrincipal.DisplayName + "'s groups.");

            messagelist.Add("groups", new Message(infomessage, failmessage));

            infomessage = "Disabling the account:";
            failmessage = String.Format("Unable to disable '" + userPrincipal.DisplayName + "'s account.");

            messagelist.Add("disable", new Message(infomessage, failmessage));

            return messagelist;
        }
    }
}

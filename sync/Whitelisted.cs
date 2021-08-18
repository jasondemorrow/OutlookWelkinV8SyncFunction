namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;

    public class Whitelisted
    {
        public static List<string> Emails(string key)
        {
            List<string> emails = new List<string>();
            string delimited = Environment.GetEnvironmentVariable(key);
            if (delimited != null)
            {
                string[] addresses = delimited.Split(';');
                if (addresses != null) 
                {
                    foreach(string email in addresses)
                    {
                        if (IsValidEmail(email))
                        {
                            emails.Add(email.ToLowerInvariant().Trim());
                        }
                    }
                }
            }
            return emails;
        }

        private static bool IsValidEmail(string email)
        {
            try {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch {
                return false;
            }
        }
    }
}
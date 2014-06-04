namespace PublicAPI.Samples
{
    using System;
    using System.IO;
    using System.Net;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Configuration;
    using MailChimp;
    using MailChimp.Lists;
    using MailChimp.Helper;
    using Newtonsoft.Json;
    using CommandLine;
    using BlockEncrypter;
    using System.Xml;
    using System.Net.Mail;
    using System.Text.RegularExpressions;

    class Options
    {
        [Option('l', "listID", Required = true,
          HelpText = "ListID.")]
        public string MailChimpListID{ get; set; }

        [Option('g', "groupID", Required = true,
          HelpText = "groupID.")]
        public string MailChimpGroupID { get; set; }

        [Option('w', "WA", Required = true,
          HelpText = "WildApricotKey.")]
        public string WildApricotKey { get; set; }

        [Option('m', "MC", Required = true,
          HelpText = "MailChimpKey.")]
        public string MailChimpKey { get; set; }

        [Option('e', "email", Required = true,
          HelpText = "Enter the field name of the email field.")]
        public string eMailField { get; set; }

        [Option('f', "email2", Required = true,
          HelpText = "Enter the field name of the email field 2.")]
        public string eMailField2 { get; set; }

        [Option('c', "crypt", Required = false,
          HelpText = "Use crypt.")]
        public bool crypt{ get; set; }

        [Option('v', "verbose", DefaultValue = true,
          HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }
      }
    

    class Program
   
    {

        public static string MailChimpListID = "";
        public static string MailChimpGroupID = "";
        public static string MailChimpKey = "";
        public static string WildApricotKey = "";
        public static string EMailField = "";
        public static string EMailField2 = "";
        public static bool crypt = false;
        
        const string CONST_BaseUrl = "https://api.wildapricot.org/";
        static string AccountId;
                               
        static void Main(string[] args)
        {

            if (GetConfigurationValue("Crypt",false) == "True")
            {
                crypt = true;
            }

            if (GetConfigurationValue("Crypt",false) == "True" & GetConfigurationValue("CryptDone", false) == "False")
            {
                //Crypt
                Console.WriteLine("Please wait until the Configuration file is encrypted.");
                UpdateConfigurationValue("MailChimpListID",true);
                UpdateConfigurationValue("MailChimpGroupID", true);
                UpdateConfigurationValue("MailChimpKey", true);
                UpdateConfigurationValue("WildApricotKey", true);
                UpdateConfigurationValue("EMailField", false);
                UpdateConfigurationValue("EMailField2", false);
                SetAppConfigValue("/configuration/appSettings/add[@key='CryptDone']", "True", AppDomain.CurrentDomain.SetupInformation.ConfigurationFile,false);
                Console.WriteLine("Configuration file is encrypted. Press enter and restart the application.");
                var name = Console.ReadLine();
                Environment.Exit(0);
             }
            
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                // Values from command line
                if (options.Verbose) crypt = options.crypt;
                if (options.Verbose) Console.WriteLine("MailChimpListID: {0}", options.MailChimpListID );
                if (options.Verbose) MailChimpListID = options.MailChimpListID;
                if (options.Verbose) Console.WriteLine("MailChimpGroupID: {0}", options.MailChimpGroupID );
                if (options.Verbose) MailChimpGroupID = options.MailChimpGroupID;
                if (options.Verbose) Console.WriteLine("Wild Apricot Key: {0}", options.WildApricotKey);
                if (options.Verbose) WildApricotKey = options.WildApricotKey;
                if (options.Verbose) Console.WriteLine("MailChimp Key: {0}", options.MailChimpKey);
                if (options.Verbose) MailChimpKey = options.MailChimpKey;
                if (options.Verbose) EMailField = options.eMailField;
                if (options.Verbose) EMailField2 = options.eMailField2;
                DoJob(MailChimpListID, MailChimpGroupID, MailChimpKey, WildApricotKey);
             }                   
            else
            {
                string _MailChimpListID = GetConfigurationValue("MailChimpListID",true);
                string _MailChimpGroupID = GetConfigurationValue("MailChimpGroupID",true);
                string _MailChimpKey = GetConfigurationValue("MailChimpKey",true);
                string _WildApricotKey = GetConfigurationValue("WildApricotKey", true);
                    
                if (_MailChimpListID != "" & _MailChimpGroupID != "" & _MailChimpKey != "" & _WildApricotKey != "")
                {
                    Console.WriteLine("Start working...");
                    DoJob(_MailChimpListID, _MailChimpGroupID, _MailChimpKey, _WildApricotKey);                
                }   
                else
                {
                    Console.WriteLine("Configuration file does not work.");
                    var name = Console.ReadLine();
                    Environment.Exit(0);
                }
            }

            Console.WriteLine("All done.");            
        }   
    

        static string GetConfigurationValue(string key, bool cryptMe)
        {
            // Read values from configuration
            string setting = ConfigurationManager.AppSettings[key];
            if (setting == "")
            {
                return "Error";
            }
            else
            {
                if (crypt == true & cryptMe == true)
                {
                    return BlockEncrypter.DecryptStringBlock(setting, Encoding.ASCII.GetBytes("Pa55w0rd"));
                }
                else { return setting; }
            }
        }


        static void SetConfigurationValue(string key)
        {
            string value;
            Console.WriteLine("Enter " + key + ": ");
            value = Console.ReadLine();

            
            if (value != null)
            {
                Console.WriteLine("Saved: " + value);
                Console.WriteLine("");
                try
                {
                    // Load the app.config file
                    Console.WriteLine("Setting value: " + AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
                    SetAppConfigValue("/configuration/appSettings/add[@key='" + key + "']", value, AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, true);
                    Console.WriteLine("Set value done.");
                   
                }
                catch (ConfigurationErrorsException)
                {
                    Console.WriteLine("Error writing app settings");
                }
             } 
        }

        static void UpdateConfigurationValue(string key, bool cryptMe1)
        {
            string value;

            if (GetConfigurationValue("CryptDone",false) == "True")
            {
                value = GetConfigurationValue(key, cryptMe1);
            }
            else 
            {
                value = GetConfigurationValue(key, false);
            }

                            
            if (value != null)
            {
                try
                {
                    // Load the app.config file
                    SetAppConfigValue("/configuration/appSettings/add[@key='" + key + "']", value, AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, cryptMe1);
                  }
                catch (ConfigurationErrorsException)
                {
                    Console.WriteLine("Error writing app settings");
                }
            }
        }


        static void DoJob(string MailChimpListID, string MailChimpGroupID, string MailChimpKey, string WildApricotKey)
        {
            Console.Write("Processing...");

            ConnectToEntryPoint(WildApricotKey);
            LoadVersionResources(WildApricotKey);
            LoadAccounts(WildApricotKey);
            LoadContactFields(WildApricotKey);
            LoadMembershipLevels(WildApricotKey);
            int _groupID = Convert.ToInt32(MailChimpGroupID);
            LoadContacts("'Email delivery disabled' eq false AND Archived eq false AND 'Member emails and newsletters' eq true", MailChimpListID, _groupID, "Member", MailChimpKey, WildApricotKey);

        }
        
        /// <summary>
        /// This helper method retrieves data from API and parses it to dynamic object.
        /// </summary>
        private static dynamic LoadObject(string url, string WildApricotKey)
        {
            object result;
            var client = new WebClient();
            client.Headers.Add(HttpRequestHeader.Accept, "application/json");
            client.Headers.Add("APIKey", WildApricotKey);

            var stream = client.OpenRead(url);
            using (var reader = new StreamReader(stream))
            {
                var str = reader.ReadToEnd();
                result = JsonConvert.DeserializeObject(str);
            }
            return result as dynamic;
        }

        public static void ConnectToEntryPoint(string WildApricotKey)
        {
            Console.WriteLine();
            Console.WriteLine("Connecting to API entry point");
            var entryPointResources = LoadObject(CONST_BaseUrl, WildApricotKey);
            foreach (var resource in entryPointResources)
            {
                Console.WriteLine("Version:" + resource.Version);
                Console.WriteLine("Name:" + resource.Name);
                Console.WriteLine("Url:" + resource.Url);
            }
        }

        public static void LoadVersionResources(string WildApricotKey)
        {
            Console.WriteLine();
            Console.WriteLine("Loading version resources");

            var versionResources = LoadObject(CONST_BaseUrl + "/v1", WildApricotKey);
            foreach (var resource in versionResources)
            {
                Console.WriteLine("Name:" + resource.Name);
                Console.WriteLine("Url:" + resource.Url);
            }
        }

        public static void LoadAccounts(string WildApricotKey)
        {
            Console.WriteLine();
            Console.WriteLine("Loading account info");

            var accounts = LoadObject(CONST_BaseUrl + "/v1/accounts", WildApricotKey);
            foreach (var account in accounts)
            {
                Console.WriteLine("Id:" + account.Id);
                Console.WriteLine("Url:" + account.Url);
                Console.WriteLine("Name:" + account.Name);
                Console.WriteLine("PrimaryDomainName:" + account.PrimaryDomainName);

                AccountId = account.Id.ToString();

                Console.WriteLine("Resources");
                foreach (var resource in account.Resources)
                {
                    Console.WriteLine("  Name:" + resource.Name);
                    Console.WriteLine("  Url:" + resource.Url);
                    Console.WriteLine("  ------");
                }
            }
        }

        public static void LoadContactFields(string WildApricotKey)
        {
            Console.WriteLine();
            Console.WriteLine("Loading contact fields description");

            var url = string.Format("{0}/v1/accounts/{1}/ContactFields/", CONST_BaseUrl, AccountId);
            var contactFields = LoadObject(url, WildApricotKey);

            foreach (var contactField in contactFields)
            {
                Console.WriteLine("FieldName:" + contactField.FieldName);
                Console.WriteLine("Type:" + contactField.Type);
                Console.WriteLine("Description:" + contactField.Description);
                Console.WriteLine("FieldInstructions:" + contactField.FieldInstructions);

                if (contactField.AllowedValues != null)
                {
                    Console.WriteLine("Allowed values");
                    foreach (var allowedValue in contactField.AllowedValues)
                    {
                        Console.WriteLine("  Id: {0}, Label: {1}", allowedValue.Id, allowedValue.Label);
                    }
                }
                Console.WriteLine("------");
            }
        }

        public static void LoadMembershipLevels(string WildApricotKey)
        {
            Console.WriteLine();
            Console.WriteLine("Loading membership levels");

            var url = string.Format("{0}/v1/accounts/{1}/MembershipLevels/", CONST_BaseUrl, AccountId);
            var levels = LoadObject(url, WildApricotKey);

            foreach (var level in levels)
            {
                Console.WriteLine("Id:" + level.Id);
                Console.WriteLine("Name:" + level.Name);
                Console.WriteLine("Type:" + level.Type);
                Console.WriteLine("MembershipFee:" + level.MembershipFee);
                Console.WriteLine("------");
            }
        }

        public static void LoadContacts(string myFilter, string MailChimpListID, int GroupID, string GroupData, string MailChimpKey, string WildApricotKey)

        {
            Console.WriteLine();
            Console.WriteLine("Loading contacts");

            // Filters results to all members with with email address at gmail.com.
            // 'e-mail' is a system field, but its name can be modified by account admin.
            //  If you have problem running program with this field name, check actual name in admin view or just set filterExpression = "$filter=Member eq true"
            var filterExpression = "$filter=Member eq true and substringof('e-Mail','@gmail.com')";

            // Retrieves only the specified fields. Leave it empty to retrieve all fields.
            // 'First name',Phone,'e-mail' are system fields, but their names can be modified by account admin.
            // If you have problem running program with these fields names, check actual names in admin view or just set selectExpression = String.Empty
            var selectExpression = "$select='First name',Phone,'e-mail','Member since','Member Id'";

            // build url
            //var url = string.Format("{0}/v1/accounts/{1}/Contacts/?{2}&{3}", CONST_BaseUrl, AccountId, filterExpression, selectExpression);
            selectExpression = "";
            string url = "";
                        string _filterExpression;
            if (myFilter == "")
            {
                _filterExpression = "";
            }
            else
            {
                _filterExpression = String.Concat("$filter=", myFilter);
            }         

        
            if (filterExpression == "")
            {
                url = String.Format("{0}/v1/accounts/{1}/Contacts", CONST_BaseUrl, AccountId, filterExpression, selectExpression);
            }
            else
            {
                url = String.Format("{0}/v1/accounts/{1}/Contacts/?{2}&{3}", CONST_BaseUrl, AccountId, _filterExpression, selectExpression);
            }


            var request = LoadObject(url, WildApricotKey);
        


            while (true)
            {
                System.Threading.Thread.Sleep(3000);

                request = LoadObject(request.ResultUrl.ToString(), WildApricotKey);
                string state = request.State.ToString();
              //  Console.WriteLine("Request state is '{0}' at {1}", state, DateTime.Now);


              
                switch (state)
                {
                    case "Complete":
                        {
                            foreach (var contact in request.Contacts)
                            {
                               // Console.WriteLine("Contact #{0}:", contact.Id);
                                string FirstName = "";
                                string LastName = "";
                                string eMailAddress = "";
                                string eMailAddressPersonal = "";
                                string isMember = "";
                                string fieldEMail = "Work e-Mail";
                                string fieldPersonalEMail = "Personal e-Mail";
                                
                                if (EMailField != "")
                                {
                                    fieldEMail = EMailField;
                                }

                                if (EMailField2 != "")
                                {
                                    fieldPersonalEMail = EMailField2;
                                }
                                
                                foreach (var field in contact.FieldValues)
                                                                    
                                {

                                   // Console.WriteLine("  {0}: {1}", field.FieldName, field.Value);

                                    string caseSwitch = field.FieldName;
                                    switch (caseSwitch)
                                    {
                                        case "First name":
                                            FirstName = field.Value;
                                            break;
                                        case "Last name":
                                            LastName = field.Value;
                                            break;
                                        case "E-Mail":
                                            eMailAddress = field.Value;
                                            break;
                                        case "Member":
                                            isMember = field.Value;
                                            break;
                                        default:
                                            if (field.FieldName == fieldEMail)
                                            {
                                                eMailAddress = field.Value;
                                            }
                                            if (field.FieldName == fieldPersonalEMail)
                                            {
                                                eMailAddressPersonal = field.Value;
                                            }                       
                                            break;
                                    }                                                                          
                                }
                                
                                string FullName = FirstName + " " + LastName + " " + eMailAddress;
                                //if (FullName.StartsWith("Oliver"))
                                if (FullName != "")
                                {
                                    //Console.WriteLine(FullName);

                                    MyMergeVar myMergeVars = new MyMergeVar();

                                    if (GroupID != 0)
                                    { 
                                    myMergeVars.Groupings = new List<Grouping>();
                                    myMergeVars.Groupings.Add(new Grouping());
                                    myMergeVars.Groupings[0].Id = GroupID ; // replace with your grouping id
                                    myMergeVars.Groupings[0].GroupNames = new List<string>();
                                    
                                    if (isMember == "True")
                                    {
                                        myMergeVars.Groupings[0].GroupNames.Add("Member");
                                    }
                                    else
                                    {
                                        myMergeVars.Groupings[0].GroupNames.Add("Contact");
                                    }
                                    }
                                    
                                    myMergeVars.FirstName = FirstName;
                                    myMergeVars.LastName = LastName;

                                    MailChimpManager mc = new MailChimpManager(MailChimpKey);

                                    //  Create the email parameter
                                    EmailParameter email = new EmailParameter();
                                    //Console.WriteLine(eMailAddress);
                                    string _Email = "";

                                    if (eMailAddress != "")
                                    {
                                        _Email = eMailAddress;
                                        if (IsValidEmailAddressByRegex(_Email) == false)
                                        {
                                            Console.WriteLine(_Email + " not valid for " + FirstName + " " + LastName);
                                            if (eMailAddressPersonal != "")
                                            {
                                            _Email = eMailAddressPersonal;
                                            Console.WriteLine("Found this " + _Email + " for " + FirstName + " " + LastName);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        _Email = eMailAddressPersonal;
                                     }
                                  
                                                                                                                           
                                    EmailParameter emailPar = new EmailParameter()
                                    {
                                        Email = _Email                                    
                                    };

                                    //Check if user unsubscribed

                                    //  Get the first 100 members of each list:
                                    MembersResult mresults = mc.GetAllMembersForList(MailChimpListID, "unsubscribed");
                                    
                                    //  Write out each member's email address:
                                    bool SubscribeMe = true;
                                    foreach (var member in mresults.Data)
                                    {
                                        //Console.WriteLine("member.Email:"  + member.Email + " -> " + member.Status);
                                        if (member.Email == eMailAddress)
                                        { 
                                            SubscribeMe = false;
                                            //Console.WriteLine("Will not subscribe: " + member.Email);
                                        }
                                    }

                                    if (SubscribeMe)
                                    {
                                        try
                                        {
                                            EmailParameter results = mc.Subscribe(MailChimpListID, emailPar, myMergeVars, "html", false, true, false, false);
                                        }
                                        catch
                                        {
                                            Console.WriteLine("Can't subscribe: " + _Email + " (EMail1: " + eMailAddress + ") (EMail2: " + eMailAddressPersonal + ") " + FirstName + " " + LastName);
                                            Console.ReadLine();
                                        }
                                        
                                    }                                  
                                                                                                                                   
                                }
                                 
                                
                            }
                            return;
                        }
                    case "Failed":
                        {
                            Console.WriteLine("Error:'{0}'", request.ErrorDetails);
                            return;
                        }
                }
            }
        }

        public static bool IsValidEmailAddressByRegex(string mailAddress)
        {
            Regex mailIDPattern = new Regex(@"[\w-]+@([\w-]+\.)+[\w-]+");

            if (!string.IsNullOrEmpty(mailAddress) && mailIDPattern.IsMatch(mailAddress))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static void SetAppConfigValue(string XPathQuery, string value, string filename, bool cryptValue)
        {
                XmlDocument doc = new XmlDocument();
                try
                {
                    doc.Load(filename);
                }
                catch 
                {
                    Console.Write("Can't find " + filename);
                }
                                                
                try
                {
                    XmlElement itemTitle = (XmlElement)doc.SelectSingleNode(XPathQuery);
                    string val = "";
                    if (crypt == true & cryptValue == true) { val = BlockEncrypter.EncryptStringBlock(value, Encoding.ASCII.GetBytes("Pa55w0rd"));}
                    else { val = value; }
                    itemTitle.Attributes["value"].Value = val;                  
                }
                catch 
                {
                    Console.Write("Can't add " + value + " to " + XPathQuery + " in " + filename);
                }

                doc.Save(filename);
                ConfigurationManager.RefreshSection("appSettings");
                //Console.Write("Saving " + filename);
                }        
        
    }

      

    // create a class that inherits MergeVar and add any additional merge variable fields:
    [System.Runtime.Serialization.DataContract]
    public class MyMergeVar : MergeVar
    {
        [System.Runtime.Serialization.DataMember(Name = "FNAME")]
        public string FirstName { get; set; }
        [System.Runtime.Serialization.DataMember(Name = "LNAME")]
        public string LastName { get; set; }
    }


}
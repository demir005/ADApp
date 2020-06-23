using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.IO;
using System.Net.Mail;


namespace ADApp
{
    public class Korisnik
    {
        public int Id { get; set; }
        public string Username { get; set; } //samaccountname
        public string DisplayName { get; set; } //displayname      
        public bool isEnabled { get; set; } //useraccountcontrol
        public bool PassNevExp { get; set; }   //pwdlastset 
    }

    class Program
    {
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private static int noviKorisnika = 0;
        private static int izmjenjenihKorisnika = 0;

        [Obsolete]
        static void Main(string[] args)
        {

            foreach (Korisnik korisnik in VratiKorisnike())
            {
                //Vraca listu korisnika
            }
            //Console.WriteLine($"Ukupno novi korisnika : " + noviKorisnika);
            //Console.WriteLine($"Ukupno izmjenjenih korisnika : " + izmjenjenihKorisnika);

            //*********LOG*******//
            log.Info($"Ukupno novi korisnika : " + noviKorisnika);
            log.Info($"Ukupno izmjenjenih korisnika : " + izmjenjenihKorisnika);
            //Console.ReadLine();

            //**********Send Mails********//
            if (noviKorisnika != 0 || izmjenjenihKorisnika != 0)
            {

                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                SmtpClient mySmtpClient = new SmtpClient("smtp.bih.net.ba");


                string email = config.AppSettings.Settings["from"].Value;
                string pass = config.AppSettings.Settings["Password"].Value;

                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new
                System.Net.NetworkCredential(email, pass);
                mySmtpClient.Credentials = basicAuthenticationInfo;
                mySmtpClient.EnableSsl = true;
                mySmtpClient.Port = 587;


                string sender = config.AppSettings.Settings["from"].Value;
                string reciver = config.AppSettings.Settings["to"].Value;

                MailAddress from = new MailAddress(sender, "ActiveDirectoryInformation");
                MailAddress to = new MailAddress(reciver);
                MailMessage myMail = new MailMessage(from, to);

                myMail.Subject = "ActiveDirectory";
                myMail.SubjectEncoding = System.Text.Encoding.UTF8;

                //set body-message and encoding
                myMail.Body = @"Ukupno novih korisnika:" + noviKorisnika + "<br>" +
                              @"Ukupno izmjenjenih korisnika: " + izmjenjenihKorisnika;
                myMail.BodyEncoding = System.Text.Encoding.UTF8;
                // text or html
                myMail.IsBodyHtml = true;
                mySmtpClient.Send(myMail);

                config.Save();
            }
        }

        public void ExcStrPrc(string Username, string DisplayName, bool isEnable, bool PassNevExp)
        {
            using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DesignSaoOsig1"].ConnectionString))
            {
                SqlCommand cmd = new SqlCommand("ADProcTemp", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@Username", Username.ToString().Trim());
                cmd.Parameters.AddWithValue("@DisplayName", DisplayName.ToString().Trim());
                cmd.Parameters.AddWithValue("@isEnabled", Convert.ToInt32(isEnable));
                cmd.Parameters.AddWithValue("@PassNevExp", Convert.ToInt32(PassNevExp));


                cmd.Parameters.Add("@addedUser", SqlDbType.Int).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("@updatedUser", SqlDbType.Int).Direction = ParameterDirection.Output;

                conn.Open();
                int k = cmd.ExecuteNonQuery();
                var addedUserCount = (int)cmd.Parameters["@addedUser"].Value;
                var updatedUserCount = (int)cmd.Parameters["@updatedUser"].Value;


                if (k != 0)
                {
                }
                noviKorisnika += addedUserCount;
                izmjenjenihKorisnika += updatedUserCount;

                conn.Close();
            }
        }

        public static List<Korisnik> VratiKorisnike()
        {
            List<Korisnik> lstADUsers = new List<Korisnik>();
            string sDomainName = "sarajevoosigura";
            string DomainPath = "LDAP://" + sDomainName;

            string fileLoc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            DirectoryEntry searchRoot = new DirectoryEntry(DomainPath);
            DirectorySearcher search = new DirectorySearcher(searchRoot);

            search.Filter = "(&(objectClass=user)(objectCategory=person))";
            search.PropertiesToLoad.Add("samaccountname"); // Username
            search.PropertiesToLoad.Add("displayname"); // display name
            search.PropertiesToLoad.Add("userAccountControl");  // isEnabled
            search.PropertiesToLoad.Add("pwdLastSet"); //passwordExpires


            DataTable resultsTable = new DataTable();
            resultsTable.Columns.Add("samaccountname");
            resultsTable.Columns.Add("displayname");
            resultsTable.Columns.Add("Neaktivan");
            resultsTable.Columns.Add("dontexpirepassword");

            SearchResult result;
            SearchResultCollection resultCol = search.FindAll();


            if (resultCol != null)
            {
                for (int counter = 0; counter < resultCol.Count; counter++)
                {
                    string UserNameEmailString = string.Empty;

                    result = resultCol[counter];

                    if (result.Properties.Contains("samaccountname")
                        && result.Properties.Contains("displayname"))
                    {
                        int userAccountControl = Convert.ToInt32(result.Properties["userAccountControl"][0]);
                        string samAccountName = Convert.ToString(result.Properties["samAccountName"][0]);

                        int isEnable;
                        int Dont_Expire_Password;


                        if ((userAccountControl & 2) > 0)
                        {
                            isEnable = 0;
                        }
                        else
                        {
                            isEnable = 1;
                        }



                        if ((userAccountControl & 65536) > 0)
                        {
                            Dont_Expire_Password = 1;
                        }
                        else
                        {
                            Dont_Expire_Password = 0;
                        }


                        Korisnik korisnik = new Korisnik();
                        korisnik.Username = (result.Properties["samaccountname"][0]).ToString();
                        korisnik.DisplayName = result.Properties["displayname"][0].ToString();
                        korisnik.isEnabled = Convert.ToBoolean(result.Properties["userAccountControl"][0]);


                        DataRow dr = resultsTable.NewRow();
                        dr["samaccountname"] = korisnik.Username.ToString();
                        dr["displayname"] = korisnik.DisplayName.ToString();
                        dr["neaktivan"] = Math.Abs(isEnable);
                        dr["dontexpirepassword"] = Dont_Expire_Password;


                        resultsTable.Rows.Add(dr);

                        // Poziva se store procedura
                        Program p = new Program();
                        p.ExcStrPrc(korisnik.Username.ToString().Trim(), korisnik.DisplayName.ToString().Trim(), Convert.ToBoolean(isEnable), Convert.ToBoolean(Dont_Expire_Password));
                        lstADUsers.Add(korisnik);
                    }

                }
                var json = JsonConvert.SerializeObject(resultCol, Formatting.Indented);
                var res = json;
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(fileLoc, "output.txt"), true))
                {
                    outputFile.WriteLine(fileLoc, json);
                }
                //File.WriteAllText(fileLoc, json);
            }
            return lstADUsers;
        }
    }
}

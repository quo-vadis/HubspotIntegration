using HubSpotIntegration.Model;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace HubSpotIntegration
{
    class Program
    {
        static string hubApi = "";

        static void Main(string[] args)
        {
            var contacts = A(new DateTime(2018, 1, 24, 19, 27, 15));
            B(contacts);
        }

        private static List<Contact> A(DateTime modifiedOnOrAfter)
        {
            string customersUrl = $"https://api.hubapi.com/contacts/v1/lists/recently_updated/contacts/recent?hapikey={hubApi}";
            var contacts = new List<Contact>();

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(customersUrl);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "GET";

            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();

                JObject jObject = JObject.Parse(result);
                JToken jhasmore = jObject["contacts"];
                             
                foreach (var item in jhasmore)
                {
                    var lastmodifieddate = (double)item["properties"]["lastmodifieddate"]["value"];
                    DateTime lastModifiedDate = ConvertFromUnixTimestamp(lastmodifieddate);

                    if (lastModifiedDate > modifiedOnOrAfter)
                    {
                        var companyName = item["properties"]["company"] != null ? item["properties"]["company"]["value"].ToString() : "";
                        contacts.Add(new Contact()
                        {
                            vid =   (int)item["vid"],
                            firstname =      item["properties"]["firstname"]        != null ? item["properties"]["firstname"]["value"].ToString()       : "",
                            lastname =       item["properties"]["lastname"]         != null ? item["properties"]["lastname"]["value"].ToString()        : "",
                            lifecyclestage = item["properties"]["lifecyclestage"]   != null ? item["properties"]["lifecyclestage"]["value"].ToString()  : "",
                            associated_company = GetCompanyByName(companyName)
                        });
                    }
                    
                }
            }
            return contacts;
        }

        public static Company GetCompanyByName(string companyName)
        {
            if (!String.IsNullOrEmpty(companyName)){

                Company company = null;
                string companiesUrl = $"https://api.hubapi.com/companies/v2/companies/recent/modified?hapikey={hubApi}";

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(companiesUrl);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "GET";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();

                    JObject jObject = JObject.Parse(result);
                    JToken jhasmore = jObject["results"];
                    foreach (var item in jhasmore)
                    {
                        if (item["properties"]["name"] != null && item["properties"]["name"]["value"].ToString() == companyName)
                        {
                             company = new Company()
                            {
                                company_id = (int)item["companyId"],
                                name =    item["properties"]["name"]     != null ? item["properties"]["name"]    ["value"].ToString()   :"",
                                website = item["properties"]["website"]  != null ? item["properties"]["website"] ["value"].ToString()   :"",
                                city =    item["properties"]["city"]     != null ? item["properties"]["city"]    ["value"].ToString()   :"",
                                state =   item["properties"]["state"]    != null ? item["properties"]["state"]   ["value"].ToString()   : "",
                                zip =     item["properties"]["zip"]      != null ? item["properties"]["zip"]     ["value"].ToString()   : "",
                                phone =   item["properties"]["phone"]    != null ? item["properties"]["phone"]   ["value"].ToString()   :""
                            };
                        }  
                    }
                }
                return company;
            }
            else
            {
                return null;
            }

        }

        public static void B(List<Contact> contacts)
        {
            string path = Path.Combine( Directory.GetCurrentDirectory(), "Contacts.xlsx") ;

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (ExcelPackage p = new ExcelPackage())
            {
                using (FileStream stream = File.Create(path))
                {
                    p.Load(stream);
                    //deleting worksheet if already present in excel file
                    var wk = p.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Contacts");
                    if (wk != null) { p.Workbook.Worksheets.Delete(wk); }

                    p.Workbook.Worksheets.Add("Contacts");
                    p.Workbook.Worksheets.MoveToEnd("Contacts");
                    ExcelWorksheet worksheet = p.Workbook.Worksheets[p.Workbook.Worksheets.Count];

                    worksheet.InsertRow(1, contacts.Count + 1);

                    worksheet.Cells[1, 1].Value = "vid";
                    worksheet.Cells[1, 2].Value = "firstname";
                    worksheet.Cells[1, 3].Value = "lastname";
                    worksheet.Cells[1, 4].Value = "lifecyclestage";
                    worksheet.Cells[1, 5].Value = "company_id";
                    worksheet.Cells[1, 6].Value = "name";
                    worksheet.Cells[1, 7].Value = "website";
                    worksheet.Cells[1, 8].Value = "city";
                    worksheet.Cells[1, 9].Value = "state";
                    worksheet.Cells[1, 10].Value = "zip";
                    worksheet.Cells[1, 11].Value = "phone";

                    for (int row = 2; row < contacts.Count + 2; row++)
                    {
                        for (int col = 1; col <= 11; col++)
                        {
                            worksheet.Cells[row, 1].Value = contacts[row - 2].vid;
                            worksheet.Cells[row, 2].Value = contacts[row - 2].firstname;
                            worksheet.Cells[row, 3].Value = contacts[row - 2].lastname;
                            worksheet.Cells[row, 4].Value = contacts[row - 2].lifecyclestage;
                            if (contacts[row - 2].associated_company != null)
                            {
                                worksheet.Cells[row, 5].Value = contacts[row - 2].associated_company.company_id;
                                worksheet.Cells[row, 6].Value = contacts[row - 2].associated_company.name;
                                worksheet.Cells[row, 7].Value = contacts[row - 2].associated_company.website;
                                worksheet.Cells[row, 8].Value = contacts[row - 2].associated_company.city;
                                worksheet.Cells[row, 9].Value = contacts[row - 2].associated_company.state;
                                worksheet.Cells[row, 10].Value = contacts[row - 2].associated_company.zip;
                                worksheet.Cells[row, 11].Value = contacts[row - 2].associated_company.phone;
                            }

                        }
                    }   
                }
                Byte[] bin = p.GetAsByteArray();
                File.WriteAllBytes(path, bin);
            }
            System.Diagnostics.Process.Start(path);

        }

        public static DateTime ConvertFromUnixTimestamp(double timestamp)
        {
            DateTime origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return origin.AddSeconds(timestamp / 1000); // convert from milliseconds to seconds
        }
    }
}



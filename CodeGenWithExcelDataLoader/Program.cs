using CodeGenWithExcelDataLoader.EarlyboundEntities;
using CodeGenWithExcelDataLoader.Enums;
using CodeGenWithExcelDataLoader.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Activities.Expressions;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CodeGenWithExcelDataLoader
{
    class Program
    {
        private static readonly string _DataMigrationConnectionString = ConfigurationManager.AppSettings["DataMigrationConnectionString"];

        private static bool _DryRunMode = bool.Parse(ConfigurationManager.AppSettings["DryRunMode"]);

        private static bool _MigrateClients = bool.Parse(ConfigurationManager.AppSettings["MigrateClients"]);
        private static bool _MigrateAccounts = bool.Parse(ConfigurationManager.AppSettings["MigrateAccounts"]);
        private static bool _MigrateProducts = bool.Parse(ConfigurationManager.AppSettings["MigrateProducts"]);
        private static bool _MigrateInvoices = bool.Parse(ConfigurationManager.AppSettings["MigrateInvoices"]);

        private static int _ClientIdNumStart = int.Parse(ConfigurationManager.AppSettings["ClientIdNumStart"]);
        private static int _AccountIdNumStart = int.Parse(ConfigurationManager.AppSettings["AccountIdNumStart"]);
        private static int _ProductIdNumStart = int.Parse(ConfigurationManager.AppSettings["ProductIdNumStart"]);
        private static int _InvoiceIdNumStart = int.Parse(ConfigurationManager.AppSettings["InvoiceIdNumStart"]);

        public static IOrganizationService _serviceProxy;
        static void Main(string[] args)
        {
            _serviceProxy = GetOrganizationService();
            MigrateClients();
            MigrateAccounts();
            MigrateProducts();
            MigrateInvoices();
        }
        private static bool MigrateClients() {
            if (!_MigrateClients) return true;

            Console.WriteLine("Cleaning up Clients");
            _serviceProxy.RetrieveEntities<EarlyboundEntities.Contact>()
                .Where(x => x.new_migrationid >= _ClientIdNumStart).ToList()
                .ForEach(x => _serviceProxy.Delete(x.LogicalName, x.Id));

            Console.WriteLine("Importing Clients");
            var index = _ClientIdNumStart;
            var records = GetClients();

            foreach (var record in records)
            {
                try
                {
                    Console.WriteLine($"Migrating {index}: ({record.id}) {record.last_name}, {record.first_name}");
                    var contact = new Contact();
                    //ADD MAPPNG HERE
                    contact.new_migrationid = int.Parse(record.id);
                    contact.FirstName = record.first_name;
                    contact.LastName = record.last_name;
                    contact.EMailAddress1 = record.email;
                    contact.Address1_Line1 = record.address;
                    contact.MobilePhone = record.cell_phone;
                    contact.Telephone1 = record.home_phone;
                    contact.Telephone2 = record.work_phone;

                    var maxRetryAttempts = 3;
                    var pauseBetweenFailures = TimeSpan.FromSeconds(2);
                    Extensions.RetryOnException(maxRetryAttempts, pauseBetweenFailures, () =>
                    {
                        if (!_DryRunMode) contact.Id = _serviceProxy.Create(contact);
                        index++;
                    });

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }

            return true;

        }
        private static bool MigrateAccounts() {
            if (!_MigrateAccounts) return true;

            Console.WriteLine("Cleaning up Accounts");
            _serviceProxy.RetrieveEntities<EarlyboundEntities.Account>()
                .Where(x => x.new_migrationid >= _AccountIdNumStart).ToList()
                .ForEach(x => _serviceProxy.Delete(x.LogicalName, x.Id));

            Console.WriteLine("Importing Accounts");
            var index = _AccountIdNumStart;
            var records = GetAccounts();

            foreach (var record in records)
            {
                try
                {
                    Console.WriteLine($"Migrating {index}: ({record.id}) {record.name}");
                    var account = new EarlyboundEntities.Account();
                    //ADD MAPPNG HERE
                    account.new_migrationid = int.Parse(record.id);
                    account.Name = record.name;
                    account.Address1_Line1 = record.address;
                    account.Telephone1 = record.phone;
                    account.WebSiteURL = record.website;
                    account.TickerSymbol = record.ticker;
                    account.Fax = record.fax;

                    var maxRetryAttempts = 3;
                    var pauseBetweenFailures = TimeSpan.FromSeconds(2);
                    Extensions.RetryOnException(maxRetryAttempts, pauseBetweenFailures, () =>
                    {
                        if (!_DryRunMode) account.Id = _serviceProxy.Create(account);
                        index++;
                    });

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }

            return true;
        }
        private static bool MigrateProducts() {
            if (!_MigrateProducts) return true;

            Console.WriteLine("Cleaning up Products");
            _serviceProxy.RetrieveEntities<EarlyboundEntities.Product>()
                .Where(x => x.new_migrationid >= _ProductIdNumStart).ToList()
                .ForEach(x => _serviceProxy.Delete(x.LogicalName, x.Id));

            Console.WriteLine("Importing Products");
            var index = _ProductIdNumStart;
            var records = GetProducts();

            foreach (var record in records)
            {
                try
                {
                    Console.WriteLine($"Migrating {index}: ({record.id}) {record.name}");
                    var product = new EarlyboundEntities.Product();
                    //ADD MAPPNG HERE
                    product.new_migrationid = int.Parse(record.id);
                    product.Name = record.name;
                    product.ProductNumber = record.product_id;
                    product.ValidFromDate =DateTime.Parse(record.valid_from);
                    product.ValidToDate = DateTime.Parse(record.valid_to);
                    product.Description = record.description;

                    var maxRetryAttempts = 3;
                    var pauseBetweenFailures = TimeSpan.FromSeconds(2);
                    Extensions.RetryOnException(maxRetryAttempts, pauseBetweenFailures, () =>
                    {
                        if (!_DryRunMode) product.Id = _serviceProxy.Create(product);
                        index++;
                    });

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }

            return true;
        }
        private static bool MigrateInvoices() {
            if (!_MigrateInvoices) return true;

            Console.WriteLine("Cleaning up Invoices");
            _serviceProxy.RetrieveEntities<EarlyboundEntities.Invoice>()
                .Where(x => x.new_migrationid >= _InvoiceIdNumStart).ToList()
                .ForEach(x => _serviceProxy.Delete(x.LogicalName, x.Id));

            Console.WriteLine("Importing Invoices");
            var index = _InvoiceIdNumStart;
            var records = GetInvoices();

            foreach (var record in records)
            {
                try
                {
                    Console.WriteLine($"Migrating {index}: ({record.id}) {record.name}");
                    var invoice = new EarlyboundEntities.Invoice();
                    //ADD MAPPNG HERE
                    invoice.new_migrationid = int.Parse(record.id);
                    invoice.Name = record.name;
                    invoice.InvoiceNumber = record.invoice_id;


                    var maxRetryAttempts = 3;
                    var pauseBetweenFailures = TimeSpan.FromSeconds(2);
                    Extensions.RetryOnException(maxRetryAttempts, pauseBetweenFailures, () =>
                    {
                        if (!_DryRunMode) invoice.Id = _serviceProxy.Create(invoice);
                        index++;
                    });

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }

            return true;
        }
        private static List<Models.Client> GetClients()
        {
            Console.WriteLine("Loading Clients");
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(_DataMigrationConnectionString);
            Worksheet xlWorksheet = xlWorkbook.Sheets[WorksheetsEnum.Clients];
            Range xlRange = xlWorksheet.UsedRange;
            var records = new List<Client>();
            try
            {
                foreach (Range row in xlRange.Rows)
                {
                    int rowNumber = row.Row;
                    if (rowNumber == 1) continue;
                    var record = new Client();
                    //ADD MAPPNG HERE
                    record.id = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.id].Text;
                    record.first_name = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.first_name].Text;
                    record.last_name = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.last_name].Text;
                    record.email = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.email].Text;
                    record.address = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.address].Text;
                    record.cell_phone = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.cell_phone].Text;
                    record.home_phone = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.home_phone].Text;
                    record.work_phone = (string)xlRange.Cells[rowNumber, (int)ClientColumnEnum.work_phone].Text;

                    records.Add(record);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            return records;
        }
        private static List<Models.Account> GetAccounts()
        {
            Console.WriteLine("Loading Accounts");
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(_DataMigrationConnectionString);
            Worksheet xlWorksheet = xlWorkbook.Sheets[WorksheetsEnum.Accounts];
            Range xlRange = xlWorksheet.UsedRange;
            var records = new List<Models. Account>();
            try
            {
                foreach (Range row in xlRange.Rows)
                {
                    int rowNumber = row.Row;
                    if (rowNumber == 1) continue;
                    var record = new Models.Account();
                    //ADD MAPPNG HERE
                    record.id = (string)xlRange.Cells[rowNumber, (int)AccountColumnEnum.id].Text;
                    record.name = (string)xlRange.Cells[rowNumber, (int)AccountColumnEnum.name].Text;
                    record.address = (string)xlRange.Cells[rowNumber, (int)AccountColumnEnum.address].Text;
                    record.phone = (string)xlRange.Cells[rowNumber, (int)AccountColumnEnum.phone].Text;
                    record.website = (string)xlRange.Cells[rowNumber, (int)AccountColumnEnum.website].Text;
                    record.ticker = (string)xlRange.Cells[rowNumber, (int)AccountColumnEnum.ticker].Text;
                    record.fax = (string)xlRange.Cells[rowNumber, (int)AccountColumnEnum.fax].Text;


                    records.Add(record);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            return records;
        }
        private static List<Models.Product> GetProducts()
        {
            Console.WriteLine("Loading Products");
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(_DataMigrationConnectionString);
            Worksheet xlWorksheet = xlWorkbook.Sheets[WorksheetsEnum.Products];
            Range xlRange = xlWorksheet.UsedRange;
            var records = new List<Models.Product>();
            try
            {
                foreach (Range row in xlRange.Rows)
                {
                    int rowNumber = row.Row;
                    if (rowNumber == 1) continue;
                    var record = new Models.Product();
                    //ADD MAPPNG HERE
                    record.id = (string)xlRange.Cells[rowNumber, (int)ProductColumnEnum.id].Text;
                    record.name = (string)xlRange.Cells[rowNumber, (int)ProductColumnEnum.name].Text;
                    record.product_id = (string)xlRange.Cells[rowNumber, (int)ProductColumnEnum.product_id].Text;
                    record.valid_from = (string)xlRange.Cells[rowNumber, (int)ProductColumnEnum.valid_from].Text;
                    record.valid_to = (string)xlRange.Cells[rowNumber, (int)ProductColumnEnum.valid_to].Text;
                    record.description = (string)xlRange.Cells[rowNumber, (int)ProductColumnEnum.description].Text;


                    records.Add(record);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            return records;
        }
        private static List<Models.Invoice> GetInvoices()
        {
            Console.WriteLine("Loading Invoices");
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(_DataMigrationConnectionString);
            Worksheet xlWorksheet = xlWorkbook.Sheets[WorksheetsEnum.Invoices];
            Range xlRange = xlWorksheet.UsedRange;
            var records = new List<Models.Invoice>();
            try
            {
                foreach (Range row in xlRange.Rows)
                {
                    int rowNumber = row.Row;
                    if (rowNumber == 1) continue;
                    var record = new Models.Invoice();
                    //ADD MAPPNG HERE
                    record.id = (string)xlRange.Cells[rowNumber, (int)InvoiceColumnEnum.id].Text;
                    record.name = (string)xlRange.Cells[rowNumber, (int)InvoiceColumnEnum.name].Text;
                    record.invoice_id = (string)xlRange.Cells[rowNumber, (int)InvoiceColumnEnum.invoice_id].Text;


                    records.Add(record);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            return records;
        }
        private static IOrganizationService GetOrganizationService()
        {
            var crmConnectionOrgName = ConfigurationManager.AppSettings["CrmConnectionOrgName"];
            var crmConnectionUsername = ConfigurationManager.AppSettings["CrmConnectionUsername"];
            var crmConnectionPassword = ConfigurationManager.AppSettings["CrmConnectionPassword"];
            var crmConenctionRegion = ConfigurationManager.AppSettings["CrmConenctionRegion"];

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            var pwd = ConvertToSecureString(crmConnectionPassword);
            CrmServiceClient conn = new CrmServiceClient(crmConnectionUsername, pwd, crmConenctionRegion, crmConnectionOrgName, isOffice365: true);
            var orgSvc = conn.OrganizationWebProxyClient ?? (IOrganizationService)conn.OrganizationServiceProxy;
            return orgSvc;
        }
        private static System.Security.SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("missing pwd");

            var securePassword = new System.Security.SecureString();
            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }
    }
}

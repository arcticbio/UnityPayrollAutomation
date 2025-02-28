using System;
using System.Collections.Generic;
using Interop.QBFC16;
using System.Text;

namespace QuickBooksDiagnostics
{
    public class Program
    {
        private static QBSessionManager sessionManager;
        private const string APP_NAME = "QB Diagnostics Tool";
        private const string APP_ID = "UnityDispatch.Diagnostics";

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("QuickBooks Diagnostics Tool");
            Console.WriteLine("---------------------------");

            try
            {
                if (!InitializeQBConnection())
                {
                    Console.WriteLine("Failed to connect to QuickBooks. Ensure QuickBooks is open and try again.");
                    return;
                }

                // Get and display company information
                DisplayCompanyInfo();

                // Get and display preferences
                DisplayPreferences();

                // Get and display employee statistics
                DisplayEmployeeStats();

                // Get and display vendor statistics
                DisplayVendorStats();

                // Get and display customer statistics
                DisplayCustomerStats();

                // Get and display payroll item information
                DisplayPayrollItems();

                // Get and display account list summary
                DisplayAccountsSummary();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                }
            }
            finally
            {
                CloseQBConnection();
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
        }

        private static bool InitializeQBConnection()
        {
            try
            {
                sessionManager = new QBSessionManager();
                sessionManager.OpenConnection2(APP_ID, APP_NAME, ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing QuickBooks connection: {ex.Message}");
                return false;
            }
        }

        private static void CloseQBConnection()
        {
            if (sessionManager != null)
            {
                try
                {
                    sessionManager.EndSession();
                    sessionManager.CloseConnection();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error closing QuickBooks connection: {ex.Message}");
                }
            }
        }

        private static void DisplayCompanyInfo()
        {
            Console.WriteLine("\nCompany Information:");
            Console.WriteLine("-------------------");

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            ICompanyQuery companyQuery = requestMsgSet.AppendCompanyQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet.ResponseList != null && responseMsgSet.ResponseList.Count > 0)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    ICompanyRet companyRet = (ICompanyRet)response.Detail;
                    Console.WriteLine($"Company Name: {companyRet.CompanyName.GetValue()}");
                    Console.WriteLine($"Legal Company Name: {companyRet.LegalCompanyName.GetValue()}");

                    Console.WriteLine($"First Month of Fiscal Year: {companyRet.FirstMonthFiscalYear.GetValue()}");
                    Console.WriteLine($"First Month of Income Tax Year: {companyRet.FirstMonthIncomeTaxYear.GetValue()}");
                }
            }
        }

        private static void DisplayPreferences()
        {
            Console.WriteLine("\nPreferences:");
            Console.WriteLine("------------");

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            IPreferencesQuery preferencesQuery = requestMsgSet.AppendPreferencesQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet.ResponseList != null && responseMsgSet.ResponseList.Count > 0)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IPreferencesRet preferencesRet = (IPreferencesRet)response.Detail;
                    //Console.WriteLine($"Using Account Numbers: {preferencesRet.AccountingPreferences.UseAccountNumbers.GetValue()}");
                    //Console.WriteLine($"Using Classes: {preferencesRet.AccountingPreferences.UseClasses.GetValue()}");
                    //Console.WriteLine($"Using Custom Transaction Numbers: {preferencesRet.AccountingPreferences.UseCustomTransactionNumbers.GetValue()}");
                    //Console.WriteLine($"Multiple Users Enabled: {preferencesRet.AccountingPreferences.IsUsingMultiUserMode.GetValue()}");
                }
            }
        }

        private static void DisplayEmployeeStats()
        {
            Console.WriteLine("\nEmployee Statistics:");
            Console.WriteLine("-------------------");

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            IEmployeeQuery employeeQuery = requestMsgSet.AppendEmployeeQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet.ResponseList != null && responseMsgSet.ResponseList.Count > 0)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IEmployeeRetList employeeList = (IEmployeeRetList)response.Detail;
                    Console.WriteLine($"Total Employees: {employeeList.Count}");

                    Console.WriteLine("\nEmployee List:");
                    for (int i = 0; i < employeeList.Count; i++)
                    {
                        IEmployeeRet employee = employeeList.GetAt(i);
                        Console.WriteLine($"- {employee.Name.GetValue()} (ID: {employee.ListID.GetValue()})");

                        if (employee.IsActive != null)
                            Console.WriteLine($"  Active: {employee.IsActive.GetValue()}");

                        if (employee.PrintAs != null)
                            Console.WriteLine($"  Print As: {employee.PrintAs.GetValue()}");
                    }
                }
            }
        }

        private static void DisplayVendorStats()
        {
            Console.WriteLine("\nVendor Statistics:");
            Console.WriteLine("-----------------");

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            IVendorQuery vendorQuery = requestMsgSet.AppendVendorQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet.ResponseList != null && responseMsgSet.ResponseList.Count > 0)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IVendorRetList vendorList = (IVendorRetList)response.Detail;
                    Console.WriteLine($"Total Vendors: {vendorList.Count}");

                    int activeVendors = 0;
                    for (int i = 0; i < vendorList.Count; i++)
                    {
                        IVendorRet vendor = vendorList.GetAt(i);
                        if (vendor.IsActive != null && vendor.IsActive.GetValue())
                            activeVendors++;
                    }

                    Console.WriteLine($"Active Vendors: {activeVendors}");
                    Console.WriteLine($"Inactive Vendors: {vendorList.Count - activeVendors}");
                }
            }
        }

        private static void DisplayCustomerStats()
        {
            Console.WriteLine("\nCustomer Statistics:");
            Console.WriteLine("-------------------");

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            ICustomerQuery customerQuery = requestMsgSet.AppendCustomerQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet.ResponseList != null && responseMsgSet.ResponseList.Count > 0)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    ICustomerRetList customerList = (ICustomerRetList)response.Detail;
                    Console.WriteLine($"Total Customers: {customerList.Count}");

                    int activeCustomers = 0;
                    for (int i = 0; i < customerList.Count; i++)
                    {
                        ICustomerRet customer = customerList.GetAt(i);
                        if (customer.IsActive != null && customer.IsActive.GetValue())
                            activeCustomers++;
                    }

                    Console.WriteLine($"Active Customers: {activeCustomers}");
                    Console.WriteLine($"Inactive Customers: {customerList.Count - activeCustomers}");
                }
            }
        }

        private static void DisplayPayrollItems()
        {
            Console.WriteLine("\nPayroll Items:");
            Console.WriteLine("--------------");

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            IPayrollItemWageQuery payrollQuery = requestMsgSet.AppendPayrollItemWageQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet.ResponseList != null && responseMsgSet.ResponseList.Count > 0)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IPayrollItemWageRetList payrollList = (IPayrollItemWageRetList)response.Detail;
                    Console.WriteLine($"Total Payroll Items: {payrollList.Count}");

                    Console.WriteLine("\nPayroll Item List:");
                    for (int i = 0; i < payrollList.Count; i++)
                    {
                        IPayrollItemWageRet item = payrollList.GetAt(i);
                        Console.WriteLine($"- {item.Name.GetValue()} (ID: {item.ListID.GetValue()})");
                        if (item.IsActive != null)
                            Console.WriteLine($"  Active: {item.IsActive.GetValue()}");
                    }
                }
            }
        }

        private static void DisplayAccountsSummary()
        {
            Console.WriteLine("\nAccounts Summary:");
            Console.WriteLine("----------------");

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            IAccountQuery accountQuery = requestMsgSet.AppendAccountQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet.ResponseList != null && responseMsgSet.ResponseList.Count > 0)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IAccountRetList accountList = (IAccountRetList)response.Detail;

                    Dictionary<ENAccountType, int> accountTypes = new Dictionary<ENAccountType, int>();

                    for (int i = 0; i < accountList.Count; i++)
                    {
                        IAccountRet account = accountList.GetAt(i);
                        ENAccountType accountType = (ENAccountType)account.AccountType.GetValue();

                        if (!accountTypes.ContainsKey(accountType))
                            accountTypes[accountType] = 0;

                        accountTypes[accountType]++;
                    }

                    Console.WriteLine($"Total Accounts: {accountList.Count}");
                    Console.WriteLine("\nAccounts by Type:");
                    foreach (var kvp in accountTypes)
                    {
                        Console.WriteLine($"- {kvp.Key}: {kvp.Value}");
                    }
                }
            }
        }
    }
}
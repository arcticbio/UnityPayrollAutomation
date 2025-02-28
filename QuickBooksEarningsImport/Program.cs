using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using Interop.QBFC16;
using System.Windows.Forms;
using System.Text;

namespace QuickBooksEarningsImport
{
    public class Program
    {
        private static QBSessionManager sessionManager;
        private const string APP_NAME = "Unity Earnings Import";
        private const string APP_ID = "UnityDispatch.EarningsImport";
        private const string APP_DESC = "Application to import employee earnings into QuickBooks";
        private const string APP_SUPPORT = "https://unitydispatch.net";

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("QuickBooks Earnings Import Tool");
            Console.WriteLine("-------------------------------");

            try
            {
                // Initialize QuickBooks connection
                if (!InitializeQBConnection())
                {
                    Console.WriteLine("Failed to connect to QuickBooks. Ensure QuickBooks is open and try again.");
                    return;
                }

                // Display company information
                DisplayCompanyInfo();

                // List all employees in QuickBooks
                Console.WriteLine("\nAvailable employees in QuickBooks:");
                Console.WriteLine("----------------------------------");
                List<EmployeeInfo> employees = ListAllEmployees();
                foreach (var emp in employees)
                {
                    Console.WriteLine($"ID: {emp.ListID}, Name: {emp.Name}, Full Name: {emp.FullName}");
                }
                Console.WriteLine("----------------------------------\n");

                // Get the CSV file path
                Console.Write("Enter the path to your CSV file: ");
                string csvPath = Console.ReadLine();

                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"File not found: {csvPath}");
                    return;
                }

                // Read earnings data from CSV
                List<EarningsRecord> earningsRecords = ReadEarningsFromCSV(csvPath);
                Console.WriteLine($"Read {earningsRecords.Count} records from CSV file.");

                // Display a preview of the CSV records
                Console.WriteLine("\nCSV Data Preview:");
                Console.WriteLine("----------------");
                foreach (var record in earningsRecords.GetRange(0, Math.Min(5, earningsRecords.Count)))
                {
                    Console.WriteLine($"Employee: '{record.EmployeeName}', Rate: {record.Rate}, Hours: {record.Hours}");
                }
                Console.WriteLine("----------------\n");

                // Import earnings into QuickBooks
                int successCount = ImportEarningsToQuickBooks(earningsRecords, employees);
                Console.WriteLine($"Successfully imported {successCount} of {earningsRecords.Count} earnings records.");
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
                // Clean up the QuickBooks connection
                CloseQBConnection();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
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
            Console.WriteLine("-------------------\n");
        }

        private static bool InitializeQBConnection()
        {
            try
            {
                sessionManager = new QBSessionManager();

                // Initialize the QBFC session
                sessionManager.OpenConnection2(APP_ID, APP_NAME, ENConnectionType.ctLocalQBD);

                // Begin a session with the currently open company file
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

        private static List<EmployeeInfo> ListAllEmployees()
        {
            List<EmployeeInfo> employees = new List<EmployeeInfo>();

            try
            {
                Console.WriteLine("\nQuerying QuickBooks for employees...");

                // Create a message set request for employees
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);

                // Add more specific query parameters
                IEmployeeQuery employeeQuery = requestMsgSet.AppendEmployeeQueryRq();

                // Set active status to All
                employeeQuery.ORListQuery.ListFilter.ActiveStatus.SetValue(ENActiveStatus.asAll);

                Console.WriteLine("Executing employee query...");
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                if (responseMsgSet == null)
                {
                    Console.WriteLine("Error: Received null response from QuickBooks");
                    return employees;
                }

                if (responseMsgSet.ResponseList == null)
                {
                    Console.WriteLine("Error: Response list is null");
                    return employees;
                }

                // Process the response
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response == null)
                {
                    Console.WriteLine("Error: First response is null");
                    return employees;
                }

                Console.WriteLine($"Response status code: {response.StatusCode}");
                Console.WriteLine($"Response status message: {response.StatusMessage}");

                if (response.StatusCode != 0)
                {
                    Console.WriteLine($"Error retrieving employees: {response.StatusMessage}");
                    return employees;
                }

                if (response.Detail == null)
                {
                    Console.WriteLine("Error: Response detail is null");
                    return employees;
                }

                // Debug response type information
                int responseType = response.Detail.Type.GetValue();
                Console.WriteLine($"Response Type Value: {responseType}");
                Console.WriteLine($"Expected Type Value: {(int)ENResponseType.rtEmployeeQueryRs}");

                // Try to process regardless of type
                try
                {
                    IEmployeeRetList employeeRetList = (IEmployeeRetList)response.Detail;
                    Console.WriteLine($"\nFound {employeeRetList.Count} employees in QuickBooks");

                    for (int i = 0; i < employeeRetList.Count; i++)
                    {
                        IEmployeeRet employeeRet = employeeRetList.GetAt(i);

                        string displayName = employeeRet.Name.GetValue();
                        string fullName = displayName;
                        string firstName = "";
                        string lastName = "";

                        if (employeeRet.FirstName != null)
                        {
                            firstName = employeeRet.FirstName.GetValue();
                        }
                        if (employeeRet.LastName != null)
                        {
                            lastName = employeeRet.LastName.GetValue();
                        }

                        if (!string.IsNullOrEmpty(firstName) && !string.IsNullOrEmpty(lastName))
                        {
                            fullName = $"{firstName} {lastName}";
                        }

                        EmployeeInfo employee = new EmployeeInfo
                        {
                            ListID = employeeRet.ListID.GetValue(),
                            Name = displayName,
                            FullName = fullName,
                            FirstName = firstName,
                            LastName = lastName
                        };

                        employees.Add(employee);
                        Console.WriteLine($"Found Employee: ListID={employee.ListID}, Name={employee.Name}, FullName={employee.FullName}");
                    }
                }
                catch (Exception castEx)
                {
                    Console.WriteLine($"Error casting response to EmployeeRetList: {castEx.Message}");
                    Console.WriteLine($"Actual response type: {response.Detail.GetType().FullName}");

                    // Try to get more information about the response
                    try
                    {
                        dynamic detail = response.Detail;
                        Console.WriteLine("Response Detail Properties:");
                        foreach (var prop in detail.GetType().GetProperties())
                        {
                            try
                            {
                                var value = prop.GetValue(detail);
                                Console.WriteLine($"  {prop.Name}: {value}");
                            }
                            catch
                            {
                                Console.WriteLine($"  {prop.Name}: <Unable to read value>");
                            }
                        }
                    }
                    catch (Exception debugEx)
                    {
                        Console.WriteLine($"Error examining response detail: {debugEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ListAllEmployees: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }

            Console.WriteLine($"\nTotal employees found: {employees.Count}");
            return employees;
        }

        private static List<EarningsRecord> ReadEarningsFromCSV(string filePath)
        {
            List<EarningsRecord> records = new List<EarningsRecord>();

            try
            {
                // Skip header row and read all lines
                string[] lines = File.ReadAllLines(filePath);

                for (int i = 1; i < lines.Length; i++) // Skip header row
                {
                    string line = lines[i];
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        string[] parts = line.Split(',');
                        if (parts.Length >= 3)
                        {
                            EarningsRecord record = new EarningsRecord
                            {
                                EmployeeName = parts[0].Trim(),
                                Rate = decimal.Parse(parts[1].Trim(), CultureInfo.InvariantCulture),
                                Hours = decimal.Parse(parts[2].Trim(), CultureInfo.InvariantCulture)
                            };

                            records.Add(record);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading CSV: {ex.Message}");
            }

            return records;
        }

        private static int ImportEarningsToQuickBooks(List<EarningsRecord> records, List<EmployeeInfo> employees)
        {
            int successCount = 0;

            try
            {
                // Create a message set request to hold our requests
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                // Process each record
                foreach (var record in records)
                {
                    try
                    {
                        // Find the employee in our cached list using flexible matching
                        string employeeListID = FindEmployeeListID(record.EmployeeName, employees);

                        if (string.IsNullOrEmpty(employeeListID))
                        {
                            Console.WriteLine($"Employee not found: '{record.EmployeeName}'. Skipping record.");
                            Console.WriteLine("Please check that the employee name in the CSV matches one of the names listed above.");
                            continue;
                        }

                        // Create a TimeTracking Add request
                        ITimeTrackingAdd timeTrackAdd = requestMsgSet.AppendTimeTrackingAddRq();

                        // Set the TxnDate to today
                        timeTrackAdd.TxnDate.SetValue(DateTime.Today);

                        // Specify the entity (employee) by their ListID
                        timeTrackAdd.EntityRef.ListID.SetValue(employeeListID);

                        // For non-billable payroll entries, we don't need to set a customer
                        timeTrackAdd.IsBillable.SetValue(false);

                        // Set item info for earnings
                        string payrollItemID = GetPayrollItemListID("Regular Pay");
                        if (string.IsNullOrEmpty(payrollItemID))
                        {
                            Console.WriteLine("No payroll item found. Skipping record.");
                            continue;
                        }
                        timeTrackAdd.ItemServiceRef.ListID.SetValue(payrollItemID);

                        // Set the duration in hours and minutes
                        // Convert total hours to whole hours and minutes
                        short wholeHours = (short)Math.Floor(record.Hours);
                        short minutes = (short)Math.Round((record.Hours - wholeHours) * 60);
                        timeTrackAdd.Duration.SetValue(wholeHours, minutes, 0, false);

                        // Set hourly rate - convert decimal to double for SDK compatibility
                        timeTrackAdd.Rate.SetValue(Convert.ToDouble(record.Rate));

                        // Add a descriptive note
                        timeTrackAdd.Notes.SetValue($"Imported earnings for {record.EmployeeName}");

                        // Log what we're doing
                        Console.WriteLine($"Adding earnings for {record.EmployeeName}: {record.Hours} hours at ${record.Rate}/hour");

                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing record for {record.EmployeeName}: {ex.Message}");
                    }
                }

                // Submit all the requests to QuickBooks at once
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                // Process the response
                ProcessResponse(responseMsgSet);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during import: {ex.Message}");
            }

            return successCount;
        }

        private static string FindEmployeeListID(string employeeName, List<EmployeeInfo> employees)
        {
            // Try multiple matching strategies

            // 1. Exact match on Name or constructed FullName
            var exactMatch = employees.Find(e =>
                string.Equals(e.Name, employeeName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(e.FullName, employeeName, StringComparison.OrdinalIgnoreCase));

            if (exactMatch != null)
            {
                Console.WriteLine($"Found exact match for '{employeeName}': {exactMatch.FullName}");
                return exactMatch.ListID;
            }

            // 2. First name only match (for cases where CSV has just the first name)
            var firstNameMatch = employees.Find(e =>
                !string.IsNullOrEmpty(e.FirstName) &&
                string.Equals(e.FirstName, employeeName, StringComparison.OrdinalIgnoreCase));

            if (firstNameMatch != null)
            {
                Console.WriteLine($"Found first name match for '{employeeName}': {firstNameMatch.FullName}");
                return firstNameMatch.ListID;
            }

            // 3. Contains match (for partial names)
            var containsMatch = employees.Find(e =>
                (e.FullName != null && e.FullName.IndexOf(employeeName, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (e.Name != null && e.Name.IndexOf(employeeName, StringComparison.OrdinalIgnoreCase) >= 0));

            if (containsMatch != null)
            {
                Console.WriteLine($"Found partial match for '{employeeName}': {containsMatch.FullName}");
                return containsMatch.ListID;
            }

            // 4. If employee name has multiple parts (like "John Doe"), try matching first name + last name
            string[] nameParts = employeeName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (nameParts.Length > 1)
            {
                var firstLastMatch = employees.Find(e =>
                    !string.IsNullOrEmpty(e.FirstName) &&
                    !string.IsNullOrEmpty(e.LastName) &&
                    string.Equals(e.FirstName, nameParts[0], StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(e.LastName, nameParts[nameParts.Length - 1], StringComparison.OrdinalIgnoreCase));

                if (firstLastMatch != null)
                {
                    Console.WriteLine($"Found first+last name match for '{employeeName}': {firstLastMatch.FullName}");
                    return firstLastMatch.ListID;
                }
            }

            return string.Empty; // No match found
        }

        private static string GetDefaultCustomerListID()
        {
            string customerListID = string.Empty;

            try
            {
                Console.WriteLine("\nQuerying QuickBooks for customers...");

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                ICustomerQuery customerQuery = requestMsgSet.AppendCustomerQueryRq();

                // Set active status to All
                customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue(ENActiveStatus.asAll);

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                if (responseMsgSet != null && responseMsgSet.ResponseList != null)
                {
                    IResponse response = responseMsgSet.ResponseList.GetAt(0);
                    if (response.StatusCode == 0)
                    {
                        if (response.Detail != null && response.Detail.Type.GetValue() == (int)ENResponseType.rtCustomerQueryRs)
                        {
                            ICustomerRetList customerRetList = (ICustomerRetList)response.Detail;

                            Console.WriteLine($"\nAvailable Customers in QuickBooks:");
                            Console.WriteLine("--------------------------------");

                            if (customerRetList.Count > 0)
                            {
                                for (int i = 0; i < customerRetList.Count; i++)
                                {
                                    ICustomerRet customerRet = customerRetList.GetAt(i);
                                    string name = customerRet.FullName.GetValue();
                                    string id = customerRet.ListID.GetValue();
                                    bool isActive = customerRet.IsActive != null && customerRet.IsActive.GetValue();

                                    Console.WriteLine($"Customer: {name} (ID: {id}, Active: {isActive})");

                                    // Take the first active customer as default
                                    if (string.IsNullOrEmpty(customerListID) && isActive)
                                    {
                                        customerListID = id;
                                        Console.WriteLine($"\nUsing customer as default: {name}");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("No customers found in QuickBooks.");
                            }
                            Console.WriteLine("--------------------------------\n");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Error querying customers: {response.StatusMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting customer list: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                }
            }

            if (string.IsNullOrEmpty(customerListID))
            {
                Console.WriteLine("\nWARNING: Could not find any active customers in QuickBooks");
                Console.WriteLine("Please create at least one active customer in QuickBooks to proceed with the import.");
                Console.WriteLine("The customer is required for time tracking entries even if they are non-billable.\n");
            }

            return customerListID;
        }

        private static string GetPayrollItemListID(string payrollItemName)
        {
            string payrollItemListID = string.Empty;

            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
            IPayrollItemWageQuery payrollItemQuery = requestMsgSet.AppendPayrollItemWageQueryRq();

            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

            if (responseMsgSet != null && responseMsgSet.ResponseList != null)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0)
                {
                    if (response.Detail != null && response.Detail.Type.GetValue() == (int)ENResponseType.rtPayrollItemWageQueryRs)
                    {
                        IPayrollItemWageRetList payrollItemRetList = (IPayrollItemWageRetList)response.Detail;
                        for (int i = 0; i < payrollItemRetList.Count; i++)
                        {
                            IPayrollItemWageRet payrollItemRet = payrollItemRetList.GetAt(i);
                            string name = payrollItemRet.Name.GetValue();
                            if (string.Equals(name, payrollItemName, StringComparison.OrdinalIgnoreCase))
                            {
                                payrollItemListID = payrollItemRet.ListID.GetValue();
                                Console.WriteLine($"Using payroll item: {name}");
                                break;
                            }
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(payrollItemListID))
            {
                Console.WriteLine($"Warning: Could not find payroll item: {payrollItemName} in QuickBooks");

                // List available payroll items to help troubleshoot
                Console.WriteLine("\nAvailable payroll items:");
                Console.WriteLine("------------------------");
                ListAvailablePayrollItems();
                Console.WriteLine("------------------------\n");
            }

            return payrollItemListID;
        }

        private static void ListAvailablePayrollItems()
        {
            try
            {
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                IPayrollItemWageQuery payrollItemQuery = requestMsgSet.AppendPayrollItemWageQueryRq();

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                if (responseMsgSet != null && responseMsgSet.ResponseList != null)
                {
                    IResponse response = responseMsgSet.ResponseList.GetAt(0);
                    if (response.StatusCode == 0)
                    {
                        if (response.Detail != null && response.Detail.Type.GetValue() == (int)ENResponseType.rtPayrollItemWageQueryRs)
                        {
                            IPayrollItemWageRetList payrollItemRetList = (IPayrollItemWageRetList)response.Detail;
                            for (int i = 0; i < payrollItemRetList.Count; i++)
                            {
                                IPayrollItemWageRet payrollItemRet = payrollItemRetList.GetAt(i);
                                string name = payrollItemRet.Name.GetValue();
                                Console.WriteLine($"Payroll Item: {name}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing payroll items: {ex.Message}");
            }
        }

        private static void ProcessResponse(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null || responseMsgSet.ResponseList == null) return;

            for (int i = 0; i < responseMsgSet.ResponseList.Count; i++)
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(i);

                // Check if there are any errors
                if (response.StatusCode != 0)
                {
                    string msg = $"Error: {response.StatusCode} - {response.StatusMessage}";
                    if (response.StatusSeverity == "Error")
                    {
                        Console.WriteLine(msg);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: {msg}");
                    }
                }
                else
                {
                    // Success - get details based on response type
                    if (response.Detail != null && response.Detail.Type.GetValue() == (int)ENResponseType.rtTimeTrackingAddRs)
                    {
                        ITimeTrackingRet timeTrackingRet = (ITimeTrackingRet)response.Detail;
                        Console.WriteLine($"Successfully added time entry with TxnID: {timeTrackingRet.TxnID.GetValue()}");
                    }
                }
            }
        }
    }

    public class EarningsRecord
    {
        public string EmployeeName { get; set; }
        public decimal Rate { get; set; }
        public decimal Hours { get; set; }
    }

    public class EmployeeInfo
    {
        public string ListID { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }
}
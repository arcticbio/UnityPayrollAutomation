using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using Interop.QBFC16;
using System.Windows.Forms;
using System.Text;

namespace QuickBooksCheckAddImport
{
    public class QuickBooksCheckAddImport : IDisposable
    {
        private static QBSessionManager sessionManager;
        private const string APP_NAME = "Unity Check Add Import";
        private const string APP_ID = "UnityDispatch.CheckAddImport";
        private const string APP_DESC = "Application to import employee earnings as checks into QuickBooks";
        private const string APP_SUPPORT = "https://unitydispatch.net";
        private static bool isSessionOpen = false;
        private static bool isConnectionOpen = false;

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("QuickBooks Check Add Import Tool");
            Console.WriteLine("--------------------------------");

            using (var program = new QuickBooksCheckAddImport())
            {
                try
                {
                    // Initialize QuickBooks connection
                    if (!program.InitializeQBConnection())
                    {
                        Console.WriteLine("Failed to connect to QuickBooks. Ensure QuickBooks is open and try again.");
                        return;
                    }

                    // Display company information
                    program.DisplayCompanyInfo();

                    // List all employees in QuickBooks
                    Console.WriteLine("\nAvailable employees in QuickBooks:");
                    Console.WriteLine("----------------------------------");
                    List<EmployeeInfo> employees = program.ListAllEmployees();
                    foreach (var emp in employees)
                    {
                        Console.WriteLine($"ID: {emp.ListID}, Name: {emp.Name}, Full Name: {emp.FullName}");
                    }
                    Console.WriteLine("----------------------------------\n");

                    // List available payroll wage items
                    Console.WriteLine("\nAvailable payroll wage items:");
                    Console.WriteLine("-----------------------------");
                    List<PayrollItemInfo> payrollItems = program.ListPayrollWageItems();
                    foreach (var item in payrollItems)
                    {
                        Console.WriteLine($"ID: {item.ListID}, Name: {item.Name}");
                    }
                    Console.WriteLine("-----------------------------\n");

                    // Get the CSV file path
                    Console.Write("Enter the path to your CSV file: ");
                    string csvPath = Console.ReadLine();

                    if (!File.Exists(csvPath))
                    {
                        Console.WriteLine($"File not found: {csvPath}");
                        return;
                    }

                    // Read earnings data from CSV
                    List<EarningsRecord> earningsRecords = program.ReadEarningsFromCSV(csvPath);
                    Console.WriteLine($"Read {earningsRecords.Count} records from CSV file.");

                    // Display a preview of the CSV records
                    Console.WriteLine("\nCSV Data Preview:");
                    Console.WriteLine("----------------");
                    foreach (var record in earningsRecords.GetRange(0, Math.Min(5, earningsRecords.Count)))
                    {
                        Console.WriteLine($"Employee: '{record.EmployeeName}', Amount: {record.Amount}, Type: {record.EarningsType}");
                    }
                    Console.WriteLine("----------------\n");

                    // Import earnings into QuickBooks
                    Console.WriteLine("\nImporting earnings into QuickBooks...");
                    int successCount = program.ImportEarningsToQuickBooks(earningsRecords, employees, payrollItems);
                    Console.WriteLine($"Successfully imported {successCount} of {earningsRecords.Count} earnings records.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                    }
                    Console.WriteLine($"Stack trace: {ex.StackTrace}");
                }
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        public void Dispose()
        {
            CloseQBConnection();
        }

        private void DisplayCompanyInfo()
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

        private bool InitializeQBConnection()
        {
            try
            {
                Console.WriteLine("Initializing QuickBooks connection...");
                sessionManager = new QBSessionManager();

                // Initialize the QBFC session
                sessionManager.OpenConnection2(APP_ID, APP_NAME, ENConnectionType.ctLocalQBD);
                isConnectionOpen = true;
                Console.WriteLine("Connection opened successfully.");

                // Begin a session with the currently open company file
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                isSessionOpen = true;
                Console.WriteLine("Session established successfully.");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing QuickBooks connection: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                }
                return false;
            }
        }

        private void CloseQBConnection()
        {
            if (sessionManager != null)
            {
                try
                {
                    if (isSessionOpen)
                    {
                        Console.WriteLine("Ending QuickBooks session...");
                        sessionManager.EndSession();
                        isSessionOpen = false;
                        Console.WriteLine("Session ended successfully.");
                    }

                    if (isConnectionOpen)
                    {
                        Console.WriteLine("Closing QuickBooks connection...");
                        sessionManager.CloseConnection();
                        isConnectionOpen = false;
                        Console.WriteLine("Connection closed successfully.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error closing QuickBooks connection: {ex.Message}");
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                    }
                }
            }
        }

        private List<EmployeeInfo> ListAllEmployees()
        {
            List<EmployeeInfo> employees = new List<EmployeeInfo>();

            try
            {
                Console.WriteLine("Querying QuickBooks for employees...");

                // Create a message set request for employees
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);

                // Add employee query
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

                try
                {
                    IEmployeeRetList employeeRetList = (IEmployeeRetList)response.Detail;
                    Console.WriteLine($"Found {employeeRetList.Count} employees in QuickBooks");

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
                    }
                }
                catch (Exception castEx)
                {
                    Console.WriteLine($"Error casting response to EmployeeRetList: {castEx.Message}");
                    Console.WriteLine($"Actual response type: {response.Detail.GetType().FullName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ListAllEmployees: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
            }

            Console.WriteLine($"Total employees found: {employees.Count}");
            return employees;
        }

        private List<PayrollItemInfo> ListPayrollWageItems()
        {
            List<PayrollItemInfo> payrollItems = new List<PayrollItemInfo>();

            try
            {
                Console.WriteLine("Querying QuickBooks for payroll wage items...");

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                IPayrollItemWageQuery payrollItemQuery = requestMsgSet.AppendPayrollItemWageQueryRq();

                // Execute the query
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                if (responseMsgSet != null && responseMsgSet.ResponseList != null)
                {
                    IResponse response = responseMsgSet.ResponseList.GetAt(0);
                    if (response.StatusCode == 0 && response.Detail != null)
                    {
                        if (response.Detail.Type.GetValue() == (int)ENResponseType.rtPayrollItemWageQueryRs)
                        {
                            IPayrollItemWageRetList payrollItemRetList = (IPayrollItemWageRetList)response.Detail;
                            Console.WriteLine($"Found {payrollItemRetList.Count} payroll wage items in QuickBooks");

                            for (int i = 0; i < payrollItemRetList.Count; i++)
                            {
                                IPayrollItemWageRet payrollItemRet = payrollItemRetList.GetAt(i);
                                PayrollItemInfo item = new PayrollItemInfo
                                {
                                    ListID = payrollItemRet.ListID.GetValue(),
                                    Name = payrollItemRet.Name.GetValue()
                                };
                                payrollItems.Add(item);
                            }
                        }
                        else
                        {
                            IPayrollItemWageRetList payrollItemRetList = (IPayrollItemWageRetList)response.Detail;
                            Console.WriteLine($"Found {payrollItemRetList.Count} payroll wage items in QuickBooks - test");

                            for (int i = 0; i < payrollItemRetList.Count; i++)
                            {
                                IPayrollItemWageRet payrollItemRet = payrollItemRetList.GetAt(i);
                                PayrollItemInfo item = new PayrollItemInfo
                                {
                                    ListID = payrollItemRet.ListID.GetValue(),
                                    Name = payrollItemRet.Name.GetValue()
                                };
                                payrollItems.Add(item);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Error retrieving payroll items: {response.StatusMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing payroll items: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                }
            }

            return payrollItems;
        }

        private List<EarningsRecord> ReadEarningsFromCSV(string filePath)
        {
            List<EarningsRecord> records = new List<EarningsRecord>();

            try
            {
                Console.WriteLine($"Reading earnings data from {filePath}...");

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
                                Amount = decimal.Parse(parts[1].Trim(), CultureInfo.InvariantCulture),
                                EarningsType = parts[2].Trim()
                            };

                            records.Add(record);
                        }
                        else
                        {
                            Console.WriteLine($"Warning: Line {i + 1} has invalid format: {line}");
                        }
                    }
                }

                Console.WriteLine($"Successfully read {records.Count} earnings records.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading CSV: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                }
            }

            return records;
        }

        private int ImportEarningsToQuickBooks(List<EarningsRecord> records, List<EmployeeInfo> employees, List<PayrollItemInfo> payrollItems)
        {
            int successCount = 0;

            try
            {
                // Ask user for check date
                Console.Write("Enter check date (MM/DD/YYYY) [Default: Today]: ");
                string dateStr = Console.ReadLine();
                DateTime checkDate = DateTime.Today;
                if (!string.IsNullOrWhiteSpace(dateStr))
                {
                    try
                    {
                        checkDate = DateTime.ParseExact(dateStr, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                        Console.WriteLine("Invalid date format. Using today's date instead.");
                    }
                }

                // Process each record
                foreach (var record in records)
                {
                    try
                    {
                        // Find the employee in our cached list
                        string employeeListID = FindEmployeeListID(record.EmployeeName, employees);

                        if (string.IsNullOrEmpty(employeeListID))
                        {
                            Console.WriteLine($"Employee not found: '{record.EmployeeName}'. Skipping record.");
                            continue;
                        }

                        // Find the appropriate payroll item
                        string payrollItemID = FindPayrollItemID(record.EarningsType, payrollItems);

                        if (string.IsNullOrEmpty(payrollItemID))
                        {
                            Console.WriteLine($"Payroll item not found for type: '{record.EarningsType}'. Skipping record.");
                            continue;
                        }

                        // Create a message set request for this employee earnings
                        IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);

                        // Add a Check Add request
                        ICheckAdd checkAdd = requestMsgSet.AppendCheckAddRq();

                        // Set basic check properties
                        checkAdd.PayeeEntityRef.ListID.SetValue(employeeListID);
                        checkAdd.TxnDate.SetValue(checkDate);
                        checkAdd.IsToBePrinted.SetValue(true);
                        checkAdd.Memo.SetValue($"Earnings: {record.EarningsType}");

                        // Add expense line for the earnings
                        IExpenseLineAdd expenseLine = checkAdd.ExpenseLineAddList.Append();
                        expenseLine.AccountRef.ListID.SetValue(payrollItemID);
                        Console.WriteLine($"PayrollItemID: '{payrollItemID}'.");
                        expenseLine.Amount.SetValue((double)record.Amount);
                        expenseLine.Memo.SetValue($"{record.EarningsType} for {record.EmployeeName}");

                        // Log what we're doing
                        Console.WriteLine($"Adding {record.EarningsType} for {record.EmployeeName}: ${record.Amount}");

                        // Execute the request
                        IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                        // Process the response
                        if (ProcessCheckAddResponse(responseMsgSet))
                        {
                            successCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing record for {record.EmployeeName}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during import: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner error: {ex.InnerException.Message}");
                }
            }

            return successCount;
        }

        private bool ProcessCheckAddResponse(IMsgSetResponse responseMsgSet)
        {
            bool success = false;

            if (responseMsgSet == null || responseMsgSet.ResponseList == null) return false;

            IResponse response = responseMsgSet.ResponseList.GetAt(0);

            if (response.StatusCode != 0)
            {
                Console.WriteLine($"Error: {response.StatusCode} - {response.StatusMessage}");
                return false;
            }

            if (response.Detail != null && response.Detail.Type.GetValue() == (int)ENResponseType.rtCheckAddRs)
            {
                ICheckRet checkRet = (ICheckRet)response.Detail;
                Console.WriteLine($"Success: Created check with TxnID: {checkRet.TxnID.GetValue()}");
                success = true;
            }
            else
            {
                Console.WriteLine("Warning: Unexpected response type received.");
            }

            return success;
        }

        private string FindEmployeeListID(string employeeName, List<EmployeeInfo> employees)
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

        private string FindPayrollItemID(string earningsType, List<PayrollItemInfo> payrollItems)
        {
            // Map common earnings types to potential QuickBooks payroll item names
            Dictionary<string, List<string>> earningsTypeMap = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
            {
                { "Commission", new List<string> { "Commission", "Sales Commission", "Commissions" } },
                { "Bonus", new List<string> { "Bonus", "Bonuses", "Employee Bonus" } },
                { "Salary", new List<string> { "Salary", "Regular Salary", "Base Salary" } },
                { "Regular", new List<string> { "Regular Pay", "Hourly Rate", "Regular Wages" } },
                { "Overtime", new List<string> { "Overtime", "OT", "Overtime Pay" } }
            };

            // Determine possible payroll item names for the given earnings type
            List<string> possibleItemNames = new List<string>();
            if (earningsTypeMap.ContainsKey(earningsType))
            {
                possibleItemNames.AddRange(earningsTypeMap[earningsType]);
            }
            // Always add the earnings type itself as a possible match
            possibleItemNames.Add(earningsType);

            // Look for exact matches first
            foreach (string itemName in possibleItemNames)
            {
                PayrollItemInfo exactMatch = payrollItems.Find(p =>
                    string.Equals(p.Name, itemName, StringComparison.OrdinalIgnoreCase));

                if (exactMatch != null)
                {
                    Console.WriteLine($"Found exact payroll item match for '{earningsType}': {exactMatch.Name}");
                    return exactMatch.ListID;
                }
            }

            // Look for contains matches next
            foreach (string itemName in possibleItemNames)
            {
                PayrollItemInfo containsMatch = payrollItems.Find(p =>
                    p.Name.IndexOf(itemName, StringComparison.OrdinalIgnoreCase) >= 0);

                if (containsMatch != null)
                {
                    Console.WriteLine($"Found partial payroll item match for '{earningsType}': {containsMatch.Name}");
                    return containsMatch.ListID;
                }
            }

            // When nothing matches, check if we have any default items to fall back to
            var defaultItem = payrollItems.Find(p =>
                p.Name.IndexOf("Regular", StringComparison.OrdinalIgnoreCase) >= 0 ||
                p.Name.IndexOf("Salary", StringComparison.OrdinalIgnoreCase) >= 0);

            if (defaultItem != null)
            {
                Console.WriteLine($"WARNING: Using default payroll item '{defaultItem.Name}' for '{earningsType}'");
                return defaultItem.ListID;
            }

            // If we still didn't find anything, return the first payroll item if available
            if (payrollItems.Count > 0)
            {
                Console.WriteLine($"WARNING: No matching payroll item for '{earningsType}'. Using first available item: {payrollItems[0].Name}");
                return payrollItems[0].ListID;
            }

            Console.WriteLine($"ERROR: No payroll items found for '{earningsType}'");
            return string.Empty;
        }
    }

    public class EarningsRecord
    {
        public string EmployeeName { get; set; }
        public decimal Amount { get; set; }
        public string EarningsType { get; set; }
    }

    public class EmployeeInfo
    {
        public string ListID { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }

    public class PayrollItemInfo
    {
        public string ListID { get; set; }
        public string Name { get; set; }
    }
}
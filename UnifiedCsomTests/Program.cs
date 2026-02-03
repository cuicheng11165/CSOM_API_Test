using System;
using UnifiedCsomTests.Scenarios;

namespace UnifiedCsomTests
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("===============================================");
            Console.WriteLine("   SharePoint CSOM API Unified Test Console");
            Console.WriteLine("===============================================\n");

            while (true)
            {
                DisplayMainMenu();
                var choice = Console.ReadLine();

                try
                {
                    switch (choice)
                    {
                        case "1":
                            FileOperationsMenu();
                            break;
                        case "2":
                            PermissionMenu();
                            break;
                        case "3":
                            TenantApiMenu();
                            break;
                        case "4":
                            ListApiMenu();
                            break;
                        case "5":
                            ContainerMenu();
                            break;
                        case "6":
                            ViewMenu();
                            break;
                        case "7":
                            ExceptionHandlingMenu();
                            break;
                        case "8":
                            CamlQueryMenu();
                            break;
                        case "9":
                            TimeZoneMenu();
                            break;
                        case "10":
                            TaxonomyMenu();
                            break;
                        case "11":
                            WebPropertiesMenu();
                            break;
                        case "12":
                            OtherScenariosMenu();
                            break;
                        case "0":
                            Console.WriteLine("\nExiting program...");
                            return;
                        default:
                            Console.WriteLine("\nInvalid selection, please try again.\n");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\nError: {ex.Message}");
                    Console.WriteLine($"Details: {ex.StackTrace}\n");
                }

                Console.WriteLine("\nPress any key to continue...");
                Console.ReadKey();
                Console.Clear();
            }
        }

        static void DisplayMainMenu()
        {
            Console.WriteLine("Main Menu:");
            Console.WriteLine("  1. File Operations Scenarios");
            Console.WriteLine("  2. Permission Management Scenarios");
            Console.WriteLine("  3. Tenant API Scenarios");
            Console.WriteLine("  4. List API Scenarios");
            Console.WriteLine("  5. Container Scenarios");
            Console.WriteLine("  6. View Operations Scenarios");
            Console.WriteLine("  7. Exception Handling Scenarios");
            Console.WriteLine("  8. CAML Query Scenarios");
            Console.WriteLine("  9. TimeZone Test Scenarios");
            Console.WriteLine("  10. Managed Metadata Scenarios");
            Console.WriteLine("  11. Web Properties Scenarios");
            Console.WriteLine("  12. Other Scenarios");
            Console.WriteLine("  0. Exit");
            Console.Write("\nPlease select: ");
        }

        static void FileOperationsMenu()
        {
            Console.Clear();
            Console.WriteLine("=== File Operations Scenarios ===\n");
            Console.WriteLine("  1. Add File with Bytes (AddFileWithBytes)");
            Console.WriteLine("  2. Add File with Stream (AddFileWithStream)");
            Console.WriteLine("  3. Add Large File with Stream (AddLargeFileWithStream)");
            Console.WriteLine("  4. Add File with SaveBytes (AddFileWithSaveBytes)");
            Console.WriteLine("  5. Add File with SaveStream (AddFileWithSaveStream)");
            Console.WriteLine("  6. Add File with Chunked Upload (AddFileWithContinueUpload)");
            Console.WriteLine("  7. Update Managed Metadata Default Value (UpdateManagedMetadataDefaultValue)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    FileAddScenarios.AddFileWithBytes();
                    break;
                case "2":
                    FileAddScenarios.AddFileWithStream();
                    break;
                case "3":
                    FileAddScenarios.AddLargeFileWithStream();
                    break;
                case "4":
                    FileAddScenarios.AddFileWithSaveBytes();
                    break;
                case "5":
                    FileAddScenarios.AddFileWithSaveStream();
                    break;
                case "6":
                    FileAddScenarios.AddFileWithContinueUpload();
                    break;
                case "7":
                    FileAddScenarios.UpdateManagedMetadataDefaultValue();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void PermissionMenu()
        {
            Console.Clear();
            Console.WriteLine("=== Permission Management Scenarios ===\n");
            Console.WriteLine("Please enter site URL (relative path, e.g. /sites/yoursite):");
            var siteRelative = Console.ReadLine() ?? "/sites/simmon1456";
            
            Console.WriteLine("Please enter user login name (e.g. i:0#.f|membership|user@domain.com):");
            var userLogin = Console.ReadLine() ?? "i:0#.f|membership|simmon@baron.space";

            Console.WriteLine("\n  1. Create Default Groups");
            Console.WriteLine("  2. Get User Effective Permissions");
            Console.WriteLine("  3. Get Contributor Role Definition");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            var siteUrl = CSOM.Common.EnvConfig.GetSiteUrl(siteRelative);

            switch (choice)
            {
                case "1":
                    PermissionScenarios.CreateDefaultGroups(siteUrl, userLogin);
                    break;
                case "2":
                    var perms = PermissionScenarios.GetUserEffectivePermissions(siteUrl, userLogin);
                    Console.WriteLine($"Permission Value: {perms}");
                    break;
                case "3":
                    var role = PermissionScenarios.GetContributorRoleDefinition(siteUrl);
                    Console.WriteLine($"Role: {role.Name}");
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void TenantApiMenu()
        {
            Console.Clear();
            Console.WriteLine("=== Tenant API Scenarios ===\n");
            Console.WriteLine("  1. Get Container Types (GetSPOContainerTypes)");
            Console.WriteLine("  2. Get Containers by Application (GetSPOContainersByApplicationId)");
            Console.WriteLine("  3. Set Site Deny Add and Customize Pages (DenyAddAndCustomizePages)");
            Console.WriteLine("  4. Get Hub Site Properties (GetHubSitesProperties)");
            Console.WriteLine("  5. Test Container API (TestContainer)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    TenantApiScenarios.TestCGetSPOContainerTypes();
                    break;
                case "2":
                    TenantApiScenarios.TestGetSPOContainersByApplicationId();
                    break;
                case "3":
                    TenantApiScenarios.TestDenyAddAndCustomizePages();
                    break;
                case "4":
                    TenantApiScenarios.TestGetHubSitesProperties();
                    break;
                case "5":
                    TenantApiScenarios.TestContainer();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void ListApiMenu()
        {
            Console.Clear();
            Console.WriteLine("=== List API Scenarios ===\n");
            Console.WriteLine("  1. Print Site Title (PrintSiteTitle)");
            Console.WriteLine("  2. Set Column Default Value and Add File (SetDefaultValueAndAddFile)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    SPListApiScenarios.PrintSiteTitle();
                    break;
                case "2":
                    SetColumnDefaultValueScenarios.SetDefaultValueAndAddFile();
                    break;
                case "3":
                    SPListApiScenarios.UpdateListParser();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void ContainerMenu()
        {
            Console.Clear();
            Console.WriteLine("=== Container Scenarios ===\n");
            Console.WriteLine("  1. Dump Containers by Application ID (DumpContainersByApplicationId)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    ContainerScenarios.DumpContainersByApplicationId();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void OtherScenariosMenu()
        {
            Console.Clear();
            Console.WriteLine("=== Other Scenarios ===\n");
            Console.WriteLine("  1. Set Compliance Tag on Bulk Items (SetComplianceTagOnBulkItems)");
            Console.WriteLine("  2. Update Conflict Test (UpdateConflict)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    ComplianceTagScenarios.DemoSetComplianceTagOnBulkItems();
                    break;
                case "2":
                    UpdateConflictScenarios.Run();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void ViewMenu()
        {
            Console.Clear();
            Console.WriteLine("=== View Operations Scenarios ===\n");
            Console.WriteLine("  1. Test View and View Fields (TestViewAndViewFields)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    ViewScenarios.TestViewAndViewFields();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void ExceptionHandlingMenu()
        {
            Console.Clear();
            Console.WriteLine("=== Exception Handling Scenarios ===\n");
            Console.WriteLine("  1. Test Try/Catch Folder Creation (TestTryCatchFolderCreation)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    ExceptionHandlingScopeScenarios.TestTryCatchFolderCreation();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void CamlQueryMenu()
        {
            Console.Clear();
            Console.WriteLine("=== CAML Query Scenarios ===\n");
            Console.WriteLine("  1. Basic CAML Query (BasicCamlQuery)");
            Console.WriteLine("  2. Paginated CAML Query (PaginatedCamlQuery)");
            Console.WriteLine("  3. Create All Items Query (CreateAllItemsQuery)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    CamlQueryScenarios.BasicCamlQuery();
                    break;
                case "2":
                    CamlQueryScenarios.PaginatedCamlQuery();
                    break;
                case "3":
                    CamlQueryScenarios.CreateAllItemsQuery();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void TimeZoneMenu()
        {
            Console.Clear();
            Console.WriteLine("=== TimeZone Test Scenarios ===\n");
            Console.WriteLine("  1. Test Document TimeZone (TestClientAPI_Document)");
            Console.WriteLine("  2. Test List Item TimeZone (TestClientAPI_ListItem)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    TimeZoneScenarios.TestClientAPI_Document();
                    break;
                case "2":
                    TimeZoneScenarios.TestClientAPI_ListItem();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void TaxonomyMenu()
        {
            Console.Clear();
            Console.WriteLine("=== Managed Metadata Scenarios ===\n");
            Console.WriteLine("  1. Create Group, TermSet and Terms (CreateGroupTermSetAndTerms)");
            Console.WriteLine("  2. List Terms in TermSet (ListTermsInTermSet)");
            Console.WriteLine("  3. Get TermSet by Name (GetTermSetByName)");
            Console.WriteLine("  4. Create Managed Metadata Field (CreateManagedMetadataField)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    TaxonomyScenarios.CreateGroupTermSetAndTerms();
                    break;
                case "2":
                    TaxonomyScenarios.ListTermsInTermSet();
                    break;
                case "3":
                    TaxonomyScenarios.GetTermSetByName();
                    break;
                case "4":
                    TaxonomyScenarios.CreateManagedMetadataField();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }

        static void WebPropertiesMenu()
        {
            Console.Clear();
            Console.WriteLine("=== Web Properties Scenarios ===\n");
            Console.WriteLine("  1. Update Web AllProperties (UpdateWebAllProperties)");
            Console.WriteLine("  2. List All Web AllProperties (ListWebAllProperties)");
            Console.WriteLine("  3. Toggle Deny Add and Customize Pages (ToggleDenyAddAndCustomizePages)");
            Console.WriteLine("  4. Decode Index Property Keys (DecodeIndexPropertyKeys)");
            Console.WriteLine("  0. Return to Main Menu");
            Console.Write("\nPlease select: ");

            var choice = Console.ReadLine();
            Console.WriteLine();

            switch (choice)
            {
                case "1":
                    UpdateWebPropertiesScenarios.UpdateWebAllProperties();
                    break;
                case "2":
                    UpdateWebPropertiesScenarios.ListWebAllProperties();
                    break;
                case "3":
                    UpdateWebPropertiesScenarios.ToggleDenyAddAndCustomizePages();
                    break;
                case "4":
                    UpdateWebPropertiesScenarios.DecodeIndexPropertyKeys();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid selection");
                    break;
            }
        }
    }
}
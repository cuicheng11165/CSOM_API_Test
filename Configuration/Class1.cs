﻿using System.Net;
using System.Runtime.CompilerServices;

namespace CSOM.Common
{
    public class EnvConfig
    {
        static EnvConfig()
        {
            string configDir = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "..", "..", "..", "..", "Config"
            );
            configDir = Path.GetFullPath(configDir);

            try
            {
                HostName = File.ReadAllText(Path.Combine(configDir, "HostName.txt"));
            }
            catch (Exception ex)
            {
                HostName = string.Empty;
                Console.WriteLine($"Error reading HostName.txt: {ex.Message}");
            }

            try
            {
                Authorization = File.ReadAllText(Path.Combine(configDir, "Authorization.txt"));
            }
            catch (Exception ex)
            {
                Authorization = string.Empty;
                Console.WriteLine($"Error reading Authorization.txt: {ex.Message}");
            }

            try
            {
                ClientId = File.ReadAllText(Path.Combine(configDir, "ClientId.txt"));
            }
            catch (Exception ex)
            {
                ClientId = string.Empty;
                Console.WriteLine($"Error reading ClientId.txt: {ex.Message}");
            }

            try
            {
                TenantId = File.ReadAllText(Path.Combine(configDir, "TenantId.txt"));
            }
            catch (Exception ex)
            {
                TenantId = string.Empty;
                Console.WriteLine($"Error reading TenantId.txt: {ex.Message}");
            }

            try
            {
                UserName = File.ReadAllText(Path.Combine(configDir, "UserName.txt"));
            }
            catch (Exception ex)
            {
                UserName = string.Empty;
                Console.WriteLine($"Error reading UserName.txt: {ex.Message}");
            }

            try
            {
                Password = File.ReadAllText(Path.Combine(configDir, "Password.txt"));
            }
            catch (Exception ex)
            {
                Password = string.Empty;
                Console.WriteLine($"Error reading Password.txt: {ex.Message}");
            }

            try
            {
                CertificateThumbprint = File.ReadAllText(Path.Combine(configDir, "CertificateThumbprint.txt"));
            }
            catch (Exception ex)
            {
                CertificateThumbprint = string.Empty;
                Console.WriteLine($"Error reading CertificateThumbprint.txt: {ex.Message}");
            }
        }

        public static string HostName { get; set; }

        public static String Authorization { set; get; }

        public static String ClientId { set; get; }

        public static String TenantId { set; get; }
        public static String UserName { set; get; }
        public static String Password { set; get; }

        public static String CertificateThumbprint { set; get; }


        public static string GetSiteUrl(string siteRelativeUrl)
        {
            return $"https://{HostName}/{siteRelativeUrl.TrimStart('/')}";
        }

        public static string GetAdminCenterUrl()
        {
            return "https://" + HostName.Replace(".sharepoint.com", "-admin.sharepoint.com");
        }

        public static string GetToken()
        {
            return Authorization;
        }
    }

}
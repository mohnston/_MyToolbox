using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Runtime.InteropServices;

namespace Westmark
{
    public class INIWrapper
    {
        public INIWrapper()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [DllImport("KERNEL32.DLL",
          EntryPoint = "GetPrivateProfileString")]
        protected internal static extern int
     GetPrivateProfileString(string lpAppName,
        string lpKeyName, string lpDefault,
        StringBuilder lpReturnedString, int nSize,
        string lpFileName);

        [DllImport("KERNEL32.DLL")]
        protected internal static extern int
     GetPrivateProfileInt(string lpAppName,
        string lpKeyName, int iDefault,
        string lpFileName);

        [DllImport("KERNEL32.DLL",
          EntryPoint = "WritePrivateProfileString")]
        protected internal static extern bool
     WritePrivateProfileString(string lpAppName,
        string lpKeyName, string lpString,
        string lpFileName);

        [DllImport("KERNEL32.DLL",
          EntryPoint = "GetPrivateProfileSection")]
        protected internal static extern int
     GetPrivateProfileSection(string lpAppName,
        byte[] lpReturnedString, int nSize,
        string lpFileName);

        [DllImport("KERNEL32.DLL",
          EntryPoint = "WritePrivateProfileSection")]
        protected internal static extern bool
      WritePrivateProfileSection(string lpAppName,
        byte[] data, string lpFileName);

        [DllImport("KERNEL32.DLL",
          EntryPoint = "GetPrivateProfileSectionNames")]
        protected internal static extern int
     GetPrivateProfileSectionNames(
       byte[] lpReturnedString,
       int nSize, string lpFileName);

        public static String GetINIValue(String filename,
           String section, String key)
        {
            StringBuilder buffer = new StringBuilder(256);
            string sDefault = "";
            if (GetPrivateProfileString(section, key, sDefault,
               buffer, buffer.Capacity, filename) != 0)
            {
                return buffer.ToString();
            }
            else
            {
                return null;
            }
        }

        public static bool WriteINIValue(String filename,
           String section, String key, String sValue)
        {
            return WritePrivateProfileString(section, key,
               sValue, filename);
        }

        public static int GetINIInt(String filename,
           String section, String key)
        {
            int iDefault = -1;
            return GetPrivateProfileInt(section, key,
               iDefault, filename);
        }

        public static StringCollection GetINISection(
           String filename, String section)
        {
            StringCollection items = new StringCollection();
            byte[] buffer = new byte[32768];
            int bufLen = 0;
            bufLen = GetPrivateProfileSection(section, buffer,
               buffer.GetUpperBound(0), filename);
            if (bufLen > 0)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < bufLen; i++)
                {
                    if (buffer[i] != 0)
                    {
                        sb.Append((char)buffer[i]);
                    }
                    else
                    {
                        if (sb.Length > 0)
                        {
                            items.Add(sb.ToString());
                            sb = new StringBuilder();
                        }
                    }
                }
            }
            return items;
        }
        public static bool WriteINISection(string filename, string
           section, StringCollection items)
        {
            byte[] b = new byte[32768];
            int j = 0;
            foreach (string s in items)
            {
                ASCIIEncoding.ASCII.GetBytes(s, 0, s.Length, b, j);
                j += s.Length;
                b[j] = 0;
                j += 1;
            }
            b[j] = 0;
            return WritePrivateProfileSection(section, b, filename);
        }

        public static StringCollection GetINISectionNames(
           String filename)
        {
            StringCollection sections = new StringCollection();
            byte[] buffer = new byte[32768];
            int bufLen = 0;
            bufLen = GetPrivateProfileSectionNames(buffer,
               buffer.GetUpperBound(0), filename);
            if (bufLen > 0)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < bufLen; i++)
                {
                    if (buffer[i] != 0)
                    {
                        sb.Append((char)buffer[i]);
                    }
                    else
                    {
                        if (sb.Length > 0)
                        {
                            sections.Add(sb.ToString());
                            sb = new StringBuilder();
                        }
                    }
                }
            }
            return sections;
        }


    }
}

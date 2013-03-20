using System;
using pgDBTools;
using System.Net.Mail;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
//using iTextSharp.text;
//using iTextSharp.text.pdf;

namespace Westmark
{
	class HItem
	{
		public HItem()
		{
		}
		public HItem(string inventoryNumber)
		{
			GetModelInfoFromInventoryNumber(inventoryNumber);
		}
		public HItem(int HardwareID)
		{
			pgDB odb = new pgDB();
			string sql = "SELECT " +
				"manufacturer, " +
				"model, " +
				"wm_hardware_number, " +
				"wm_model_number, " +
				"description, " +
				"finish, " +
				"size, " +
				"is_valid, " +
				"is_stock, " +
				"image_path, " +
				"cutsheet_path, " +
				"category, " +
				"uom " +
				"FROM hardware_models " +
				"WHERE hardware_model_id = " + HardwareID.ToString();
			DataTable dt = odb.pgDataTable("spec", sql);
			if (dt.Rows.Count > 0)
			{
				this.Description = dt.Rows[0]["description"].ToString();
				this.Finish = dt.Rows[0]["finish"].ToString();
				this.InventoryNumber = dt.Rows[0]["wm_hardware_number"].ToString();
				this.IsValid = (bool)dt.Rows[0]["is_valid"];
				this.ListingNumber = dt.Rows[0]["wm_model_number"].ToString();
				this.ManModel = dt.Rows[0]["model"].ToString();
				this.Manufacturer = dt.Rows[0]["manufacturer"].ToString();
				this.Size = dt.Rows[0]["size"].ToString();
				this.UnitOfMeasure = dt.Rows[0]["uom"].ToString();
				this.IsStock = (bool)dt.Rows[0]["is_stock"];
				this.Category = dt.Rows[0]["category"].ToString();
				this.ImagePath = dt.Rows[0]["image_path"].ToString();
				this.CutsheetPath = dt.Rows[0]["cutsheet_path"].ToString();

			}

		}

		private string _Manufacturer = string.Empty;
		private string _ManModel = string.Empty;
		private string _InventoryNumber = string.Empty;
		private string _ListingNumber = string.Empty;
		private string _Description = string.Empty;
		private string _Finish = string.Empty;
		private string _Size = string.Empty;
		private bool _IsValid = false;
		private string _UnitOfMeasure = string.Empty;

		private string _ImagePath = string.Empty;
		private string _CutsheetPath;
		private bool _IsStock;
		private string _Category;

		public string Category
		{
			get { return _Category; }
			set { _Category = value; }
		}

		public bool IsStock
		{
			get { return _IsStock; }
			set { _IsStock = value; }
		}

		public string CutsheetPath
		{
			get { return _CutsheetPath; }
			set { _CutsheetPath = value; }
		}

		public string ImagePath
		{
			get { return _ImagePath; }
			set { _ImagePath = value; }
		}


		/*
			hardware_model_id integer NOT NULL DEFAULT nextval(('public.hardware_models_hardware_model_id_seq'::text)::regclass),
			manufacturer text,
			model text,
			description text,
			finish text,
			list_order integer DEFAULT 0,
			image_path text,
			provider text,
			wm_model_number text,
			wm_hardware_number text,
			wm_description text,
			psi_value double precision,
			uom text,
			size text,
			cutsheet_path text,
			is_valid boolean DEFAULT true,
			item_tag_string text,
			lead_days integer,
			block_path text,
			is_stock boolean DEFAULT false,
			category text,
			born_on timestamp without time zone DEFAULT now(),
			added_by text,
			updated_by text,
			updated_on timestamp without time zone DEFAULT now(),
			spec_props text,
			print_count integer DEFAULT 0,
		 */

		public string UnitOfMeasure
		{
			get { return _UnitOfMeasure; }
			set { _UnitOfMeasure = value; }
		}

		public bool IsValid
		{
			get { return _IsValid; }
			set { _IsValid = value; }
		}

		public string Size
		{
			get { return _Size; }
			set { _Size = value; }
		}

		public string Finish
		{
			get { return _Finish; }
			set { _Finish = value; }
		}

		public string Description
		{
			get { return _Description; }
			set { _Description = value; }
		}

		public string ManModel
		{
			get { return _ManModel; }
			set { _ManModel = value; }
		}

		public string ListingNumber
		{
			get { return _ListingNumber; }
			set { _ListingNumber = value; }
		}

		public string InventoryNumber
		{
			get { return _InventoryNumber; }
			set { _InventoryNumber = value; }
		}

		public string Manufacturer
		{
			get { return _Manufacturer; }
			set { _Manufacturer = value; }
		}

		public string GetModelInfoFromInventoryNumber(string inventoryNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT " +
				"manufacturer, " +
				"model, " +
				"wm_hardware_number, " +
				"wm_model_number, " +
				"description, " +
				"finish, " + 
				"size, " +
				"is_valid, " +
				"is_stock, " +
				"image_path, " +
				"cutsheet_path, " +
				"category, " +
				"uom " +
				"FROM hardware_models " +
				"WHERE wm_hardware_number = '" + odb.PrepForSQL(inventoryNumber) + "'";
			DataTable dt = odb.pgDataTable("spec", sql);
			if (dt.Rows.Count > 0)
			{
				this.Description = dt.Rows[0]["description"].ToString();
				this.Finish = dt.Rows[0]["finish"].ToString();
				this.InventoryNumber = dt.Rows[0]["wm_hardware_number"].ToString();
				this.IsValid = (bool)dt.Rows[0]["is_valid"];
				this.ListingNumber = dt.Rows[0]["wm_model_number"].ToString();
				this.ManModel = dt.Rows[0]["model"].ToString();
				this.Manufacturer = dt.Rows[0]["manufacturer"].ToString();
				this.Size = dt.Rows[0]["size"].ToString();
				this.UnitOfMeasure = dt.Rows[0]["uom"].ToString(); 
				this.IsStock = (bool)dt.Rows[0]["is_stock"];
				this.Category = dt.Rows[0]["category"].ToString();
				this.ImagePath = dt.Rows[0]["image_path"].ToString();
				this.CutsheetPath = dt.Rows[0]["cutsheet_path"].ToString();

			}
			return res;
		}

		internal bool TryGetStockItem(string testString)
		{
			bool res = false;
			testString = testString.Trim();
			if (testString.Length > 0)
			{
				pgDB odb = new pgDB();
				string sql = "SELECT * " +
					"FROM hardware_models " +
					"WHERE wm_hardware_number = '" + testString + "'";
				DataTable dt = odb.pgDataTable("spec", sql);
				if (dt.Rows.Count < 1)
				{
					// Look for and use more detailed information
					int bkl = testString.LastIndexOf("[");
					int bkr = testString.LastIndexOf("]");
					if (bkr > (bkl + 2))
					{
						string hmodel = testString.Substring(bkl + 1, (bkr - bkl) - 1);
						sql = "SELECT * " +
							"FROM hardware_models " +
							"WHERE wm_hardware_number = '" + hmodel + "'";
						dt = odb.pgDataTable("spec", sql);
					}
				}

				if (dt.Rows.Count > 0)
				{
					res = true;
					// fill in the rest of the info
					this.Description = dt.Rows[0]["description"].ToString();
					this.Finish = dt.Rows[0]["finish"].ToString();
					this.InventoryNumber = dt.Rows[0]["wm_hardware_number"].ToString();
					this.IsValid = (bool)dt.Rows[0]["is_valid"];
					this.ListingNumber = dt.Rows[0]["wm_model_number"].ToString();
					this.ManModel = dt.Rows[0]["model"].ToString();
					this.Manufacturer = dt.Rows[0]["manufacturer"].ToString();
					this.Size = dt.Rows[0]["size"].ToString();
					this.UnitOfMeasure = dt.Rows[0]["uom"].ToString();
				}
			}
			return res;
		}

		internal static string GetPost(string itemno)
		{
			string res = string.Empty;
			int ires = 0;
			if (itemno.Length > 1)
			{
				string lastChar = itemno.Substring(itemno.Length - 1); 
				if (int.TryParse(lastChar, out ires))
				{

				}
				else
				{
					res = lastChar;
				}
			}
			return res;
		}

		internal static int GetNum(string itemno)
		{
			int res = 0;
			int itemp = 0;
			if (itemno.Length > 1)
			{
				string firstChar = itemno.Substring(0, 1);
				if (int.TryParse(firstChar,out itemp))
				{
					// first char is a number - BAD NEWS
				}
				else
				{
					if (GetPost(itemno).Length == 1)
					{
						itemno = itemno.Substring(0, itemno.Length - 1);
					}
					if (int.TryParse(itemno.Substring(1),out res))
					{
						if (firstChar.Equals("G"))
						{
							res = res + 1000000;
						}
					}
				}
			}
			return res;
		}
	}

	class Laminate
	{
		public Laminate()
		{
		}
		public Laminate(
			string jobNumber,
			string plNumber)
		{
			pgDB odb = new pgDB();
			string sql = "SELECT pl_color_plus('" + jobNumber + "', '" + plNumber + "')";
			string res = odb.pgValue("eng", sql);

			if (res.Trim().Length > 0)
			{
				string[] sep = new string[] { "::" };
				string[] ares = res.Split(sep, StringSplitOptions.None);
				if (ares.GetUpperBound(0) >= 0)
				{
					this.Manufacturer = ares[0].Trim();
				}
				if (ares.GetUpperBound(0) > 0)
				{
					this.Model = ares[1].Trim();
				}
				if (ares.GetUpperBound(0) > 1)
				{
					this.Description = ares[2].Trim();
				}
			}

		}

		public void GetLaminateFromItem(
			string jobNumber,
			string itemNumber)
		{
			pgDB odb = new pgDB();
			DataTable dt = new DataTable();
			string sql = string.Empty;
			string man = string.Empty;
			string model = string.Empty;
			if (itemNumber.StartsWith("H"))
			{
				// Try the hitemlist table first
				bool found = false;
				sql = "SELECT manufacturer, model, description " +
					"FROM hitemlist " +
					"WHERE job_number = " + jobNumber + " " +
					"AND itemno = '" + itemNumber + "'";
				dt = odb.pgDataTable("eng", sql);
				if (dt.Rows.Count == 1)
				{
					man = dt.Rows[0]["manufacturer"].ToString().Trim();
					model = dt.Rows[0]["model"].ToString().Trim();
					if ((man.Length > 0) && (model.Length > 0))
					{
						found = true;
						this.Manufacturer = man;
						this.Model = model;
						this.Description = dt.Rows[0]["description"].ToString().Trim();
					}
				}
				// if nothing then try the itemlist table (color)
				if (!found)
				{
					sql = "SELECT notes, color " +
						"FROM itemlist " +
						"WHERE job_number = " + jobNumber + " " +
						"AND itemno = '" + itemNumber + "'";
					dt = odb.pgDataTable("eng", sql);
					if (dt.Rows.Count == 1)
					{
						string color = dt.Rows[0]["color"].ToString().Trim();
						if (color.Length > 3)
						{
							GetLaminateFromPLNumber(jobNumber, color);
						}
						else
						{
							// try to pull the color from the notes
							string snote = dt.Rows[0]["notes"].ToString();
							int plLoc = snote.IndexOf("PL-");
							if (plLoc >= 0)
							{
								string plno = snote.Substring(plLoc, 5);
								GetLaminateFromPLNumber(jobNumber, plno);
							}
						}
					}
				}
			}
			else // not an H item
			{
				sql = "SELECT notes, color " +
				   "FROM itemlist " +
				   "WHERE job_number = " + jobNumber + " " +
				   "AND itemno = '" + itemNumber + "'";
				dt = odb.pgDataTable("eng", sql);
				if (dt.Rows.Count == 1)
				{
					string color = dt.Rows[0]["color"].ToString().Trim();
					if (color.Length > 3)
					{
						GetLaminateFromPLNumber(jobNumber, color);
					}
				}
			}
		}

		public void GetLaminateFromPLNumber(
			string jobNumber,
			string plNumber)
		{
			pgDB odb = new pgDB();
			string sql = "SELECT pl_color_plus('" + jobNumber + "', '" + plNumber + "')";
			string res = odb.pgValue("eng", sql);
			
			if (res.Trim().Length > 0)
			{
				string[] sep = new string[] { "::" };
				string[] ares = res.Split(sep, StringSplitOptions.None);
				if (ares.GetUpperBound(0) >= 0)
				{
					this.Manufacturer = ares[0].Trim();
				}
				if (ares.GetUpperBound(0) > 0)
				{
					this.Model = ares[1].Trim();
				}
				if (ares.GetUpperBound(0) > 1)
				{
					this.Description = ares[2].Trim();
				}
			}

		}
		
		private string _Manufacturer = string.Empty;
		private string _Model = string.Empty;
		private string _Description = string.Empty;

		public string Manufacturer 
		{
			get { return _Manufacturer; }
			set { _Manufacturer = value; }
		}
		public string Model 
		{
			get { return _Model; }
			set { _Model = value; }
		}
		public string Description
		{
			get { return _Description; }
			set { _Description = value; }
		}

		internal static string GetSheetSizeFromDesc(string inDesc)
		{
			string res = string.Empty;
			string digs = "0123456789";
			inDesc = inDesc.ToUpper();
			inDesc = inDesc.Replace("'", "");
			inDesc = inDesc.Replace("\"", "");
			int xloc = inDesc.IndexOf(" X ");
			if (xloc >= 1)
			{
				string before = inDesc.Substring(0, xloc).Trim();
				string after = inDesc.Substring(xloc + 2).Trim();
				int idx = before.Length - 1;
				while (idx > 0 && digs.Contains(before.Substring(idx,1)))
				{
					idx--;   
				}
				before = before.Substring(idx);
				idx = 0;
				while (idx < after.Length && digs.Contains(after.Substring(idx,1)))
				{
					idx++;
				}
				after = after.Substring(0, idx + 1);
				res = before + " x " + after;
			}
			return res;
		}
	}
	
	class Employee
	{
		public Employee()
		{ }

		public static string UserName()
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT firstname " +
				"FROM employees " +
				"WHERE login = '" + Environment.UserName + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}
		public static string UserEmailInternal()
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT emailinternal " +
				"FROM employees " +
				"WHERE login = '" + Environment.UserName + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}
		public static string UserEmailExternal()
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT emailexternal " +
				"FROM employees " +
				"WHERE login = '" + Environment.UserName + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}
		public static string UserPhoneExtension()
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT extension " +
				"FROM employees " +
				"WHERE login = '" + Environment.UserName + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}


		internal static string EngineerInitials(string UserName)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT initials " +
				"FROM employees " +
				"WHERE login = '" + UserName + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}
		internal static string EngineerInitials()
		{
			return EngineerInitials(Environment.UserName);
		}

	}

	class Files
	{
		public Files()
		{ }
		public static string SettingFilePath = @"C:\Program Files\_Westmark\WestmarkSettings.xml";
		static string SettingFileDir = @"C:\Program Files\_Westmark";
		private static string ColSetting = "Setting";
		private static string ColValue = "Value";
		
		public partial class FileItem
		{
			private string _FullPath = string.Empty;
			private string _ShortName = string.Empty;

			public FileItem(string fullPath)
			{
				this._FullPath = fullPath;
				this._ShortName = FileNameFromFullPath(fullPath, true);
			}

			public string FullPath
			{
				get
				{
					return _FullPath;
				}
			}
			public string ShortName
			{
				get
				{
					return _ShortName;
				}
			}
		}

		public enum SettingCategory
		{
			Global,
			AutoCAD,
			MTO
		}
		static string[] TableNames = { "Global", "AutoCAD", "MTO" };
		public static string CleanedFileName(string dirtyFileName)
		{
			string res = dirtyFileName;
			string[] badChars = {"~", "`", "!", "@", "#",
									"$", "%", "^", "&", "*", "|",
									"}", "{", "\"", "\'", ";", ":",
									".", ",", "<", ">", "?", "/"};
			
			for (int i = 0; i <= badChars.GetUpperBound(0); i++)
			{
				res = res.Replace(badChars[i], "");
			}
			return res;
		}

		public static string FileNameFromFullPath(string fullFilePath)
		{
			return FileNameFromFullPath(fullFilePath, false);
		}
		public static string FileNameFromFullPath(string fullFilePath, bool RemoveExtension)
		{
			string res = string.Empty;
			if (File.Exists(fullFilePath))
			{
				FileInfo fi = new FileInfo(fullFilePath);
				res = fi.Name;
				if (RemoveExtension)
				{
					res = res.Replace(fi.Extension, "");
				}
			}
			return res;
		}

		public static bool WriteWestmarkSetting(
			SettingCategory Category,
			string SettingName,
			string SettingValue)
		{
			bool bres = false;
			// Does Westmark directory exist?
			if (!Directory.Exists(SettingFileDir))
			{
				// if not - Does the Program Files dir exist?
				if (!Directory.Exists(@"C:\Program Files"))
				{
					// if not - return false - There is a real problem
					return bres;
				}
				// Create the Westmark settings directory
				Directory.CreateDirectory(SettingFileDir);
			}
			// Does Westmark file exist?
			DataSet ds = new DataSet();
			if (File.Exists(SettingFilePath))
			{
				// open the settings file (DataSet)
				ds.ReadXml(SettingFilePath);
			}

			// Does the table (category) exist?
			string tblName = TableNames[(int)Category];
			DataTable dt;
			if (ds.Tables.Contains(tblName))
			{
				dt = ds.Tables[tblName];
			}
			else
			{
				dt = NewSettingsTable(tblName);
				ds.Tables.Add(dt);
			}

			// Does the setting exist?
			bool settingFound = false;
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				if (dt.Rows[i][ColSetting].ToString().Equals(SettingName))
				{
					// if yes then edit
					dt.Rows[i][ColValue] = SettingValue;
					dt.AcceptChanges();
					settingFound = true;
					break;
				}
			}
			// if no the add
			if (!settingFound)
			{
				DataRow dr = dt.NewRow();
				dr[ColSetting] = SettingName;
				dr[ColValue] = SettingValue;
				dt.Rows.Add(dr);
				dt.AcceptChanges();
			}
			// Save setting file
			ds.WriteXml(SettingFilePath, XmlWriteMode.WriteSchema);
			bres = true;
			return bres;
		}

		public static string ReadWestmarkSetting(
			SettingCategory Category,
			string SettingName)
		{
			string res = string.Empty;
			// Does Westmark file exist?
			DataSet ds = new DataSet();
			if (File.Exists(SettingFilePath))
			{
				// open the settings file (DataSet)
				ds.ReadXml(SettingFilePath);
			}
			else
			{
				return res;
			}

			// Does the table (category) exist?
			string tblName = TableNames[(int)Category];
			DataTable dt;
			if (ds.Tables.Contains(tblName))
			{
				dt = ds.Tables[tblName];
			}
			else
			{
				return res;
			}

			// Does the setting exist?
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				if (dt.Rows[i][ColSetting].ToString().Equals(SettingName))
				{
					// if yes then get the value
					res = dt.Rows[i][ColValue].ToString();
					break;
				}
			}
			return res;
		}

		private static DataTable NewSettingsTable(string tblName)
		{
			DataTable dt = new DataTable(tblName);
			DataColumn dc = new DataColumn(ColSetting, typeof(string));
			dt.Columns.Add(dc);
			dc = new DataColumn(ColValue, typeof(string));
			dt.Columns.Add(dc);
			return dt;
		}

		internal static string IncrementFileName(
			string directoryPath, 
			string fileName, 
			string extensionName)
		{
			string resPath = string.Empty;
			// if there are more than 100 increments fail
			int max = 100;

			// Directory should already exist before calling this
			// Just in case check - fail if it doesn't
			if (Directory.Exists(directoryPath))
			{
				if (extensionName.StartsWith("."))
				{
					extensionName = extensionName.Substring(1);
				}
				string testFileName = directoryPath + @"\" + fileName + "." + extensionName;
				int incr = 1;
				while (File.Exists(testFileName))
				{
					// don't let this go into an extended loop
					if (incr > max)
					{
						testFileName = resPath;
						break;
					}
					testFileName = directoryPath + @"\" + fileName + "[" + incr.ToString() + "]." + extensionName;
					incr++;
				}
				resPath = testFileName;
			}
			
			return resPath;
		}

		internal static string PageNumberOnly(string FileName)
		{
			string res = FileName.Trim();
			if (res.Contains("."))
			{
				res = res.Substring(0, res.LastIndexOf(".") - 1);
			}
			if (res.Contains("-"))
			{
				res = res.Substring(res.LastIndexOf("-") + 1);
			}
			return GeneralTools.SLZ(res);
		}

		internal static bool IsFileinUse(FileInfo file)
		{
			FileStream stream = null;
			try
			{
				stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
				
			}
			catch (IOException)
			{
				//the file is unavailable because it is:
				//still being written to
				//or being processed by another thread
				//or does not exist (has already been processed)
				return true;
			}
			finally
			{
				if (stream != null)
					stream.Close();
			}
			return false;
		}
		internal static string FileAccessIssues(string FullFilePath)
		{
			string res = string.Empty;
			try
			{
				if (!File.Exists(FullFilePath))
				{
					res = "File does not exist";
				}
				else
				{
					// File in use?
					FileInfo fi = new FileInfo(FullFilePath);
					if (IsFileinUse(fi))
					{
						res = "File is in use";
					}
					// By whom? TODO
				}
			}
			catch (Exception ex)
			{
				res = ex.Message;
			}
			return res;
		}


		public static string SYMPath = @"\\cad_svr\catalog\CABWEST\SYM";
	}

	public class SendMail
	{
		public SendMail()
		{ }

        //public static void SendLotusMail(string to, string subject, string body)
        //{
        //    SendLotusMail(to, subject, body, string.Empty);
        //}

        //public static void SendLotusMail(string to, string subject, string body, string attachmentPath)
        //{
        //    try
        //    {
        //        SmtpClient client = new SmtpClient();
        //        client.Host = "domino1.westmarkproducts.com";
        //        MailMessage message = new MailMessage();

        //        MailAddress mFrom = new MailAddress(Westmark.Employee.UserEmailExternal());
        //        message.From = mFrom;

        //        message.To.Add(new MailAddress(to));

        //        message.Subject = subject;
        //        message.Body = body;
        //        if (attachmentPath.Trim().Length > 0)
        //        {
        //            if (System.IO.File.Exists(attachmentPath))
        //            {
        //                Attachment att = new Attachment(attachmentPath);
        //                message.Attachments.Add(att);
        //            }
        //            else
        //            {
        //                message.Body = message.Body + Environment.NewLine +
        //                    "Attempt to attach file at " + attachmentPath + " failed." + Environment.NewLine +
        //                    "File does not exist.";
        //            }
        //        }

        //        client.Send(message);
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
        //        System.Windows.Forms.MessageBox.Show(
        //            "Error in " + trace.GetFrame(0).GetFileName() + Environment.NewLine +
        //            "Method: " + trace.GetFrame(0).GetMethod().Name + Environment.NewLine +
        //            "Line: " + trace.GetFrame(0).GetFileLineNumber().ToString() + Environment.NewLine +
        //            ex.Message,
        //            "Error report", System.Windows.Forms.MessageBoxButtons.OK);

        //    }
        //}
		
        //public static void SendLotusMail(string to, string subject, string body, string[] attachmentPaths)
        //{
        //    try
        //    {
        //        SmtpClient client = new SmtpClient();
        //        client.Host = "domino1.westmarkproducts.com";
        //        MailMessage message = new MailMessage();

        //        MailAddress mFrom = new MailAddress(Westmark.Employee.UserEmailExternal());
        //        message.From = mFrom;

        //        message.To.Add(new MailAddress(to));

        //        message.Subject = subject;
        //        message.Body = body;
        //        if (attachmentPaths.GetUpperBound(0) >= 0)
        //        {
        //            for (int i = 0; i <= attachmentPaths.GetUpperBound(0); i++)
        //            {
        //                string attachmentPath = attachmentPaths[i];
        //                if (attachmentPath.Trim().Length > 0)
        //                {
        //                    if (System.IO.File.Exists(attachmentPath))
        //                    {
        //                        Attachment att = new Attachment(attachmentPath);
        //                        message.Attachments.Add(att);
        //                    }
        //                    else
        //                    {
        //                        message.Body = message.Body + Environment.NewLine +
        //                            "Attempt to attach file at " + attachmentPath + " failed." + Environment.NewLine +
        //                            "File does not exist.";
        //                    }
        //                }

        //            }
        //        }
        //        client.Send(message);
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
        //        System.Windows.Forms.MessageBox.Show(
        //            "Error in " + trace.GetFrame(0).GetFileName() + Environment.NewLine +
        //            "Method: " + trace.GetFrame(0).GetMethod().Name + Environment.NewLine +
        //            "Line: " + trace.GetFrame(0).GetFileLineNumber().ToString() + Environment.NewLine +
        //            ex.Message,
        //            "Error report", System.Windows.Forms.MessageBoxButtons.OK);

        //    }
        //}

	}

	public class JobTools
	{
		public JobTools()
		{ }

		public static string GetJobName(int jobNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT name " +
				"FROM job " +
				"WHERE jobid = '" + jobNumber.ToString() + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}

		public static string GetJobRootPath(int jobNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT filepath " +
				"FROM job " +
				"WHERE jobid = '" + jobNumber.ToString() + "'";
			res = odb.pgValue("eng", sql);
			if (!Directory.Exists(res))
			{
				res = string.Empty;
			}
			return res;
		}
		public static string GetJobCADSupportPath(int jobNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT filepath " +
				"FROM job " +
				"WHERE jobid = '" + jobNumber.ToString() + "'";
			res = odb.pgValue("eng", sql);
			res += @"\" + jobNumber.ToString() + "-CADSupport";
			if (!Directory.Exists(res))
			{
				res = string.Empty;
			}
			return res;
		}
		public static string GetJobInHousePath(int jobNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT filepath " +
				"FROM job " +
				"WHERE jobid = '" + jobNumber.ToString() + "'";
			res = odb.pgValue("eng", sql);
			res += @"\" + jobNumber.ToString() + "-InHouse";
			if (!Directory.Exists(res))
			{
				res = string.Empty;
			}
			return res;
		}
		public static string GetJobCorrespondencePath(int jobNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT filepath " +
				"FROM job " +
				"WHERE jobid = '" + jobNumber.ToString() + "'";
			res = odb.pgValue("eng", sql);
			res += @"\" + jobNumber.ToString() + "-Correspondence";
			if (!Directory.Exists(res))
			{
				res = string.Empty;
			}
			return res;
		}

		internal static string GetEngineerFullName(int JobNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT employees.firstname " +
				"FROM employees " +
				"JOIN job " +
				"ON job.engid = employees.uniqueid " +
				"WHERE job.jobid = '" + JobNumber.ToString() + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}
		
		internal static string GetEngineerLogin(int JobNumber)
		{
			string res = string.Empty;
			pgDB odb = new pgDB();
			string sql = "SELECT employees.login " +
				"FROM employees " +
				"JOIN job " +
				"ON job.engid = employees.uniqueid " +
				"WHERE job.jobid = '" + JobNumber.ToString() + "'";
			res = odb.pgValue("eng", sql);
			return res;
		}

		internal static int GetJobNumberFromDwgName(string dwgName)
		{
			int jno = 0;
			if (File.Exists(dwgName))
			{
				FileInfo fi = new FileInfo(dwgName);
				string dname = fi.Name;
				if (dname.Contains("-"))
				{
					int.TryParse(dname.Substring(0, dname.IndexOf("-")), out jno);
				}
			}
			return jno;
		}

		internal static System.Collections.Generic.List<string> GetFinishList(int jno)
		{
			System.Collections.Generic.List<string> resList = new System.Collections.Generic.List<string>();
			pgDB odb = new pgDB();
			string sql = "select lamname, pl_number_only(lamname) as num " +
				"from p_profile_finish " +
				"where jobid = '" + jno.ToString() + "' " +
				"order by num";
			DataTable dt = odb.pgDataTable("eng", sql);
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				resList.Add(dt.Rows[i]["lamname"].ToString());
			}
			return resList;
		}

		internal static string ShortFilePath(int JobNumber, string FullPath)
		{
			string root = GetJobRootPath(JobNumber);
			return FullPath.Replace(root, "...");
		}

		internal static bool JobExists(int jno)
		{
			pgDB odb = new pgDB();
			string sql = "SELECT jobid " +
				"FROM job " +
				"WHERE jobid = '" + jno.ToString() + "'";
			return (odb.pgValue("eng",sql).Length > 0);
		}

		internal static string JustRoomNo(string NumDotDesc)
		{
			string res = string.Empty;
			if (NumDotDesc.Contains("·"))
			{
				res = BeforeChar(NumDotDesc, "·");
			}
			else
			{
				res = NumDotDesc;
			}
			return res;
		}
		internal static string JustRoomDesc(string NumDotDesc)
		{
			string res = string.Empty;
			res = AfterChar(NumDotDesc, "·");
			return res;
		}

		//string test1 = "hello*";
		//string test2 = "*hello";
		//string test3 = "he*llo";
		//MessageBox.Show("BeforeChar(" + test1 + ") returned " + Westmark.JobTools.BeforeChar(test1,"*") + "\n" + 
		//    "BeforeChar(" + test2 + ") returned " + Westmark.JobTools.BeforeChar(test2,"*") + "\n" + 
		//    "BeforeChar(" + test3 + ") returned " + Westmark.JobTools.BeforeChar(test3,"*") + "\n\n" +
		//    "AfterChar(" + test1 + ") returned " + Westmark.JobTools.AfterChar(test1, "*") + "\n" +
		//    "AfterChar(" + test2 + ") returned " + Westmark.JobTools.AfterChar(test2, "*") + "\n" +
		//    "AfterChar(" + test3 + ") returned " + Westmark.JobTools.AfterChar(test3, "*"))
		//    ;

		internal static string BeforeChar(string FullString, string TargetChar)
		{
			string res = string.Empty;
			int idx = FullString.IndexOf(TargetChar);
			if (idx > 0)
			{
				res = FullString.Substring(0, idx);
			}
			return res;
		}
		internal static string AfterChar(string FullString, string TargetChar)
		{
			string res = string.Empty;
			int idx = FullString.IndexOf(TargetChar);
			if ((idx > -1) && (idx < (FullString.Length - 2)))
			{
				res = FullString.Substring(idx + 1);
			}
			return res;
		}


		internal static string[] GetMTOFiles(int Job_Number)
		{
			string[] files = {string.Empty};
			if (JobExists(Job_Number))
			{
				string mtoDir=GetJobCADSupportPath(Job_Number);
				if (Directory.Exists(mtoDir))
				{
					files = Directory.GetFiles(mtoDir, "*.MTO");
				}
			}
			return files;
		}

		internal static bool JobHasMTO(int Job_Number)
		{
			// check cabinet mtos, ctop/detail, and other
			bool res = false;
			string[] cabMtos = GetMTOFiles(Job_Number);
			if (cabMtos.GetUpperBound(0) >= 0)
			{
				res = true;
			}
			else
			{
				pgDB odb = new pgDB();
				string sql = "SELECT uniqueid " +
					"FROM materiallist " +
					"WHERE job_number = " + Job_Number.ToString();
				string val = odb.pgValue("eng", sql, true);
				if (val.Length > 0)
				{
					res = true;
				}
				else
				{
					sql = "SELECT ndx " +
						"FROM " + Westmark.DBNames.HItemMTO.TableName + " " +
						"WHERE job = " + Job_Number.ToString();
					val = odb.pgValue("eng", sql, true);
					if (val.Length > 0)
					{
						res = true;
					}
				}
			}
			return res;
		}

		public static string IniDirectory = @"\\adroot\cadjobs\_DataFiles";

	}

	//public class PDFTools
	//{
		
	//    public PDFTools()
	//    { }

	//    // PDF constants
	//    public static iTextSharp.text.Font fontNormal = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.NORMAL);
	//    public static iTextSharp.text.Font fontBold = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.BOLD);
	//    public static iTextSharp.text.Font fontItalic = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.ITALIC);
	//    public static iTextSharp.text.Font fontBoldItalic = FontFactory.GetFont(FontFactory.HELVETICA, 10, (iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC));
	//    public static iTextSharp.text.Font fontNormalBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.NORMAL);
	//    public static iTextSharp.text.Font fontBoldBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.BOLD);
	//    public static iTextSharp.text.Font fontItalicBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.ITALIC);
	//    public static iTextSharp.text.Font fontBoldItalicBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, (iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC));
	//    public static string helv = FontFactory.HELVETICA;
	//    public static int norm = iTextSharp.text.Font.NORMAL;
	//    public static int bold = iTextSharp.text.Font.BOLD;
	//    public static int italic = iTextSharp.text.Font.ITALIC;

	//    public static void AddTable(ref Document document, DataTable myTable, string titleString)
	//    {
	//        AddTable(ref document, myTable, titleString, null);
	//    }
	//    public static void AddTable(ref Document document, DataTable myTable, string titleString, float[] columnWidths)
	//    {
	//        AddTable(ref document, myTable, titleString, columnWidths, false);
	//    }
		
	//    public static void AddTable(
	//        ref Document document,
	//        DataTable myTable,
	//        string titleString,
	//        float[] columnWidths,
	//        bool KeepTogether)

	//    {
	//        AddTable(
	//            ref document,
	//            myTable,
	//            titleString,
	//            columnWidths,
	//            KeepTogether,
	//            0f,
	//            0f,
	//            0f);
	//    }
		
	//    public static void AddTable(
	//        ref Document document,
	//        DataTable myTable,
	//        string titleString,
	//        float[] columnWidths,
	//        bool KeepTogether,
	//        float TitleSize,
	//        float HeaderSize,
	//        float DataSize)

	//    {
	//        string errMess = "";

	//        iTextSharp.text.Font fontNormal = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.NORMAL);
	//        iTextSharp.text.Font fontBold = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.BOLD);
	//        iTextSharp.text.Font fontItalic = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.ITALIC);
	//        iTextSharp.text.Font fontBoldItalic = FontFactory.GetFont(FontFactory.HELVETICA, 10, (iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC));
	//        iTextSharp.text.Font fontNormalBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.NORMAL);
	//        iTextSharp.text.Font fontBoldBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.BOLD);
	//        iTextSharp.text.Font fontItalicBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.ITALIC);
	//        iTextSharp.text.Font fontBoldItalicBig = FontFactory.GetFont(FontFactory.HELVETICA, 14, (iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC));
	//        string app = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;

	//        try
	//        {
	//            document.Add(new Paragraph(" "));
	//            int colCount = myTable.Columns.Count;
	//            PdfPTable tbl = new PdfPTable(colCount);
	//            string stemp = string.Empty;
	//            string[] sep = new string[]{ "\r\n" };

	//            if (columnWidths == null)
	//            {
	//                #region if there is a stored width format for this table use it

	//                try
	//                {
	//                    pgDB odb = new pgDB();
	//                    string sql = "SELECT array_to_string(column_widths, ',') " +
	//                        "FROM table_column_widths " +
	//                        "WHERE project = '" + app + "' " +
	//                        "AND table_name = '" + myTable.TableName + "'";
	//                    string res = odb.pgValue("eng", sql);
	//                    if (res.Length > 0)
	//                    {

	//                        string[] widths = res.Split(new string[] { "," }, StringSplitOptions.None);
	//                        float[] fwidths = new float[widths.GetUpperBound(0) + 1];
	//                        try
	//                        {
	//                            for (int i = 0; i <= widths.GetUpperBound(0); i++)
	//                            {
	//                                float ft = float.Parse(widths[i]);
	//                                fwidths[i] = ft;
	//                            }
	//                            columnWidths = fwidths;
	//                        }
	//                        catch (Exception ex)
	//                        {
	//                        }
	//                    }
	//                }
	//                catch (Exception ex)
	//                {

	//                }

	//                #endregion
	//                if (columnWidths == null)
	//                {
	//                    #region Autofit the columns

	//                    float[] colWidths = new float[colCount];
	//                    int[] longest = new int[colCount];
	//                    for (int i = 0; i <= longest.GetUpperBound(0); i++)
	//                    {
	//                        if (myTable.Columns[i].Caption.Length > 0)
	//                        {
	//                            longest[i] = myTable.Columns[i].Caption.Length;
	//                        }
	//                        else
	//                        {
	//                            longest[i] = myTable.Columns[i].ColumnName.Length;
	//                        }
	//                    }
	//                    for (int r = 0; r < myTable.Rows.Count; r++)
	//                    {
	//                        for (int c = 0; c < myTable.Columns.Count; c++)
	//                        {
	//                            // if there are lfcr in text only get the longest line
	//                            // cells will support multi-line
	//                            stemp = myTable.Rows[r][c].ToString().Trim();
	//                            int tlen = stemp.Length;

	//                            if (stemp.Contains("\r\n"))
	//                            {
	//                                string[] lines = stemp.Split(sep, StringSplitOptions.RemoveEmptyEntries);
	//                                tlen = 5;
	//                                for (int i = 0; i <= lines.GetUpperBound(0); i++)
	//                                {
	//                                    if (lines[i].Length > tlen)
	//                                    {
	//                                        tlen = lines[i].Length;
	//                                    }
	//                                }
	//                            }
	//                            if (tlen > longest[c])
	//                            {
	//                                longest[c] = tlen;
	//                            }
	//                        }
	//                    }
	//                    int totalChars = 0;
	//                    for (int i = 0; i <= longest.GetUpperBound(0); i++)
	//                    {
	//                        totalChars += longest[i];
	//                    }
	//                    for (int i = 0; i <= colWidths.GetUpperBound(0); i++)
	//                    {
	//                        colWidths[i] = (float)(((double)longest[i] / (double)totalChars) * 100d);
	//                    }
	//                    columnWidths = colWidths;
	//                    // tbl.SetWidths(colWidths);

	//                    #endregion
	//                }
	//            }

	//            tbl.SetWidths(columnWidths);
	//            tbl.KeepTogether = KeepTogether;
	//            tbl.HorizontalAlignment = Element.ALIGN_LEFT;
	//            // Add Title
	//            if (TitleSize > 0)
	//            {
	//                fontBoldBig.Size = TitleSize;
	//            }
	//            Phrase phrase = new Phrase(titleString, fontBoldBig);
	//            PdfPCell cell = new PdfPCell(phrase);
	//            cell.HorizontalAlignment = Element.ALIGN_CENTER;
	//            // cell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
	//            cell.BorderWidth = 0;
	//            cell.Colspan = colCount;
	//            cell.Padding = 4;
	//            tbl.AddCell(cell);
	//            // Add column headers
	//            if (HeaderSize > 0)
	//            {
	//                fontBoldItalic.Size = HeaderSize;
	//            }
	//            for (int index = 0; index < myTable.Columns.Count; index++)
	//            {
	//                phrase = new Phrase(myTable.Columns[index].ColumnName, fontBoldItalic);
	//                cell = new PdfPCell(phrase);
	//                cell.Padding = 4;
	//                cell.HorizontalAlignment = Element.ALIGN_CENTER;
	//                cell.BackgroundColor = iTextSharp.text.Color.YELLOW;
	//                cell.BorderWidthBottom = 2f;
	//                tbl.AddCell(cell);
	//            }
	//            tbl.HeaderRows = 2;
	//            bool alt = false;
	//            if (DataSize > 0)
	//            {
	//                fontNormal.Size = DataSize;
	//            }
	//            for (int i = 0; i < myTable.Rows.Count; i++)
	//            {
	//                alt = ((i % 2) > 0);
	//                for (int j = 0; j < myTable.Columns.Count; j++)
	//                {
	//                    phrase = new Phrase(myTable.Rows[i][j].ToString().Replace("High Pressure Laminate", "Unknown HPL"), fontNormal);
	//                    cell = new PdfPCell(phrase);
	//                    cell.Padding = 4;
	//                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
	//                    if (alt)
	//                    {
	//                        cell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
	//                    }
	//                    tbl.AddCell(cell);
	//                }
	//            }
	//            tbl.WidthPercentage = 100;
	//            document.Add(tbl);

	//            //// *****
	//            //// for column width debug
	//            //string cwids = string.Empty;
	//            //for (int i = 0; i <= columnWidths.GetUpperBound(0); i++)
	//            //{
	//            //    cwids += "[" + columnWidths[i].ToString() + "]   ";
	//            //}
	//            //phrase = new Phrase(cwids, fontNormal);
	//            //document.Add(phrase);
	//            //// *****
	//        }
	//        catch (DocumentException de)
	//        {
	//            errMess = de.Message;
	//        }
	//        catch (IOException ioe)
	//        {
	//            errMess = errMess + "\n" + ioe.Message;
	//        }
	//        catch (Exception ex)
	//        {
	//            errMess = errMess + "\n" + ex.Message;
	//        }
	//        if (errMess.Length > 0)
	//        {
	//            System.Windows.Forms.MessageBox.Show(errMess, "In AddTable");
	//        }
	//    }

	//    public enum CellAlignVer
	//    {
	//        Top = Element.ALIGN_TOP,
	//        Mid = Element.ALIGN_MIDDLE,
	//        Bottom=Element.ALIGN_BOTTOM
	//    }
	//    public enum CellAlignHor
	//    {
	//        Left = Element.ALIGN_LEFT,
	//        Center = Element.ALIGN_CENTER,
	//        Right = Element.ALIGN_RIGHT
	//    }

	//    internal static PdfPCell Cell(
	//        string text,
	//        float fontSize,
	//        bool isBold,
	//        bool isItalic,
	//        CellAlignHor AlignmentHor,
	//        CellAlignVer AlignmentVer,
	//        float cellPadding,
	//        Color colorBack,
	//        Color colorFore)
	//    {
	//        int iFontStyle = iTextSharp.text.Font.NORMAL;
	//        if (isBold)
	//        {
	//            iFontStyle = iTextSharp.text.Font.BOLD;
	//            if (isItalic)
	//            {
	//                iFontStyle = (iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC);
	//            }
	//        }
	//        else if (isItalic)
	//        {
	//            iFontStyle = (iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC);
	//        }

	//        iTextSharp.text.Font myFont = FontFactory.GetFont(FontFactory.HELVETICA, fontSize, iFontStyle, colorFore);
	//        Phrase ph = new Phrase();
	//        ph.Font = myFont;
	//        ph.Add(text);
	//        PdfPCell cell = new PdfPCell(ph);
	//        cell.HorizontalAlignment = (int)AlignmentHor;
	//        cell.VerticalAlignment = (int)AlignmentVer;
	//        cell.Padding = cellPadding;
	//        cell.BackgroundColor = colorBack;
	//        return cell;
	//    }

	//    internal static PdfPCell Cell(
	//        string text, 
	//        float fontSize, 
	//        CellAlignHor AlignmentHor, 
	//        CellAlignVer AlignmentVer, 
	//        float cellPadding, 
	//        Color colorBackground)
	//    {
	//        return Cell(
	//            text,
	//            fontSize,
	//            false,
	//            false,
	//            AlignmentHor,
	//            AlignmentVer,
	//            cellPadding,
	//            colorBackground,
	//            Color.BLACK);
	//    }
	//}

	public class GeneralTools
	{
		public GeneralTools()
		{ }

		public enum FractionalPrecision
		{
			fracPrec1 = 1,
			fracPrec2 = 2,
			fracPrec4 = 4,
			fracPrec8 = 8,
			fracPrec16 = 16,
			fracPrec32 = 32,
			fracPrec64 = 64,
			fracPrec128 = 128
		}

		public static string SLZ(string inputString)
		{
			string res = inputString.Trim();
			while (res.StartsWith("0") && res.Length > 1)
			{
				res = res.Substring(1);
			}
			return res;
		}
		
		public static string QuotePG(string value)
		{
			string temp = value ?? string.Empty;
			temp = temp.Trim();
			if (temp.Length == 0) return "";
			temp = temp.Replace("\\", "\\\\");
			temp = temp.Replace("\'", "\\\'");
			temp = temp.Replace("'", "\'");
			return "'" + temp + "'";
		}
		public static string Quote(string value)
		{
			string res = string.Empty;
			res = "'" + value + "'";
			return res;
		}

		public static string TrueOrFalse(object DataValueObject)
		{
			string res = "false";
			if (DataValueObject != null)
			{
				if (DataValueObject.GetType() == typeof(bool))
				{
					if ((bool)DataValueObject)
					{
						res = "true";
					}
				}
				else
				{
					if (DataValueObject.GetType() == typeof(int))
					{
						if ((int)DataValueObject > 0)
						{
							res = "true";
						}
					}
				}
			}
			return res;
		}

		public static string UnQuote(string itemno)
		{
			string res = itemno;
			if (res.StartsWith("'"))
			{
				res = res.Substring(1);
			}
			if (res.EndsWith("'"))
			{
				res = res.Substring(0, res.Length - 1);
			}
			return res;
		}

		public static bool isNumeric(string val)
		{
			try
			{
				double mydbl = Convert.ToDouble(val);
				return true;
			}
			catch { return false; }
		}

		public static bool isInt(string val)
		{
			try
			{
				int myint = Convert.ToInt32(val);
				return true;
			}
			catch { return false; }
		}

		internal static decimal StringToDecimal(string tempVal)
		{
			decimal dres = 0;
			if (tempVal.Trim().Length == 0) return dres;
			// if already an int or decimal just return the number
			if (decimal.TryParse(tempVal, out dres))
			{
				return dres;
			}

			// for some columns allow for fractional or mod input or millimeter
			// 11 1/2 = 11.5; 11-1/2 = 11.5; 3m = 3.75; 4m = 5; 25.4mm = 1; 32mm = 1.25
			tempVal = tempVal.ToLower();

			if (tempVal.EndsWith("mm"))
			{
				// if mm then convert number from millimeters to decimal inches
				if (tempVal.Length > 2)
				{
					string num = tempVal.Substring(0, tempVal.Length - 2);
					int inum = 0;
					if (int.TryParse(num, out inum))
					{
						if (inum > 0)
						{
							dres = (decimal)inum / (decimal)25.4;
						}
					}
				}
			}
			else
			{
				if (tempVal.EndsWith("m"))
				{
					// if m then conaver number from 32mm mods to decimal inches
					if (tempVal.Length > 1)
					{
						string num = tempVal.Substring(0, tempVal.Length - 1);
						int inum = 0;
						if (int.TryParse(num, out inum))
						{
							if (inum > 0)
							{
								dres = (decimal)(inum * 32) / (decimal)25.4;
							}
						}
					}

				}
				else
				{
					// if has / then convert from fraction to decimal inches
					if (tempVal.Contains("/"))
					{
						int whole = 0;
						int nom = 0;
						int den = 0;
						int slash = tempVal.IndexOf("/");
						int dash = tempVal.IndexOf("-");
						int space = tempVal.IndexOf(" ");
						if (slash > 0 && slash < (tempVal.Length - 1))
						{
							// slash is not first or last
							if (dash > space)
							{
								space = dash;
							}
							if (space == -1) // no whole number - just the fraction
							{
								if (int.TryParse(tempVal.Substring(0, slash), out nom))
								{
									if (nom > 0)
									{
										if (int.TryParse(tempVal.Substring(slash + 1), out den))
										{
											if (den > 0)
											{
												dres = ((decimal)nom / (decimal)den);
											}
										}
									}
								}
							}
							else if (space > 0 && space < (tempVal.Length - 1))
							{
								if (slash > (space + 1))
								{
									// space is not first or last
									if (int.TryParse(tempVal.Substring(0, space), out whole))
									{
										// there is a whole number
										if (int.TryParse(tempVal.Substring(space + 1, (slash - space) - 1), out nom))
										{
											if (nom > 0)
											{
												if (int.TryParse(tempVal.Substring(slash + 1), out den))
												{
													if (den > 0)
													{
														dres = (decimal)whole + ((decimal)nom / (decimal)den);
													}
												}
											}
										}
									}
								}

							}
						}

					}
				}
			}
			return dres;
		}

		internal static double StringToDouble(string tempVal)
		{
			double dres = 0;
			if (tempVal.Trim().Length == 0) return dres;
			// if already an int or decimal just return the number
			if (double.TryParse(tempVal, out dres))
			{
				return dres;
			}

			// for some columns allow for fractional or mod input or millimeter
			// 11 1/2 = 11.5; 11-1/2 = 11.5; 3m = 3.75; 4m = 5; 25.4mm = 1; 32mm = 1.25
			tempVal = tempVal.ToLower();

			if (tempVal.EndsWith("mm"))
			{
				// if mm then convert number from millimeters to double inches
				if (tempVal.Length > 2)
				{
					string num = tempVal.Substring(0, tempVal.Length - 2);
					int inum = 0;
					if (int.TryParse(num, out inum))
					{
						if (inum > 0)
						{
							dres = (double)inum / (double)25.4;
						}
					}
				}
			}
			else
			{
				if (tempVal.EndsWith("m"))
				{
					// if m then conaver number from 32mm mods to double inches
					if (tempVal.Length > 1)
					{
						string num = tempVal.Substring(0, tempVal.Length - 1);
						int inum = 0;
						if (int.TryParse(num, out inum))
						{
							if (inum > 0)
							{
								dres = (double)(inum * 32) / (double)25.4;
							}
						}
					}

				}
				else
				{
					// if has / then convert from fraction to double inches
					if (tempVal.Contains("/"))
					{
						int whole = 0;
						int nom = 0;
						int den = 0;
						int slash = tempVal.IndexOf("/");
						int dash = tempVal.IndexOf("-");
						int space = tempVal.IndexOf(" ");
						if (slash > 0 && slash < (tempVal.Length - 1))
						{
							// slash is not first or last
							if (dash > space)
							{
								space = dash;
							}
							if (space == -1) // no whole number - just the fraction
							{
								if (int.TryParse(tempVal.Substring(0, slash), out nom))
								{
									if (nom > 0)
									{
										if (int.TryParse(tempVal.Substring(slash + 1), out den))
										{
											if (den > 0)
											{
												dres = ((double)nom / (double)den);
											}
										}
									}
								}
							}
							else if (space > 0 && space < (tempVal.Length - 1))
							{
								if (slash > (space + 1))
								{
									// space is not first or last
									if (int.TryParse(tempVal.Substring(0, space), out whole))
									{
										// there is a whole number
										if (int.TryParse(tempVal.Substring(space + 1, (slash - space) - 1), out nom))
										{
											if (nom > 0)
											{
												if (int.TryParse(tempVal.Substring(slash + 1), out den))
												{
													if (den > 0)
													{
														dres = (double)whole + ((double)nom / (double)den);
													}
												}
											}
										}
									}
								}

							}
						}

					}
				}
			}
			return dres;
		}

		public static string modsToFraction(string mods)
		{
			mods = mods.Trim();
			if (mods.Length == 0) return "";
			if (isInt(mods) == false) return "";
			int mo = Convert.ToInt32(mods);
			double mmact = mo * 32;
			double inact = mmact / 25.4;
			string actheight = doubleToFraction(inact, FractionalPrecision.fracPrec64);
			return actheight;
		}

		public static string nominalToFraction(string nominch)
		{
			// return fractional value as string
			// if unable to convert then return an empty string
			if (nominch.Trim().Length < 1) return "";
			if (isInt(nominch) == false) return "";
			int nom = Convert.ToInt32(nominch);
			double mmnom = nom * 25.4;
			int mods = Convert.ToInt32(mmnom / 32);
			double mmact = mods * 32;
			double inact = mmact / 25.4;
			string actheight = doubleToFraction(inact, FractionalPrecision.fracPrec64);
			return actheight;
		}

		internal static decimal nominalToDecimal(int nomInch)
		{
			decimal resDec = 0;
			if (nomInch > 0)
			{
				double mmnom = nomInch * 25.4;
				int mods = Convert.ToInt32(mmnom / 32);
				double mmact = mods * 32;
				resDec = (decimal)(mmact / 25.4);
			}
			return resDec;
		}

		public static string doubleToFraction(double val, FractionalPrecision precision)
		{
			if (val == 0) return "";
			double dec = (val % 1);
			double whole = val - dec;
			if (whole == val) return whole.ToString();
			double prec = (double)precision;
			if (prec == 0) return whole.ToString();
			double denom = prec;
			double decfrac = 1 / prec;
			double num = Convert.ToDouble(Convert.ToInt32(dec / decfrac));
			// check for another whole unit ( 64/64, 32/32 etc.)
			if (num == denom)
			{
				whole = whole + 1;
				return whole.ToString();
			}
			if (num == 0) return whole.ToString();

			// pass back the reduced fraction
			while ((num % 2) == 0)
			{
				num = num / 2;
				denom = denom / 2;
			}
			if (whole > 0)
			{
				return whole.ToString() + " " + num.ToString() + @"/" + denom.ToString();
			}
			else
			{
				return num.ToString() + @"/" + denom.ToString();
			}

		}

		public static string DecimalToFraction(decimal val, FractionalPrecision precision)
		{
			if (val == 0) return "";
			decimal dec = (val % 1);
			decimal whole = val - dec;
			if (whole == val) return ((int)whole).ToString();
			decimal prec = (decimal)precision;
			if (prec == 0) return ((int)whole).ToString();
			decimal denom = prec;
			decimal decfrac = 1 / prec;
			decimal num = Convert.ToDecimal(Convert.ToInt32(dec / decfrac));
			// check for another whole unit ( 64/64, 32/32 etc.)
			if (num == denom)
			{
				whole = whole + 1;
				return whole.ToString();
			}
			if (num == 0) return whole.ToString();

			// pass back the reduced fraction
			while ((num % 2) == 0)
			{
				num = num / 2;
				denom = denom / 2;
			}
			if (whole > 0)
			{
				return ((int)whole).ToString() + " " + num.ToString() + @"/" + denom.ToString();
			}
			else
			{
				return num.ToString() + @"/" + denom.ToString();
			}

		}

		internal static string GetHashString(System.Collections.Hashtable hTable, object KeyValue)
		{
			string res = string.Empty;
			if (hTable.ContainsKey(KeyValue))
			{
				res = hTable[KeyValue].ToString().Trim();
			}
			return res;
		}

		internal static decimal GetHashDecimal(System.Collections.Hashtable hTable, object KeyValue)
		{
			decimal res = 0;
			if (hTable.ContainsKey(KeyValue))
			{
				string sres = hTable[KeyValue].ToString();
				decimal d;
				if (decimal.TryParse(sres, out d))
				{
					res = d;
				}
			}
			return res;
		}

		internal static int GetHashInt(System.Collections.Hashtable hTable, object KeyValue)
		{
			int res = 0;
			if (hTable.ContainsKey(KeyValue))
			{
				string sres = hTable[KeyValue].ToString();
				int i;
				if (int.TryParse(sres, out i))
				{
					res = i;
				}
			}
			return res;
		}

	}

	public class DBNames
	{
		public static partial class ItemTable
		{
			// TODO: change table name when we go live
			// public const string ItemTableName = "items";
            public const string ItemTableName = "itemlist";
			public const string ItemsUnresolvedTableName = "items_unresolved";
			public const string uniqueid = "uniqueid"; // integer NOT NULL DEFAULT nextval(('"items_uniqueid_seq"'::text)::regclass),
			public const string jobno = "jobno"; // character varying(50),
			public const string fulldwgpath = "fulldwgpath"; // character varying(150),
			public const string area = "area"; // character varying(50),
			public const string floor = "floor"; // character varying(50),
			public const string building = "building"; // character varying(50),
			public const string room = "room"; // character varying(150),
			public const string elevation = "elevation"; // character varying(50),
			public const string phase = "phase"; // character varying(50) DEFAULT 1,
			public const string cycle = "cycle"; // character varying(50),
			public const string model = "model"; // character varying(150),
			public const string w = "w"; // character varying(150),
			public const string h = "h"; // character varying(50),
			public const string d = "d"; // character varying(50),
			public const string itemno = "itemno"; // character varying(50),
			public const string notes = "notes"; // text,
			public const string finish = "finish"; // character varying(150),
			public const string mtostatus = "mtostatus"; // character varying(50),
			public const string qty = "qty"; // integer,
			public const string color = "color"; // character varying(50),
			public const string handle = "handle"; // character varying(50),
			public const string locks = "locks"; // character varying(50),
			public const string lockgroup = "lockgroup"; // character varying(50),
			public const string data4 = "data4"; // character varying(50),
			public const string data2 = "data2"; // character varying(50),
			public const string data3 = "data3"; // character varying(50),
			public const string engstatus = "engstatus"; // text,
			public const string prostatus = "prostatus"; // text,
			public const string LastEditTime = "LastEditTime"; // timestamp with time zone DEFAULT now(),
			public const string LastEditBy = "LastEditBy"; // text,
			public const string shelfqty = "shelfqty"; // smallint DEFAULT 0,
			public const string mtophase = "mtophase"; // smallint,
			public const string data1 = "data1"; // text,
			public const string shelf050 = "\"shelf-050\""; // smallint DEFAULT 0,
			public const string shelf075 = "\"shelf-075\""; // smallint DEFAULT 0,
			public const string shelf100 = "\"shelf-100\""; // smallint DEFAULT 0,
			public const string room_desc = "room_desc"; // text,
			public const string added = "added"; // timestamp without time zone DEFAULT now(),
			public const string updated = "updated"; // timestamp without time zone DEFAULT now(),
			public const string job_number = "job_number"; // integer,
			public const string on_page = "on_page"; // boolean DEFAULT true,
			public const string h_is_net = "h_is_net"; // boolean,
			public const string unit_heights = "unit_heights"; // character varying(128),
			public const string unit_widths = "unit_widths"; // character varying(64),
			public const string is_modified = "is_modified"; // boolean,
			public const string is_locked = "is_locked"; // boolean,
			public const string unit_below_bottom = "unit_below_bottom"; // character varying(12) DEFAULT 0,
			public const string door_adjust_top = "door_adjust_top"; // double precision DEFAULT 0,
			public const string door_adjust_bottom = "door_adjust_bottom"; // double precision DEFAULT 0,
			public const string added_by_mto = "added_by_mto"; // boolean DEFAULT false,
			public const string sill_placed = "sill_placed"; // boolean DEFAULT false,
			public const string room_prefix = "room_prefix"; // text,
			public const string w_num = "w_num"; // numeric(7,3) DEFAULT 0,
			public const string d_num = "d_num"; // numeric(7,3) DEFAULT 0,
			public const string h_num = "h_num"; // numeric(9,5) DEFAULT 0,
			public const string shelf025 = "\"shelf-025\""; // smallint DEFAULT 0,
			public const string shelfb025 = "\"shelfb-025\""; // smallint DEFAULT 0,
			public const string shelfb050 = "\"shelfb-050\""; // smallint DEFAULT 0,
			public const string shelfb075 = "\"shelfb-075\""; // smallint DEFAULT 0,
			public const string shelfb100 = "\"shelfb-100\""; // smallint DEFAULT 0,
			public const string spec_id = "spec_id"; // integer DEFAULT 0,
			public const string unit_body_heights = "unit_body_heights"; // text,
			public const string is_sill = "is_sill"; // boolean DEFAULT false,
			public const string mark_for_delete = "mark_for_delete"; // boolean DEFAULT false,
			public const string existing_item_id = "existing_item_id";
			public const string conflicts = "conflicts";

			internal static string FormatItemNumberForDB(string itemNumber)
			{
				string res = itemNumber;
				if (itemNumber.Length > 0)
				{
					if (Westmark.GeneralTools.isNumeric(itemNumber.Substring(itemNumber.Length - 1)))
					{
						res = res.PadLeft(5, char.Parse("0"));
					}
				}
				return res;
			}
		}

		public static partial class HItemTable
		{
			// TODO: change table name for released product
			// public const string HItemTableName = "hitems";
            public const string HItemTableName = "hitemlist";
			public const string HItemUnresolvedTableName = "hitems_unresolved";
			public const string MTO_reference = "mto_reference_ndx";
			public const string hitem_id = "hitem_id";
			public const string job_number = "job_number"; // integer,
			public const string itemno = "itemno"; // text,
			public const string quantity = "quantity"; // numeric,
			public const string unit_of_measure = "unit_of_measure"; // text,
			public const string manufacturer = "manufacturer"; // text,
			public const string model = "model"; // text,
			public const string listing_design = "listing_design"; // text,
			public const string description = "description"; // text,
			public const string size = "size"; // text,
			public const string finish = "finish"; // text,
			public const string notes = "notes"; // text,
			public const string added = "added"; // timestamp without time zone,
			public const string updated = "updated"; // timestamp without time zone,
			public const string room = "room"; // text,
			public const string elevation = "elevation"; // text,
			public const string phase = "phase"; // smallint,
			public const string fulldwgpath = "fulldwgpath"; // text,
			public const string marked_for_delete = "marked_for_delete"; // boolean DEFAULT false,
			public const string mto_note = "mto_note"; // text,
			public const string mtod = "mtod"; // timestamp without time zone,
			public const string parent = "parent"; // text,
			public const string status = "status"; // smallint DEFAULT 1, -- 0-Drawn but not ready to mto...

			public const string existing_item_id = "existing_item_id";
			public const string hitem_unresolved_id = "hitem_unresolved_id";

			public const string Conflict = "conflicts";
		}

		public static partial class HItemMTO
		{
			// TODO: Change table name
			// public const string TableName = "h_f_items_mtod_temp";
            public const string TableName = "h_f_items_mtod";
			public const string TableNameUnresolved = "h_f_items_mtod_unresolved";
			public const string conflict = "conflicts";
			public const string reference_index = "ref_ndx";
			public const string ndx = "ndx"; // integer NOT NULL DEFAULT nextval(('public.hf_item_mtod_ndx_seq'::text)::regclass),
			public const string job = "job"; //  integer,
			public const string takeoff = "takeoff"; //  smallint,
			public const string phase = "phase"; //  smallint,
			public const string item = "item"; //  text,
			public const string qty = "qty"; //  integer,
			public const string w = "w"; //  text,
			public const string h = "h"; //  text,
			public const string d = "d"; //  text,
			public const string notes = "notes"; //  text,
			public const string model = "model"; //  text,
			public const string room = "room"; //  text,
			public const string elevation = "elevation"; //  text,
			public const string locked = "locked"; //  boolean DEFAULT false,
			public const string filename = "filename"; //  text,
			public const string added = "added"; //  timestamp without time zone DEFAULT now(),
			public const string item_type = "item_type"; //  text,
			public const string item_number = "item_number"; //  integer,
			public const string item_part = "item_part"; //  text,
			public const string topurchasing = "topurchasing"; //  timestamp without time zone,
			public const string added_by_mto = "added_by_mto"; //  boolean DEFAULT false,
			public const string itemheld = "itemheld"; //  boolean DEFAULT false,
			public const string manuf = "manuf"; //  text,
			public const string mmodel = "mmodel"; //  text,
			public const string size = "size"; //  text,
			public const string finish = "finish"; //  text,
			public const string mto_notes = "mto_notes"; //  text,
			public const string parent = "parent"; //  text,
			public const string uom = "uom"; //  text DEFAULT 'EA'::text,
		}

		public static partial class CADDBErrors
		{
			public static string CADDBErrorTableName = "cad_db_errors";
			public static string type = "type"; // text,
			public static string fulldwgpath = "fulldwgpath"; // text,
			public static string itemno = "itemno"; // text,
			public static string model = "model"; // text,
			public static string room = "room"; // text,
			public static string elevation = "elevation"; // text,
			public static string jobno = "jobno"; // integer,
			public static string handle = "handle"; // text,
			public static string itemlist_id = "itemlist_id"; // bigint,
			public static string description = "description"; // text,
			public static string bad_info = "bad_info"; // text,
			public static string sql_string = "sql_string"; // text,
			public static string location = "location"; // text,
			public static string time_stamp = "time_stamp"; // timestamp without time zone DEFAULT now(),
			public static string room_desc = "room_desc"; // text,
			public static string level = "level"; // text
		}
	}
	
	public class CADNames
	{
		public static partial class MultiRoom
		{
			public const string GridBlockName = "MULTI_E";
			public const string TableTypeName = "MULTIROOM";
			public const string TableTitle = "ELEVATION";

		}

		public static partial class SillTable
		{
            public const string ItemType = "ItemType";
            public const string TableTypeName = "SILLS";
            public const string TableTitle = "SILLS";
            public const string SillType_CTOP = "CTOP";
            public const string SillType_HITEM = "HITEM";
            public const string SillType_DETAIL = "DETAIL";

			public const int ColIndexItemTag = 0;
            public const int ColIndexColor = 1;
			public const int ColIndexRoom = 2;
			public const int ColIndexNote = 3;
			public const int ColIndexWidthA = 4;
			public const int ColIndexWidthAMetric = 5;
			public const int ColIndexDepthB = 6;
			public const int ColIndexDepthBMetric = 7;
			public const int ColIndexWidthC = 8;
			public const int ColIndexDepthD = 9;

			public const string PlanEndType = "plan_end_type";
			public const string SectionType = "section_type";
			public const string FinishType = "finish_type";
			public const string FrontEdgeType = "front_edge";
			public const string Substrate = "substrate";
			public const string Installation = "installation";
			public const string FinishUnder = "finish_under";
			public const string DimDeckThick = "dim_deck_thick";
			public const string DimOverhang = "dim_overhang";
			public const string DimFrontThick = "dim_front_thick";
			public const string DimWrapJamb = "dim_wrap_jamb";
			public const string DimApronHt = "dim_apron_height";
			public const string DimApronDp = "dim_apron_depth";
			public const string DimScribeW = "dim_scribe_w";
			public const string DimScribeD = "dim_scribe_d";
			public const string Elevation = "elevation";
			public const string Room = "room";
			public const string RoomDesc = "room_desc";
			public const string ArchDetail = "arch_detail";
			public const string ArchPage = "arch_page";
            public const string TableStyleName = "WM_Standard";
        }

		public static partial class XDFields
		{
			public const string ShelfQty = "ShelfQty"; //1
			public const string CabHandle = "CabHandle"; //2AE2C
			public const string Room = "Room"; //101·SCIENCE LAB
			public const string Elev = "Elev"; //1
			public const string MyHandle = "MyHandle"; //2A4E9
			public const string FinSides = "FinSides"; //L,B
			public const string Model = "Model"; //B202
			public const string CWidth = "CWidth"; //28
			public const string CHeight = "CHeight"; //19
			public const string CDepth = "CDepth"; //17
			public const string JobNo = "JobNo"; //10205
			public const string ItemNo = "ItemNo"; //0
			public const string Phase = "Phase"; //1
			public const string Elevation = "Elevation"; //
			public const string Series = "Series"; //
			public const string HingeVisible = "HingeVisible"; //False
			public const string HingeType = "HingeType"; //
			public const string PullType = "PullType"; //
			public const string ModBreaks = "ModBreaks"; //19
			public const string LockBreaks = "LockBreaks"; //
			public const string Locks = "Locks"; //True
			public const string Modified = "Modified"; //True
			public const string Status = "Status"; //
			public const string Quantity = "Qty"; //0
			public const string Comment = "Comment"; //
			public const string FinishProfile = "FinishProfile"; //
			public const string HdwrProfile = "HdwrProfile"; //
			public const string DivQty = "DivQty"; //0
			public const string DivThick = "DivThick"; //0
			public const string FixedWidth = "FixedWidth"; //0
			public const string DwgPath = "DwgPath"; //\\CAD_SVR\CADJOBS\10205-CARDOZO MS SCI RM ADD\10205-003.DWG
			public const string BlockName = "BlockName"; //CES-B202-10205-003-0
			public const string SH025 = "SH025"; //0
			public const string SH050 = "SH050"; //0
			public const string SH075 = "SH075"; //1
			public const string SH100 = "SH100"; //0
			public const string SH025b = "SH025b"; //0
			public const string SH050b = "SH050b"; //0
			public const string SH075b = "SH075b"; //0
			public const string SH100b = "SH100b"; //0
			public const string HookQty1 = "HookQty1"; //0
			public const string BodyBreakup = "BodyBreakup"; //19
			public const string WidthBreakup = "WidthBreakup"; //14|14
			public const string EndsToFloor = "EndsToFloor"; //0
			public const string ToteSize = "ToteSize"; //Large
			public const string DoorHtAdjTop = "DoorHtAdjTop"; //0
			public const string DoorHtAdjBottom = "DoorHtAdjBottom"; //0
			public const string TagVersion = "TagVersion"; //2
			public const string TagHandle = "TagHandle"; //2A4E9


			public const string ParentNo = "ParentItemNo";
			public const string ParentHandle = "ParentItemHandle";

			public const string SpecID = "SpecID";

		}

		public static partial class ItemType
		{
			public const string CabinetType = "Cabinet";
			public const string CtopType = "Countertop";
			public const string DetailType = "Detail";
			public const string MetalType = "Metal";
			public const string ToeType = "ToeKick";
			public const string OtherType = "Other";

			public const string DetailModel = "DETAIL";
			public const string MetalModel = "METAL";
			public const string CtopModel = "TOP";
			public const string SillModel = "SILL";
		}

		public static partial class DetailBar
		{
			public const string BLOCK_NAME = "s_sec_dyn";
			public const string SECTION_NUMBER = "B_SECNO";
			public const string ARCH_PAGE = "B_ARCPAGENO";
			public const string ARCH_REF = "B_ARCELEVNO";
			public const string PAGE_NUMBER = "B_PAGENO";
			public const string SECTION_DESC = "B_SECDEC";
			public const string SCALE = "B_SCALE";
		}

		public static partial class ElevBar
		{
			public const string BLOCK_NAME = "s_elev_dyn";
			public const string ELEVATION_NUMBER = "B_ELEVNO";
			public const string ROOM = "B_ROOMLO";
			public const string ARCH_PAGE = "B_ARCPAGENO";
			public const string ARCH_REF = "B_ARCELEVNO";
			public const string COMPASS_DIR = "B_DIR";
			public const string STATUS = "B_STATUS";
			public const string PHASE = "B_PHASE";
			public const string PHASE_PREVIOUS_A = "B_PREV_PHASE_A";
			public const string PHASE_PREVIOUS_B = "B_PREV_PHASE_B";
			public const string DYN_DISTANCE = "Distance";
		}
		public static partial class Title
		{
			public static string BLOCK_NAME = "TITLE";
			public static string PAGE_NUMBER = "B_PAGENO";
			public static string ELEVATION_RANGE = "B_ELEVNO";
			public static string ELEVATION_OR_PLAN = "ELEV_OR_PLAN";
			public static string DRAWN_BY = "B_DRAWN";
			public static string DATE = "B_DATE";
			public static string SCALE = "B_DWGSCALE";
			public static string JOB_NUMBER = "B_JOBNO";
			public static string JOB_NAME = "B_JOBNAME";
			public static string ADDRESS = "B_ADDRESS";
			public static string CITY_STATE_ZIP = "B_CITY";
			public static string REVISION_1 = "REVISION_1";
			public static string REVISION_2 = "REVISION_2";
			public static string REVISION_3 = "REVISION_3";
			public static string REVISION_4 = "REVISION_4";
			public static string REVISION_5 = "REVISION_5";
			public static string REVISION_6 = "REVISION_6";
			public static string REVISION_7 = "REVISION_7";
			public static string REVISION_8 = "REVISION_8";

		}
		public static partial class Panel
		{
			public static string BLOCK_NAME = "dynPanel";
			public static string ITEM_NUMBER = "ITEMNO";
			public static string ATT_EDGE_LETTERS = "EDGE_LETTERS";
			public static string COLOR = "COLOR";
			public static string DYN_SIZE_WITH_GRAIN_CUT = "CutWithGrain";
			public static string DYN_SIZE_ACROSS_GRAIN_CUT = "CutCrossGrain";
			public static string DYN_SIZE_WITH_GRAIN_SNAP = "SnapHeight";
			public static string DYN_SIZE_ACROSS_GRAIN_SNAP = "SnapWidth";
			public static string DYN_VISIBILITY = "Visibility1";
			public static string DYN_FACE_CUT_TOP = "CutToFace_Top";
			public static string DYN_FACE_SNAP_TOP = "SnapToFace_Top";
			public static string DYN_FACE_CUT_BOTTOM = "CutToFace_Bottom";
			public static string DYN_FACE_SNAP_BOTTOM = "SnapToFace_Bottom";
			public static string DYN_FACE_CUT_LEFT = "CutToFace_Left";
			public static string DYN_FACE_SNAP_LEFT = "SnapToFace_Left";
			public static string DYN_FACE_CUT_RIGHT = "CutToFace_Right";
			public static string DYN_FACE_SNAP_RIGHT = "SnapToFace_Right";
			public static string DYN_GRAIN_ROTATION = "GrainRotation";

			public static string SPEC_BLOCK_SOLID_WOOD = "SPECBOX_WOOD";
			public static string SPEC_BLOCK_PANELING = "SPECBOX_PANELING";
			public static string SPEC_BLOCK_SUBSTRATE_PANELING = "SPECBOX_SUBSTRATE";
			public static string SPEC_BLOCK_FINISH = "SPECBOX_FINISH";
			public static string SPEC_BLOCK_WOOD_CABINET = "SPECBOX_WOODCABINET";

			public static int TableColItem = 0;
			public static int TableColDimAcross = 1;
			public static int TableColDimWith = 2;
			public static int TableColColor = 3;
			public static int TableColEdgeA = 4;
			public static int TableColEdgeB = 5;
			public static int TableColEdgeC = 6;
			public static int TableColEdgeD = 7;
			public static int TableColNotes = 8;

			public static string EB = "EB";
		}
		public static partial class NonCatalog
		{
			public static int TableColRef = 0;
			public static int TableColItem = 1;
			public static int TableColDescription = 2;
			public static int TableColNotes = 3;
			public static int TableColManuf = 4;
			public static int TableColModel = 5;
			public static int TableColQuantity = 6;
			public static int TableColUOM = 7;
			public static int TableColType = 8;

			public const string TypeCharStock = "n";
			public const string TypeCharNonStock = "o";
			public const string TypeCharDetail = "z";
			public const string TypeCharMetal = "u";
			public const string TypeCharNonItem = "l";

			public static string XDFieldParentItemNo = "ParentItemNo";
			public static string XDFieldParentItemHandle = "ParentItemHandle";
			public static string XDFieldPhase = "Phase";

			public static string GoesWithMessage = "Goes with item(s) ";

			public static string TableTypeFieldName = "TableType";
			public static string TableTypeNonCatalog = "NonCatalog";
			public static string TableNumber = "TableNumber";
		}
		public static partial class Cabinet
		{
			public const string BLOCK_NAME_BASE = "SYM_CAB_B";
			public const string BLOCK_NAME_UPPER = "SYM_CAB_U";
			public const string BLOCK_NAME_MULTI = "SYM_CAB_M";
			public const string DESIGN = "DESNO";
			public const string ITEM_NUMBER = "ITEMNO";
			public const string COLOR = "COLOR";
			public const string SIZE = "SIZE";
			public const string LINE_1 = "LINE-1";
			public const string LINE_2 = "LINE-2";
			public const string LINE_3 = "LINE-3";
			public const string LINE_4 = "LINE-4";
			public const string LINE_5 = "LINE-5";
			public const string FINISHED_ENDS = "FIN";
			public const string MULTIROOM_RANGE = "RANGE";
			public const string RELEASE_STATUS = "RELEASE_STATUS";
			public const string PHASE = "PHASE";
		}
		public static partial class HFItem
		{
			public const string BLOCK_NAME_BASE = "SYM_CAB_B";
			public const string BLOCK_NAME_UPPER = "SYM_CAB_U";
			public const string BLOCK_NAME_MULTI = "SYM_CAB_M";
			public const string DESIGN = "DESNO";
			public const string ITEM_NUMBER = "ITEMNO";
			public const string COLOR = "COLOR";
			public const string SIZE = "SIZE";
			public const string LINE_1 = "LINE-1";
			public const string LINE_2 = "LINE-2";
			public const string LINE_3 = "LINE-3";
			public const string LINE_4 = "LINE-4";
			public const string FINISHED_ENDS = "FIN";
			public const string MULTIROOM_RANGE = "RANGE";
			public const string RELEASE_STATUS = "RELEASE_STATUS";
			public const string PHASE = "PHASE";
			public const string QUANTITY = "QTY";
			public const string UOM = "U.O.M.";
			public const string MANUF = "MANUFACTURER";
			public const string MODEL = "MODEL";
			public const string DESC_1 = "DESCRIPTION1";
			public const string DESC_2 = "DESCRIPTION2";
			public const string SIZE_DOT_FINISH = "SIZE_DOT_FINISH";
			public const string NOTES = "NOTES";
		}
		public static partial class Countertop
		{
			public const string TypeName = "TYPENAME";

		}
		public static partial class ToeKick
		{
			public const string BLOCK_NAME = "SYM_TOE";
			public const string BLOCK_NAME_MULTI = "SYM_TOE_M";
			public const string DESIGN = "DESNO";
			public const string ITEM_NUMBER = "ITEMNO";
			public const string COLOR = "COLOR";
			public const string LINE_1 = "LINE-1";
			public const string LINE_2 = "LINE-2";
			public const string LINE_3 = "LINE-3";
			public const string LINE_4 = "LINE-4";
			public const string MULTIROOM_RANGE = "RANGE";
			public const string RELEASE_STATUS = "RELEASE_STATUS";
			public const string PHASE = "PHASE";
			public const string WIDTH = "WIDTH";
			public const string HEIGHT = "HEIGHT";
			public const string DEPTH = "DEPTH";

			internal static bool IsValidToeDesign(string design)
			{
				bool res = false;
				string des = design.Trim();
				des = des.Replace("-", "");
				des = des.Replace("M", "");
				if (des.Contains("="))
				{
					des = des.Substring(0, des.IndexOf("=")).Trim();
				}
				pgDB odb = new pgDB();
				string sql = "SELECT model_number " +
					"FROM toe_models " +
					"WHERE model_number ILIKE '" + des + "'";
				res = odb.pgValue("eng",sql).Length > 0;
				return res;
			}
		}

		public static partial class CADFilePaths
		{
			public const string BLANK_IMAGE_FILE = @"\\CAD_SVR\CATALOG\CABWEST\SYM\Blank.bmp";
			public const string ICON_DIRECTORY_LOCAL = @"C:\Program Files\Autodesk\AutoCAD 2011\support\Icons\";
			public const string ICON_DIRECTORY_SERVER = @"\\CAD_SVR\CATALOG\CABWEST\SYM\";
			public const string BLANK_BUTTON_16_FILE = "Unknown_16.bmp";
			public const string BLANK_BUTTON_32_FILE = "Unknown_32.bmp";
			public const string SYMBOLS_DIRECTORY = @"\\CAD_SVR\CATALOG\CABWEST\SYM\";


			public const string SETTINGS_BLOCK_PATH = @"\\cad_svr\catalog\CABWEST\W_R2K_Set.dwg";

		}
		public static partial class CADMenu
		{
			public static string RIBBON_PANEL_DRAFTING = "WPI-Drafting";
			public static string RIBBON_PANEL_CABINET = "WPI-Cabinet";
			public static string RIBBON_PANEL_ITEM_TAGS = "WPI-Item Tags";
			public static string RIBBON_PANEL_DIMENSIONS = "WPI-Dimensions";
			public static string RIBBON_PANEL_TEXT = "WPI-Text";
			public static string RIBBON_PANEL_HATCH = "WPI-Hatch";
			public static string RIBBON_PANEL_BLOCKS = "WPI-Blocks";
			public static string RIBBON_PANEL_LAYERS = "WPI-Layers";
			public static string RIBBON_PANEL_TOOLBOX = "WPI-Toolbox";
			public static string RIBBON_PANEL_PANELING = "WPI-Paneling";
			public static string RIBBON_PANEL_PANEL_TABLE_EDIT = "WPI-Panel_Table_Edit";
			public static string RIBBON_PANEL_ITEM_TABLE = "WPI-Item_Table";
			public static string RIBBON_TAB_WESTMARK = "Westmark";
			public static string WORKSPACE_WESTMARK_SIDE = "Westmark-Side";
			public static string WORKSPACE_WESTMARK_SIDE_TOOLBARS = "Westmark-Side-WToolbars";
			public static string WORKSPACE_WESTMARK_TOP = "Westmark-Top";
			public static string WORKSPACE_WESTMARK_TOP_TOOLBARS = "Westmark-Top-WToolbars";
			
			public static string DB_TABLE_ACAD_PANEL_DEF = "acad_panel_def";
			public static string DB_FIELD_ID = "id"; 
			public static string DB_FIELD_PNL_NAME ="panel_name";
			public static string DB_FIELD_PNL_KEY_TIP = "key_tip";
			public static string DB_FIELD_PNL_TITLE = "title";
			public static string DB_FIELD_ELEM_ID = "element_id";
			
			public static string DB_TABLE_ACAD_PANEL_ITEMS = "acad_panel_items";
			public static string DB_FIELD_PANEL_ID = "panel_id";
			public static string DB_FIELD_ROW_INDEX = "row_index";
			public static string DB_FIELD_ITEM_ORDER  = "item_order";
			// public static string DB_FIELD_MACRO_ID  = "macro_id"; // macroid is generated by the cui system
			public static string DB_FIELD_MACRO_NAME = "macro_name";
			public static string DB_FIELD_BUTTON_STYLE  = "button_style";
			public static string DB_FIELD_TOOL_TIP  = "tool_tip";
			public static string DB_FIELD_DISPLAY_NAME  = "display_name";
			public static string DB_FIELD_BUTTON_NAME  = "button_name";
			public static string DB_FIELD_MACRO  = "macro";
			public static string DB_FIELD_SMALL_IMG  = "small_image";
			public static string DB_FIELD_LARGE_IMG  = "large_image";
			public static string DB_FIELD_LABEL  = "label";
			public static string DB_FIELD_SUBROW_INDEX = "sub_row";
			public static string DB_FIELD_SUBPANEL_INDEX = "sub_panel";
			
			public const string BUTTON_STYLE_LARGE_TXT_HOR = "large_w_text_hor";
			public const string BUTTON_STYLE_LARGE_TXT = "large_w_text";
			public const string BUTTON_STYLE_SMALL = "small_wo_text";
			public const string BUTTON_STYLE_SMALL_TXT = "small_w_text";

			public const string SLIDE_OUT = "slideout";
		}
		public static partial class Spec
		{
			public static string SPEC_BLOCK_SOLID_WOOD = "SPECBOX_WOOD";
			public static string SPEC_BLOCK_PANELING = "SPECBOX_PANELING";
			public static string SPEC_BLOCK_SUBSTRATE_PANELING = "SPECBOX_SUBSTRATE";
			public static string SPEC_BLOCK_FINISH = "SPECBOX_FINISH";
			public static string SPEC_BLOCK_WOOD_CABINET = "SPECBOX_WOODCABINET";
			public static string SPEC_BLOCK_HPDL_PANELING = "SPECBOX_HPDL";

			public static string ATTNAME_REFERENCE = "REFERENCE";
			public static string ATTNAME_WOOD_SPECIES = "SPECIES";
			public static string ATTNAME_CUT_PLAINSAWN_X = "PLAIN_SAWN_X";
			public static string ATTNAME_CUT_QUARTERSAWN_X = "QUARTER_SAWN_X";
			public static string ATTNAME_CUT_RIFTSAWN_X = "RIFT_SAWN_X";
			public static string ATTNAME_CUT_OTHER_X = "OTHER_SAWN_X";
			public static string ATTNAME_CUT_OTHER = "OTHER_SAWN";
			public static string ATTNAME_GRADE_CUSTOM_X = "CUSTOM_GRADE_X";
			public static string ATTNAME_GRADE_PREMIUM_X = "PREMIUM_GRADE_X";
			public static string ATTNAME_LEAF_MATCH_BOOK_X = "BOOK_MATCH_X";
			public static string ATTNAME_LEAF_MATCH_SLIP_X = "SLIP_MATCH_X";
			public static string ATTNAME_LEAF_MATCH_OTHER_X = "OTHER_LEAF_MATCH_X";
			public static string ATTNAME_LEAF_MATCH_OTHER = "OTHER_LEAF_MATCH";
			public static string ATTNAME_PANEL_MATCH_RUNNING_X = "RUNNING_MATCH_X";
			public static string ATTNAME_PANEL_MATCH_OTHER_X = "OTHER_PANEL_MATCH_X";
			public static string ATTNAME_PANEL_MATCH_OTHER = "OTHER_PANEL_MATCH";
			// changed to specific core selected from db list plus a notes box
			public static string ATTNAME_CORE = "CORE";

			// OBSOLETE
			public static string ATTNAME_CORE_OTHER_X = "CORE_OTHER_X";
			public static string ATTNAME_CORE_OTHER = "CORE_OTHER";
			public static string ATTNAME_CORE_PB_X = "CORE_PB_X";
			public static string ATTNAME_CORE_MDF_X = "CORE_MDF_X";
			
			public static string ATTNAME_BLUEPRINT_YES_X = "BLUEPRINT_YES_X";
			public static string ATTNAME_BLUEPRINT_NO_X = "BLUEPRINT_NO_X";
			public static string ATTNAME_END_MATCH_YES_X = "END_MATCH_YES_X";
			public static string ATTNAME_END_MATCH_NO_X = "END_MATCH_NO_X";
			public static string ATTNAME_FIRE_RATED_YES_X = "FIRE_RATED_YES_X";
			public static string ATTNAME_FIRE_RATED_NO_X = "FIRE_RATED_NO_X";
			public static string ATTNAME_FSC_YES_X = "FSC_YES_X";
			public static string ATTNAME_FSC_NO_X = "FSC_NO_X";

			public static string ATTNAME_WOODMATCH_COMP_X = "COMPATIBLE_X";
			public static string ATTNAME_WOODMATCH_WELL_X = "WELL_MATCHED_X";

			public static string ATTNAME_PAINTGRADE_YES = "PAINTGRADE_YES_X";
			public static string ATTNAME_PAINTGRADE_NO = "PAINTGRADE_NO_X";
			public static string ATTNAME_FIN_REFERENCE = "FINISH_REFERENCE";
			public static string ATTNAME_FIN_BY_WMK_YES = "FINISH_BY_WM_YES_X";
			public static string ATTNAME_FIN_BY_WMK_NO = "FINISH_BY_WM_NO_X";
			public static string ATTNAME_PRIME_BY_WMK_YES = "PRIME_BY_WM_YES_X";
			public static string ATTNAME_PRIME_BY_WMK_NO = "PRIME_BY_WM_NO_X";
			public static string ATTNAME_BACKPRIME_BY_WMK_YES = "BACKPRIME_BY_WM_YES_X";
			public static string ATTNAME_BACKPRIME_BY_WMK_NO = "BACKPRIME_BY_WM_NO_X";
			public static string ATTNAME_SHEEN_SATIN_X = "SHEEN_SATIN_X";
			public static string ATTNAME_SHEEN_OTHER_X = "SHEEN_OTHER_X";
			public static string ATTNAME_SHEEN_OTHER = "OTHER_SHEEN";
			public static string ATTNAME_STAIN_NONE_X = "STAIN_NONE_X";
			public static string ATTNAME_STAIN_MATCH_X = "STAIN_MATCH_X";
			public static string ATTNAME_STAIN_OTHER_X = "STAIN_OTHER_X";
			public static string ATTNAME_STAIN_OTHER = "OTHER_STAIN";

			public static string ATTNAME_NOTES_WOOD = "WOOD_NOTES";
			public static string ATTNAME_NOTES_FINISH = "FINISH_NOTES";
			public static string ATTNAME_NOTES_VENEER = "VENEER_NOTES";
			public static string ATTNAME_NOTES_SUBSTRATE = "SUBSTRATE_NOTES";
			public static string ATTNAME_NOTES_CABINET = "CABINET_NOTES";

			public static string ATTNAME_EB_MAT_PER_LEGEND_X = "EB_MAT_PER_LEGEND_X";
			public static string ATTNAME_EB_MAT_OTHER_X = "EB_MAT_OTHER_X";
			public static string ATTNAME_EB_MAT_OTHER = "EB_MAT_OTHER";
			public static string ATTNAME_EB_THICKNESS = "EB_THICKNESS";
			public static string ATTNAME_EB_FIN_SAME_AS_FACE_X = "EB_FIN_SAMEASFACE_X";
			public static string ATTNAME_EB_FIN_OTHER_X = "EB_FIN_OTHER_X";
			public static string ATTNAME_EB_FIN_OTHER = "EB_FIN_OTHER";
			
		}
	
	}

	public class SpecPageDBItem
	{
		public string specfamily = string.Empty;
		public string specgroup = string.Empty;
		public string blockname = string.Empty;
		public string desc = string.Empty;
		public int lineCount = 0;
		public string fNumber = string.Empty;
		private string _dbFieldName = string.Empty;
		public string DBFieldName
		{
			get
			{
				string res = specfamily;
				res += "_" + specgroup;
				return res.Replace(" ", "_");
			}
		}
	}
	public class SpecPageDBItems : System.Collections.ObjectModel.Collection<SpecPageDBItem>
	{

		public static SpecPageDBItem GetItemByBlockName(System.Collections.ObjectModel.Collection<SpecPageDBItem> SpecItems, string BlockName)
		{
			SpecPageDBItem resItem = null;
			foreach (SpecPageDBItem spi in SpecItems)
			{
				if (spi.blockname == BlockName)
				{
					resItem = spi;
					break;
				}
			}
			return resItem;
		}
		
		public const string specPageTableName = "cad_spec_page_data";
		public const string specPageUnresolvedTableName = "cad_spec_page_data_unresolved";
		public const string existing_uid = "existing_uid";
		public const string uid = "uid"; // serial NOT NULL,
		public const string door_magnetic_catch = "door_magnetic_catch"; //  text,
		public const string drawer_slide_file = "drawer_slide_file"; //  text,
		public const string drawer_slide_paper = "drawer_slide_paper"; //  text,
		public const string drawer_slide_standard = "drawer_slide_standard"; //  text,
		public const string grommet_color = "grommet_color"; //  text,
		public const string grommet_type = "grommet_type"; //  text,
		public const string hinge_concealed = "hinge_concealed"; //  text,
		public const string hinge_knuckle = "hinge_knuckle"; //  text,
		public const string lock_door = "lock_door"; //  text,
		public const string lock_drawer = "lock_drawer"; //  text,
		public const string lock_latch = "lock_latch"; //  text,
		public const string material_closed = "material_closed"; //  text,
		public const string material_open = "material_open"; //  text,
		public const string material_under_upper = "material_under_upper"; //  text,
		public const string metal_color = "metal_color"; //  text,
		public const string pull_recessed = "pull_recessed"; //  text,
		public const string pull_wire = "pull_wire"; //  text,
		public const string shelf_size = "shelf_size"; //  text,
		public const string shelf_rest = "shelf_rest"; //  text,
		public const string substrate_cabinet = "substrate_cabinet"; //  text,
		public const string substrate_drawer = "substrate_drawer"; //  text,
		public const string toe_material = "toe_material"; //  text,
		public const string wood_finish = "wood_finish"; //  text,
		public const string wood_type = "wood_type"; //  text,
		public const string job_number = "job_number"; //  integer,
		public const string fulldwgpath = "fulldwgpath"; //  text,
		public const string cab_series = "cab_series"; // text,

		public bool ContainsBlock(string BlockName)
		{
			bool res = false;
			foreach (SpecPageDBItem spi in this)
			{
				if (spi.blockname == BlockName)
				{
					res = true;
					break;
				}
			}
			return res;
		}

		internal SpecPageDBItem GetByBlock(string BlockName)
		{
			SpecPageDBItem sdi = null;
			foreach (SpecPageDBItem spi in this)
			{
				if (spi.blockname == BlockName)
				{
					sdi = spi;
					break;
				}
			}
			return sdi;
		}
	}
}

using System;
using System.Collections;
using System.Data;
using System.Net.NetworkInformation;
using System.Text;
using Npgsql;

namespace pgDBTools
{
    /// <summary>
    /// Summary description for pgDB.
    /// </summary>
    public class pgDB
    {
        private string eng_sys_db = "SERVER=webapps;" +
            "DATABASE=engineering_system;" +
            "USER ID=postgres;" +
            "PASSWORD=;";

        private string spec_db = "SERVER=webapps;" +
            "DATABASE=project_specs;" +
            "USER ID=postgres;" +
            "PASSWORD=;";

        private string mto_db = "SERVER=webapps;" +
            "DATABASE=MTO;" +
            "USER ID=postgres;" +
            "PASSWORD=;";

        //private string eng_sys_db_local = "SERVER=localhost;" +
        //    "DATABASE=engineering_system;" +
        //    "USER ID=postgres;" +
        //    "PASSWORD=engf7255;";
        //private string spec_db_local = "SERVER=localhost;" +
        //    "DATABASE=project_specs;" +
        //    "USER ID=postgres;" +
        //    "PASSWORD=engf7255;";
        //private string mto_db_local = "SERVER=localhost;" +
        //    "DATABASE=mto;" +
        //    "USER ID=postgres;" +
        //    "PASSWORD=engf7255;";

        public enum dbServer
        {
            EngineeringSystemServer,
            ProjectSpecsServer,
            MTOServer,
            Other
        }

        public pgDB()
        {
            if (!PingIsResponding("webapps"))
            {
                //if (PingIsResponding("localhost"))
                //{
                //    eng_sys_db = eng_sys_db_local;
                //    spec_db = spec_db_local;
                //}
            }
        }

        public static bool PingIsResponding(string HostOrAddress)
        {
            bool res = false;
            Ping pingSender = new Ping();
            PingOptions options = new PingOptions();

            // Use the default Ttl value which is 128,
            // but change the fragmentation behavior.
            options.DontFragment = true;

            // Create a buffer of 32 bytes of data to be transmitted.
            string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            byte[] buffer = Encoding.ASCII.GetBytes(data);

            // wait 1/2 second (500) for reply
            // changed to 2 second
            int timeout = 2000;

            try
            {
                PingReply reply = pingSender.Send(HostOrAddress, timeout, buffer, options);
                if (reply.Status == IPStatus.Success)
                {
                    res = true;
                }
            }
            catch (Exception ex)
            {
                // throw;
            } return res;
        }

        public string SLZ(string FullString)
        {
            string res = FullString;
            bool go = true;
            while (go)
            {
                if (res.StartsWith("0") && res.Length > 1)
                {
                    res = res.Substring(1);
                }
                else
                {
                    go = false;
                }
            }
            if (res == "0")
            {
                res = "";
            }
            return res;
        }

        public string PrepForSQL(string unpreped)
        {
            string temp = unpreped ?? string.Empty;
            temp = temp.Trim();
            if (temp.Length == 0) return "";
            temp = temp.Replace("\\", "\\\\");
            temp = temp.Replace("\'", "\\\'");
            temp = temp.Replace("'", "\'");

            // temp = temp.Replace("&", "\&");
            return temp;
        }

        public ArrayList pgList(string conn, string ssql)
        {
            // RETURNS A ONE DIMENSIONAL LIST OF RESULTS
            // If the result of the query is not exactly one value then returns empty arraylist
            ArrayList aList = new ArrayList();

            // create a connection
            Npgsql.NpgsqlConnection db;
            switch (conn)
            {
                case "eng":
                case "engsys":
                case "eng_sys":
                    db = new Npgsql.NpgsqlConnection(eng_sys_db);
                    break;

                case "spec":
                    db = new Npgsql.NpgsqlConnection(spec_db);
                    break;

                case "mto":
                    db = new Npgsql.NpgsqlConnection(mto_db);
                    break;

                default:
                    db = new Npgsql.NpgsqlConnection(conn);
                    break;
            }
            db.Open();

            // create an empty dataset
            System.Data.DataSet ds = new System.Data.DataSet();

            // create a pgDataAdapter
            Npgsql.NpgsqlDataAdapter da = new Npgsql.NpgsqlDataAdapter();
            da.SelectCommand = new NpgsqlCommand(ssql, db);

            // fill the dataset using the data adapter
            da.Fill(ds);
            da.Dispose();
            db.Close();
            db.Dispose();
            if (ds.Tables.Count == 0)
                return aList;
            if (ds.Tables[0].Columns.Count != 1)
                return aList;
            if (ds.Tables[0].Rows.Count < 1)
                return aList;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                aList.Add(ds.Tables[0].Rows[i][0].ToString());
            }
            return aList;
        }

        public string pgValue(string conn, string ssql)
        {
            string res = "";
            res = pgValue(conn, ssql, false);
            return res;
        }

        public string pgValue(dbServer ServerName, string ssql)
        {
            string res = string.Empty;
            res = pgValue(ServerName, ssql, false);
            return res;
        }

        public string pgValue(dbServer ServerName, string ssql, bool firstOnly)
        {
            string val = string.Empty;

            switch (ServerName)
            {
                case dbServer.EngineeringSystemServer:
                    val = pgValue(eng_sys_db, ssql, firstOnly);
                    break;

                case dbServer.ProjectSpecsServer:
                    val = pgValue(spec_db, ssql, firstOnly);
                    break;

                case dbServer.MTOServer:
                    val = pgValue(mto_db, ssql, firstOnly);
                    break;
            }
            return val;
        }

        public string pgValue(string conn, string ssql, bool firstOnly)
        {
            // RETURNS ONE STRING VALUE
            // If the result of the query is not exactly one value then returns empty string
            // unless firstOnly is true. Then returns first value.
            // create a connection
            Npgsql.NpgsqlConnection db;

            // create an empty dataset
            System.Data.DataSet ds = new System.Data.DataSet();

            // create a pgDataAdapter
            Npgsql.NpgsqlDataAdapter da = new Npgsql.NpgsqlDataAdapter();

            string val = "";

            switch (conn)
            {
                case "eng":
                case "engsys":
                case "eng_sys":
                    db = new Npgsql.NpgsqlConnection(eng_sys_db);
                    break;

                case "spec":
                    db = new Npgsql.NpgsqlConnection(spec_db);
                    break;

                case "mto":
                    db = new Npgsql.NpgsqlConnection(mto_db);
                    break;

                default:
                    db = new Npgsql.NpgsqlConnection(conn);
                    break;
            }

            try
            {
                db.Open();
                da.SelectCommand = new NpgsqlCommand(ssql, db);
                try
                {
                    // fill the dataset using the data adapter
                    da.Fill(ds);
                    if (ds.Tables.Count > 0)
                    {
                        if (ds.Tables[0].Columns.Count == 1)
                        {
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                if (firstOnly)
                                {
                                    val = ds.Tables[0].Rows[0][0].ToString();
                                }
                                else
                                {
                                    if (ds.Tables[0].Rows.Count != 1)
                                    {
                                        val = "";
                                    }
                                    else
                                    {
                                        val = ds.Tables[0].Rows[0][0].ToString();
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception exds)
                {
                    System.Windows.Forms.MessageBox.Show("Error in pgDB.pgValue\n" +
                        exds.Message + "\n" +
                        "Connection string: " + db.ConnectionString + "\n" +
                        "SQL string: " + ssql, "pgDB Error");
                }
                finally
                {
                    db.Close();
                }
            }
            catch (Exception exdb)
            {
                System.Windows.Forms.MessageBox.Show("Error in pgDB.pgValue\n" +
                    exdb.Message + "\n" +
                    "Connection string: " + db.ConnectionString, "pgDB Error");
            }
            finally
            {
                da.Dispose();
                ds.Dispose();
                db.Dispose();
            }
            return val;
        }

        public int pgSQL(string conn, string ssql)
        {
            int effCount = 0;

            // create a connection
            Npgsql.NpgsqlConnection db;

            NpgsqlCommand sqlCMD = new NpgsqlCommand();

            switch (conn)
            {
                case "eng":
                case "engsys":
                case "eng_sys":
                    db = new Npgsql.NpgsqlConnection(eng_sys_db);
                    break;

                case "spec":
                    db = new Npgsql.NpgsqlConnection(spec_db);
                    break;

                case "mto":
                    db = new Npgsql.NpgsqlConnection(mto_db);
                    break;

                default:
                    db = new Npgsql.NpgsqlConnection(conn);
                    break;
            }
            try
            {
                db.Open();
                try
                {
                    sqlCMD.Connection = db;
                    sqlCMD.CommandText = ssql;
                    effCount = sqlCMD.ExecuteNonQuery();
                }
                catch (Exception exds)
                {
                    System.Windows.Forms.MessageBox.Show("Error in pgDB.pgSQL\n" +
                        exds.Message + "\n" +
                        "Connection string: " + db.ConnectionString + "\n" +
                        "SQL string: " + ssql, "pgDB Error");
                }
                finally
                {
                    db.Close();
                }
            }
            catch (Exception exdb)
            {
                System.Windows.Forms.MessageBox.Show("Error in pgDB.pgSQL\n" +
                    exdb.Message + "\n" +
                    "Connection string: " + db.ConnectionString, "pgDB Error");
            }
            finally
            {
                sqlCMD.Dispose();
                db.Dispose();
            }
            return effCount;
        }

        public DataSet pgDataSet(string conn, string ssql)
        {
            // create a connection
            Npgsql.NpgsqlConnection db;

            // create an empty dataset
            System.Data.DataSet ds = new System.Data.DataSet();

            // create a pgDataAdapter
            NpgsqlDataAdapter da = new NpgsqlDataAdapter();
            switch (conn)
            {
                case "eng":
                case "engsys":
                case "eng_sys":
                    db = new Npgsql.NpgsqlConnection(eng_sys_db);
                    break;

                case "spec":
                    db = new Npgsql.NpgsqlConnection(spec_db);
                    break;

                case "mto":
                    db = new Npgsql.NpgsqlConnection(mto_db);
                    break;

                default:
                    db = new Npgsql.NpgsqlConnection(conn);
                    break;
            }
            try
            {
                db.Open();
                da.SelectCommand = new NpgsqlCommand(ssql, db);
                try
                {
                    // fill the dataset using the data adapter
                    da.Fill(ds);

                    //da.Dispose();
                    //db.Close();
                    //db.Dispose();
                }
                catch (Exception exds)
                {
                    System.Windows.Forms.MessageBox.Show("Error in pgDB.pgDataSet filling Dataset\n" +
                        exds.Message + "\n" +
                        "Connection string: " + db.ConnectionString.ToString() + "\n" +
                        "SQL string: " + ssql, "pgDB Error");
                }
                finally
                {
                    db.Close();
                }
            }
            catch (Exception exdb)
            {
                System.Windows.Forms.MessageBox.Show("Error in pgDB.pgDataSet opening Database\n" +
                    exdb.Message + "\n" +
                    "Connection string: " + db.ConnectionString.ToString(), "pgDB Error");
            }
            finally
            {
                da.Dispose();
                db.Dispose();
            }
            return ds;
        }

        public DataSet GetTableNames(string conn)
        {
            // create a connection
            Npgsql.NpgsqlConnection db;
            switch (conn)
            {
                case "eng":
                case "engsys":
                case "eng_sys":
                    db = new Npgsql.NpgsqlConnection(eng_sys_db);
                    break;

                case "spec":
                    db = new Npgsql.NpgsqlConnection(spec_db);
                    break;

                case "mto":
                    db = new Npgsql.NpgsqlConnection(mto_db);
                    break;

                default:
                    db = new Npgsql.NpgsqlConnection(conn);
                    break;
            }
            db.Open();
            string ssql = "SELECT relname FROM pg_class WHERE relkind = 'r' AND relname NOT LIKE 'pg_%' ORDER BY relname";

            // create an empty dataset
            System.Data.DataSet ds = new System.Data.DataSet();

            // create a pgDataAdapter
            Npgsql.NpgsqlDataAdapter da = new Npgsql.NpgsqlDataAdapter();
            da.SelectCommand = new NpgsqlCommand(ssql, db);

            // fill the dataset using the data adapter
            da.Fill(ds);
            da.Dispose();
            db.Close();
            db.Dispose();
            return ds;
        }

        public DataSet GetTableFields(string conn, string tableName)
        {
            // create a connection
            Npgsql.NpgsqlConnection db;
            switch (conn)
            {
                case "eng":
                case "engsys":
                case "eng_sys":
                    db = new Npgsql.NpgsqlConnection(eng_sys_db);
                    break;

                case "spec":
                    db = new Npgsql.NpgsqlConnection(spec_db);
                    break;

                case "mto":
                    db = new Npgsql.NpgsqlConnection(mto_db);
                    break;

                default:
                    db = new Npgsql.NpgsqlConnection(conn);
                    break;
            }
            db.Open();
            string ssql = "select attname from pg_attribute where attrelid = " +
                "(select oid from pg_class where relname = '" + tableName +
                "') AND attstattarget != 0 ORDER BY attname";

            // create an empty dataset
            System.Data.DataSet ds = new System.Data.DataSet();

            // create a pgDataAdapter
            Npgsql.NpgsqlDataAdapter da = new Npgsql.NpgsqlDataAdapter();
            da.SelectCommand = new NpgsqlCommand(ssql, db);

            // fill the dataset using the data adapter
            da.Fill(ds);
            da.Dispose();
            db.Close();
            db.Dispose();
            return ds;
        }

        public string GetUniqueValues(DataTable dt, string FieldName, string Seperator)
        {
            string res = "";
            SortedList slUniqueItems = new SortedList();
            string thisName = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                thisName = dt.Rows[i][FieldName].ToString();
                if (!slUniqueItems.Contains(thisName))
                {
                    slUniqueItems.Add(thisName, thisName);
                }
            }
            for (int i = 0; i < slUniqueItems.Count; i++)
            {
                thisName = slUniqueItems[slUniqueItems.GetKey(i)].ToString();
                if (res.Length == 0)
                {
                    res = thisName;
                }
                else
                {
                    res = res + Seperator + thisName;
                }
            }
            return res;
        }

        public DataTable pgDataTable(string conn, string ssql)
        {
            // create a connection
            Npgsql.NpgsqlConnection db;

            // create an empty dataset
            System.Data.DataSet ds = new System.Data.DataSet();
            System.Data.DataTable dt = new System.Data.DataTable();

            // create a pgDataAdapter
            Npgsql.NpgsqlDataAdapter da = new Npgsql.NpgsqlDataAdapter();
            switch (conn)
            {
                case "eng":
                case "engsys":
                case "eng_sys":
                    db = new Npgsql.NpgsqlConnection(eng_sys_db);
                    break;

                case "spec":
                    db = new Npgsql.NpgsqlConnection(spec_db);
                    break;

                case "mto":
                    db = new Npgsql.NpgsqlConnection(mto_db);
                    break;

                default:
                    db = new Npgsql.NpgsqlConnection(conn);
                    break;
            }
            try
            {
                db.Open();
                da.SelectCommand = new NpgsqlCommand(ssql, db);
                try
                {
                    // fill the dataset using the data adapter
                    da.Fill(ds);

                    //da.Dispose();
                    //db.Close();
                    //db.Dispose();
                }
                catch (Exception exds)
                {
                    System.Windows.Forms.MessageBox.Show("Error in pgDB.pgDataSet filling Dataset\n" +
                        exds.Message + "\n" +
                        "Connection string: " + db.ConnectionString.ToString() + "\n" +
                        "SQL string: " + ssql, "pgDB Error");
                }
                finally
                {
                    db.Close();
                }
            }
            catch (Exception exdb)
            {
                System.Windows.Forms.MessageBox.Show("Error in pgDB.pgDataSet opening Database\n" +
                    exdb.Message + "\n" +
                    "Connection string: " + db.ConnectionString.ToString(), "pgDB Error");
            }
            finally
            {
                da.Dispose();
                db.Dispose();
            }
            if (ds.Tables.Count > 0) dt = ds.Tables[0];
            return dt;
        }

        public string ParseHashtableForInsert(Hashtable ht)
        {
            if (ht.Count < 1)
                return string.Empty;
            string res = string.Empty;

            try
            {
                res = "(";
                string flds = string.Empty;
                string vals = string.Empty;
                foreach (DictionaryEntry de in ht)
                {
                    flds += de.Key.ToString() + ", ";

                    // quote strings
                    vals += de.Value.ToString() + ", ";
                }
                flds = flds.Substring(0, flds.Length - 2);
                vals = vals.Substring(0, vals.Length - 2);
                res = res + flds + ") VALUES (" + vals + ")";
            }
            catch (Exception ex)
            {
                throw new Exception("Error in pgDBTools.ParseHashtableForInsert", ex.InnerException);
            }
            return res;
        }

        public string ParseHashtableForInsert(Hashtable ht, bool AutoQuote)
        {
            if (!AutoQuote)
            {
                return ParseHashtableForInsert(ht);
            }
            if (ht.Count < 1)
                return string.Empty;
            string res = string.Empty;

            try
            {
                res = "(";
                string flds = string.Empty;
                string vals = string.Empty;
                foreach (DictionaryEntry de in ht)
                {
                    flds += de.Key.ToString() + ", ";

                    // quote strings
                    if (de.Value.GetType() == typeof(string) &&
                        de.Value.ToString() != Util.Now)
                    {
                        vals += QuoteIt(de.Value.ToString()) + ", ";
                    }
                    else
                    {
                        vals += de.Value.ToString() + ", ";
                    }
                }
                flds = flds.Substring(0, flds.Length - 2);
                vals = vals.Substring(0, vals.Length - 2);
                res = res + flds + ") VALUES (" + vals + ")";
            }
            catch (Exception ex)
            {
                throw new Exception("Error in pgDBTools.ParseHashtableForInsert", ex.InnerException);
            }
            return res;
        }

        public string ParseHashtableForUpdate(Hashtable ht)
        {
            if (ht.Count < 1)
                return string.Empty;
            string res = "";

            try
            {
                foreach (DictionaryEntry de in ht)
                {
                    // boolean should be text
                    res = res + de.Key.ToString() + " = " + de.Value.ToString() + ", ";
                }
                res = res.Substring(0, res.Length - 2);
            }
            catch (Exception ex)
            {
                throw new Exception("Error in pgDBTools.ParseHashtableForUpdate", ex.InnerException);
            }
            return res;
        }

        public string ParseHashtableForUpdate(Hashtable ht, bool AutoQuote)
        {
            if (!AutoQuote)
            {
                return ParseHashtableForUpdate(ht);
            }
            if (ht.Count < 1)
                return string.Empty;
            string res = "";

            try
            {
                foreach (DictionaryEntry de in ht)
                {
                    // boolean should be text
                    res = res + de.Key.ToString() + " = " + GetAutoQuoted(de.Value) + ", ";
                }
                res = res.Substring(0, res.Length - 2);
            }
            catch (Exception ex)
            {
                throw new Exception("Error in pgDBTools.ParseHashtableForUpdate", ex.InnerException);
            }
            return res;
        }

        private string GetAutoQuoted(object p)
        {
            string res = string.Empty;
            Type t = p.GetType();
            if (t == typeof(string))
            {
                res = QuoteIt(p.ToString());
            }
            else if (t == typeof(bool))
            {
                bool b = (bool)p;
                if (b)
                {
                    res = "true";
                }
                else
                {
                    res = "false";
                }
            }
            else
            {
                res = p.ToString();
            }
            return res;
        }

        public string QuoteIt(string inputString)
        {
            string res = inputString.Trim();
            if (res.Trim().Length == 0)
                return "\'\'";
            if (res.Equals(Util.Now))
                return res;
            
            if (!res.StartsWith("\'"))
            {
                res = "\'" + res;
            }
            if (res.EndsWith("\\\'"))
            {
                res += "\'";
            }
            if (!res.EndsWith("\'"))
            {
                res = res + "\'";
            }
            return res;
        }

        public partial class Util
        {
            public Util()
            { }

            public static bool GetBoolVal(object cellVal, bool DefaultValue)
            {
                bool res = DefaultValue;
                if (cellVal != null)
                {
                    string v = cellVal.ToString().ToUpper();
                    if (v.Trim().Length > 0)
                    {
                        if (v.StartsWith("T"))
                        {
                            res = true;
                        }
                        else
                        {
                            res = false;
                        }
                    }
                }
                return res;
            }
            public static int GetIntVal(object cellVal, int DefaultValue)
            {
                int res = DefaultValue;
                if (cellVal != null)
                {
                    int i;
                    if (int.TryParse(cellVal.ToString(), out i))
                    {
                        res = i;
                    }
                }
                return res;
            }
            public static decimal GetDecVal(object cellVal, decimal DefaultValue)
            {
                decimal res = DefaultValue;
                if (cellVal != null)
                {
                    decimal d;
                    if (decimal.TryParse(cellVal.ToString(), out d))
                    {
                        res = d;
                    }
                }
                return res;
            }
            public static string GetStringVal(object cellVal, string DefaultValue)
            {
                string res = DefaultValue;
                if (cellVal != null)
                {
                    
                    res = PrepForSQL(cellVal.ToString());
                }
                return res;
            }

            internal static string JustDesignNumber(string Design)
            {
                string des = Design;
                pgDB odb = new pgDB();
                string sql = "SELECT just_model_number('" + Design + "')";
                string val = odb.pgValue("eng", sql);
                if (val.Length > 0) des = val;
                return des;
            }

            public static string Now = "now()";
            
            public static string PrepForSQL(string unpreped)
            {
                string temp = unpreped ?? string.Empty;
                temp = temp.Trim();
                if (temp.Length == 0) return "";
                temp = temp.Replace("\\", "\\\\");
                temp = temp.Replace("\'", "\\\'");
                temp = temp.Replace("'", "\'");

                // temp = temp.Replace("&", "\&");
                return temp;
            }
        }
    }
}
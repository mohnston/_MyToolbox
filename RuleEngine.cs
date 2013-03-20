using System;
using System.Data;
using System.Text;
using System.Collections;
using pgDBTools;

namespace WPIRulesEngine
{
	/// <summary>
	/// Summary description for RuleEngine.
	/// </summary>
	public class RuleEngine
	{
		public RuleEngine()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		private string my_where_clause = "";
		private bool my_verbose = false;
        private string crlf = Environment.NewLine;
        private string severityCritical = "Critical";
        private string severityWarning = "Warning";
        private string severityNote = "Note";
        public enum eSeverity
        {
            Critical, Warning, Note, All
        };


        public string GetRuleViolations(string category, string subcategory)
        {
            return GetRuleViolations(category, subcategory, eSeverity.All);
        }

		public string GetRuleViolations( string category, string subcategory, eSeverity severity) 
		{
			// if the where clause is missing just send that message
			if ( WhereClause.Length == 0 ) return "The required WHERE CLAUSE is missing.";
            
			// build the sql statement from the database
			StringBuilder allMessages = new StringBuilder();
			string sql = "";
			DataSet dsRules;
			DataSet dsElems;
			DataSet dsElem;
			string ruleid;
			string sevClause = "";
			bool result;
			string stable;
			string sfield;
			string svalue;
			string soperator;
			string squote;
			string q;
            
			pgDB odb = new pgDB();

            switch (severity)
            {
                case eSeverity.All:
                    break;
                case eSeverity.Critical:
                    sevClause = " AND rule_severity = '" + severityCritical + "' ";
                    break;
                case eSeverity.Note:
                    sevClause = " AND rule_severity = '" + severityNote + "' ";
                    break;
                case eSeverity.Warning:
                    sevClause = " AND rule_severity = '" + severityWarning + "' ";
                    break;
                default:
                    break;
            }

            sql = "SELECT * FROM rules " +
                    "WHERE rule_category = '" + category +
                    "' AND rule_subcategory = '" + subcategory + "'" + sevClause;

			dsRules = odb.pgDataSet("spec", sql);
			if ( this.Verbose ) 
			{
				allMessages.Append(crlf + "Rules SQL = " + sql);
				allMessages.Append(crlf + "Returned " + dsRules.Tables[0].Rows.Count.ToString() + " rows.");
			}
			
			DataTable rsrules;
			DataTable rselems;
			DataTable rselem;

			rsrules = dsRules.Tables[0];
			if ( rsrules.Rows.Count == 0 ) return "No rules have been created for category=" +
																					category + " subcategory=" + subcategory;
			
			for ( int i = 0; i < rsrules.Rows.Count; i++ ) 
			{
				ruleid = rsrules.Rows[i]["rule_id"].ToString();
				sql = "SELECT * FROM rule_elements WHERE rule_id = " + ruleid + " " +
						"ORDER BY elem_order";
				dsElems = odb.pgDataSet("spec", sql);
				rselems = dsElems.Tables[0];
				if ( this.Verbose ) 
				{
					allMessages.Append(crlf + "Rule " + i.ToString() + " SQL = " + sql);
					allMessages.Append(crlf + "Returned " + rselems.Rows.Count.ToString() + " rows.");
				}
				result = true;
				for ( int j = 0; j < rselems.Rows.Count; j++ ) 
				{
					// check out the first rule
					stable = rselems.Rows[j]["elem_table"].ToString();
					sfield = rselems.Rows[j]["elem_key"].ToString();
					svalue = rselems.Rows[j]["elem_value"].ToString();
					soperator = rselems.Rows[j]["elem_operator"].ToString();
					squote = rselems.Rows[j]["quote"].ToString();
					if ( squote.ToUpper() == "TRUE" ) q = @"'";
					else q = "";

					if ( svalue.IndexOf("*") == 0 ) 
					{
						svalue = svalue.Substring(1);
						sql = "SELECT " + sfield + " " +
							"FROM " + stable + " " +
							"WHERE " + WhereClause + 
							" AND " + sfield + " " + soperator + " " + q + svalue + q;
					}
					else
					{
						sql = "SELECT " + sfield + " " +
							"FROM " + stable + " " +
							"WHERE " + WhereClause + 
							" AND " + sfield + " " + soperator + " " + q + svalue + q;
					}
					dsElem = odb.pgDataSet("spec", sql);
                    if (dsElem.Tables.Count > 0)
                    {
                        rselem = dsElem.Tables[0];
                        if (this.Verbose)
                        {
                            allMessages.Append(crlf + "Rule " + i.ToString() + " Element " + j.ToString() + " SQL = " + sql);
                            allMessages.Append(crlf + "Returned " + rselem.Rows.Count.ToString() + " rows.");
                        }

                        if (rselem.Rows.Count == 0)
                        {
                            result = false;
                            j = rselems.Rows.Count;
                        }
                    }
				}// end of the rule element loop
				if ( result == true ) 
				{
					allMessages.Append(crlf + rsrules.Rows[i]["rule_severity"].ToString().ToUpper() + ": " + rsrules.Rows[i]["rule_message"].ToString());
					// i = rsrules.Rows.Count;
				}

			} // end of rules loop
			


			return allMessages.ToString();
		}

        public ArrayList GetCategoryList()
        {
            pgDB odb = new pgDB();
            ArrayList al = new ArrayList();
            string sql = "SELECT distinct rule_category " +
                "FROM rules " +
                "ORDER BY rule_category";
            DataSet ds = odb.pgDataSet("spec", sql);
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        al.Add(ds.Tables[0].Rows[i]["rule_category"].ToString());
                    }
                }
            }
            return al;
        }

        public ArrayList GetSubCategoryList(string CategoryName)
        {
            pgDB odb = new pgDB();
            ArrayList al = new ArrayList();
            string sql = "SELECT distinct rule_subcategory " +
                "FROM rules " +
                "WHERE rule_category = '" + CategoryName + "' " +
                "ORDER BY rule_subcategory";
            DataSet ds = odb.pgDataSet("spec", sql);
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        al.Add(ds.Tables[0].Rows[i]["rule_subcategory"].ToString());
                    }
                }
            }
            return al;
        }

		public void CreateOrUpdateRule( 
			string rule_name, 
			string rule_description,
			string rule_category,
			string rule_subcategory,
			string severity,
			string message) 
		{
			// adds rule if it doesn't already exist
			// updates data if it does exist
			pgDB odb = new pgDB();
			string sql = "SELECT * FROM rules WHERE rule_name = '" + rule_name + "'";
			DataSet rs = odb.pgDataSet("spec", sql);
			if ( rs.Tables[0].Rows.Count == 0 ) 
			{
				sql = "INSERT INTO rules ( rule_name, rule_description, rule_category, " +
					"rule_subcategory, rule_severity, rule_message ) " +
					"VALUES " +
					"('" + rule_name + "', '" + rule_description + "', '" + rule_category + "', '" +
					rule_subcategory + "', '" + severity + "', '" + message + "')";
			}
			else 
			{
				sql = "UPDATE rules SET " +
					"rule_description = '" + rule_description + "', " +
					"rule_category = '" + rule_category + "', " +
					"rule_subcategory = '" + rule_subcategory + "', " +
					"rule_severity = '" + severity + "', " +
					"rule_message = '" + message + "' " +
					"WHERE rule_name = '" + rule_name + "'";
			}
			rs = odb.pgDataSet("spec",sql);
		}
		
		public void AddRuleElement(
			string rule_id,
			string table_name,
			string key_name,
			string value_name,
			string order,
			string operator_name) 
		{
			// adds element 
			pgDB odb = new pgDB();
			string sql = "INSERT INTO rule_elements (rule_id, elem_table, elem_key, elem_value, elem_operator, elem_order) " +
							"VALUES " +
							"(" + rule_id + ", '" + table_name + "', '" + key_name + "', '" + value_name +
							"', '" + operator_name + "', " + order + ")";
			DataSet rs = odb.pgDataSet("spec", sql);
		}

		public void UpdateRuleElement(
			string elem_id,
			string table_name,
			string key_name,
			string value_name,
			string order,
			string operator_name) 
		{
			pgDB odb = new pgDB();
			string sql = "UPDATE rule_elements SET " +
				"elem_table = '" + table_name + "', " +
				"elem_key = '" + key_name + "', " +
				"elem_value = '" + value_name + "', " +
				"elem_order = '" + order + "', " +
				"elem_operator = '" + operator_name + "' " +
				"WHERE rule_element_id = " + elem_id;
			DataSet rs = odb.pgDataSet("spec", sql);
		}	
			
	
		public string WhereClause
		{
			get
			{
				return my_where_clause;
			}
			set
			{
				my_where_clause = value;
			}
		}

		public bool Verbose 
		{
			get 
			{
				return my_verbose;
			}
			set 
			{
				my_verbose = value;
			}
		}
	}
}

using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using Teigha.Runtime;
using Teigha.DatabaseServices;
using Teigha.Colors;
using Teigha.Geometry;

namespace Teigha.Tools
{
    public class TeighaTools
    {
        #region Static names
        public static Teigha.Colors.Color clrMagenta = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, COLOR_WHITE); // acColors.Color.FromColor(System.Drawing.Color.Magenta);
        public static Teigha.Colors.Color clrGreen = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, COLOR_GREEN);
        public static Teigha.Colors.Color clrCyan = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, COLOR_CYAN);
        public static Teigha.Colors.Color clrYellow = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, COLOR_YELLOW);
        public static Teigha.Colors.Color clrWhite = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, COLOR_WHITE);
        public static Teigha.Colors.Color clrBlue = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, COLOR_BLUE);
        public static Teigha.Colors.Color clrRed = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, COLOR_RED);

        private static short COLOR_MAGENTA = 6;
        private static short COLOR_GREEN = 3;
        private static short COLOR_CYAN = 4;
        private static short COLOR_YELLOW = 2;
        private static short COLOR_WHITE = 7;
        private static short COLOR_BLUE = 5;
        private static short COLOR_RED = 1;

        public static string layerNameCabinet_Old = "CABINET";
        public static string layerNameBorder_Old = "BORDER";
        public static string layerNameHidden_Old = "HIDDEN";
        public static string layerNameGreen_Old = "GREEN";
        public static string layerNameText_Old = "TEXT";
        public static string layerNameCyan_Old = "CYAN";
        public static string layerNameLowSwing_Old = "LOWSWING";
        public static string layerNameTallSwing_Old = "TALLSWING";

        public static string layerName_2D_Cabinet = "WPI_Cabinet";
        public static string layerName_2D_Border = "WPI_Border";
        public static string layerName_2D_Hidden = "WPI_Hidden";
        public static string layerName_2D_Green = "WPI_Green";
        public static string layerName_2D_Text = "WPI_Text";
        public static string layerName_2D_Cyan = "WPI_Cyan";

        public static string layerName_3D_Cabinet = "WPI-3D-Cabinet";
        public static string layerName_3D_Cabinet_Base = "WPI-3D-Cabinet-Base";
        public static string layerName_3D_Cabinet_Upper = "WPI-3D-Cabinet-Upper";
        public static string layerName_3D_Cabinet_Panel = "WPI-3D-Cabinet-Panel";

        public static string layerName_3D_Countertop = "WPI-3D-Top";
        public static string layerName_3D_Detail_TransactionTop = "WPI-3D-Top-Transaction";
        public static string layerName_3D_Detail_WorkTop = "WPI-3D-Top-Work";
        public static string layerName_3D_Detail_Top_Lower = "WPI-3D-Top-Lower";
        public static string layerName_3D_Detail_Top_Upper = "WPI-3D-Top-Upper";

        public static string layerName_3D_Toekick = "WPI-3D-Toekick";

        public static string layerName_3D_Wall = "WPI-3D-Wall";
        public static string layerName_3D_Floor = "WPI-3D-Floor";

        public static string layerName_3D_Panel = "WPI-3D-Panel";
        public static string layerName_3D_Detail_Panel = "WPI-3D-Panel-Detail";
        public static string layerName_3D_Panel_Removable = "WPI-3D-Panel-Removable";

        public static string layerName_3D_Framework = "WPI-3D-Framework";
        public static string layerName_3D_Framework_Studs = "WPI-3D-Framework-Studs";
        public static string layerName_3D_Framework_Plate = "WPI-3D-Framework-Plate";

        public static string layerName_3D_Face = "WPI-3D-Face";
        public static string layerName_3D_Face_Panel = "WPI-3D-Face-Panel";
        public static string layerName_3D_Face_Reveal = "WPI-3D-Face-Reveal";

        public static string layerName_3D_Hardware = "WPI-3D-Hardware";
        public static string layerName_3D_Hardware_Bracket = "WPI-3D-Hardware-Bracket";
        public static string layerName_3D_Hardware_Standoff = "WPI-3D-Hardware-Standoff";
        public static string layerName_3D_Hardware_Fastener = "WPI-3D-Hardware-Fastener";

        const string LINETYPE_NAME_CONTINUOUS = "CONTINUOUS";
        const string LINETYPE_NAME_HIDDEN = "HIDDEN";
        const string LINETYPE_NAME_LOWSWING = "DASHDOT2";
        const string LINETYPE_NAME_TALLSWING = "DASHDOT"; 
        #endregion

        public ObjectId TextStyleBoldID = ObjectId.Null;
        int PitchNFamilyTreb = 34;
        int PitchNFamilyOptima = 2;
        private const string wmStd = "WM_Standard";

        private const string alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        
    }
   
    public class BlockTools
    {
        internal static string GetAttributeValue(Database db, Transaction trans, BlockReference bref, string AttTagName)
        {
            string res = string.Empty;
            AttTagName = AttTagName.ToUpper();
            try
            {
                foreach (ObjectId attid in bref.AttributeCollection)
                {
                    AttributeReference attref = (AttributeReference)trans.GetObject(attid, OpenMode.ForWrite);
                    if (attref.Tag == AttTagName)
                    {
                        res = attref.TextString;
                        break;
                    }
                    attref.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                // System.Windows.Forms.MessageBox.Show(ex.Message);
            } 
            return res;
        }

        internal static System.Collections.Hashtable GetAttributeValues(Database db, Transaction trans, BlockReference bref)
        {
            System.Collections.Hashtable htRes = new System.Collections.Hashtable();
            try
            {
                foreach (ObjectId attid in bref.AttributeCollection)
                {
                    AttributeReference attref = (AttributeReference)trans.GetObject(attid, OpenMode.ForWrite);
                    if (!htRes.ContainsKey(attref.Tag))
                    {
                        htRes.Add(attref.Tag, attref.TextString);
                    }
                    attref.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                // System.Windows.Forms.MessageBox.Show(ex.Message);
            } 
            return htRes;
        }

        internal static void SetAttributeValue(Database db, Transaction trans, BlockReference bref, string AttTagName, string AttValue)
        {
            AttTagName = AttTagName.ToUpper();
            try
            {
                foreach (ObjectId attid in bref.AttributeCollection)
                {
                    AttributeReference attref = (AttributeReference)trans.GetObject(attid, OpenMode.ForWrite);
                    if (attref.Tag == AttTagName)
                    {
                        attref.TextString = AttValue;
                        break;
                    }
                    // attref.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                // System.Windows.Forms.MessageBox.Show(ex.Message);
            } 
            
        }

        internal static void SetAttributeValues(Database db, Transaction trans, BlockReference bref, Hashtable htTagValue)
        {
            Hashtable htTV = new Hashtable();
            // force tag names to upper case
            foreach (DictionaryEntry de in htTagValue)
                htTV.Add(de.Key.ToString().Trim().ToUpper(), de.Value.ToString());

            try
            {
                foreach (ObjectId attid in bref.AttributeCollection)
                {
                    AttributeReference attref = (AttributeReference)trans.GetObject(attid, OpenMode.ForRead);
                    if (htTV.ContainsKey(attref.Tag))
                    {
                        attref.UpgradeOpen();
                        attref.TextString = htTV[attref.Tag].ToString();
                    }
                }
            }
            catch (System.Exception ex)
            {
                // System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }

    public class XDTools
    {
        public XDTools()
        {
            Initialize();
        }

        private void Initialize()
        {
            // throw new NotImplementedException();
        }

        public static void SetMyXD(
            Database db,
            Transaction trans,
            ObjectId myOID,
            string appName,
            string fieldName,
            string fieldValue,
            bool ReplaceIfExists)
        {
            // This is special to CES style xdata which stores values in xdata (strings)
            // Values are in the form FIELD_NAME::VALUE 
            // Note the double colons separating field and value

            bool foundField = false;

            // Add the appname if it doesn't already exist
            AddRegAppTableRecord(appName);
            DBObject myObj = (DBObject)trans.GetObject(myOID, OpenMode.ForWrite);
            ResultBuffer rb = myObj.GetXDataForApplication(appName); // .XData;
            ResultBuffer rbAll = myObj.XData;
            if ((rb == null) || (rbAll == null)) // no xdata at all for this object OR no xdata for the appName
            {
                rb = new ResultBuffer(
                    new TypedValue(1001, appName),
                    new TypedValue(1000, fieldName + "::" + fieldValue));
            }
            else // rb is not null and reg app is already in drawing
            {
                // Create typedvalue list so items can be added, removed and replaced
                TypedValueList tvlist = new TypedValueList(rb.AsArray());
                if (ReplaceIfExists == false)
                {
                    tvlist.Add(1000, fieldName + "::" + fieldValue);
                }
                else
                {
                    TypedValue tval = new TypedValue();
                    for (int i = 0; i < tvlist.Count; i++)
                    {
                        tval = tvlist[i];
                        if (tval.TypeCode == 1000)
                        {
                            if (tval.Value.ToString().StartsWith(fieldName + "::"))
                            {
                                foundField = true;
                                tvlist.Insert(i, new TypedValue(1000, fieldName + "::" + fieldValue));
                                int ix = i;
                                ix++;
                                tvlist.RemoveAt(ix);
                            }
                        }
                    }
                }
                if (ReplaceIfExists & (foundField == false))
                {
                    tvlist.Add(1000, fieldName + "::" + fieldValue);
                }
                rb = tvlist;
            }
            // set the xdata for this object
            myObj.XData = rb;

        }

        public static void SetMyXD(
            Database db,
            Transaction trans,
            string appName,
            string fieldName,
            string fieldValue,
            bool ReplaceIfExists)
        {
            // No ObjectID provided - Assumes xdata in ModelSpace
            ObjectId oid = SymbolUtilityServices.GetBlockModelSpaceId(db);
            SetMyXD(db, trans, oid, appName, fieldName, fieldValue, ReplaceIfExists);
        }

        public static void SetMyXD(
            ObjectId myOID,
            string appName,
            string fieldName,
            string fieldValue,
            bool ReplaceIfExists)
        {
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                SetMyXD(db, trans, myOID, appName, fieldName, fieldValue, ReplaceIfExists);
            }
        }
        public static void SetMyXD(
            string appName,
            string fieldName,
            string fieldValue,
            bool ReplaceIfExists)
        {
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                ObjectId myOID = SymbolUtilityServices.GetBlockModelSpaceId(db);
                SetMyXD(db, trans, myOID, appName, fieldName, fieldValue, ReplaceIfExists);
            }
        }

        internal static System.Collections.Hashtable GetMyXDValues(
            Database db,
            Transaction trans,
            ObjectId oid,
            string appName)
        {
            System.Collections.Hashtable htRes = new System.Collections.Hashtable();
            DBObject myObj = (DBObject)trans.GetObject(oid, OpenMode.ForRead);
            ResultBuffer rb = myObj.GetXDataForApplication(appName); // .XData;
            if (rb != null) // found xdata for the appName
            {
                foreach (TypedValue tv in rb)
                {
                    if (tv.TypeCode == 1000)
                    {
                        string tval = tv.Value.ToString();
                        string name = ParseXDFieldName(tval);
                        string val = ParseXDValue(tval);
                        if (!htRes.ContainsKey(name))
                        {
                            htRes.Add(name, val);
                        }
                    }
                }
            }
            return htRes;
        }
        
        public static string GetMyXDValue(
            Database db,
            Transaction trans,
            ObjectId myOID,
            string appName,
            string fieldName)
        {
            string res = string.Empty;
            DBObject myObj = (DBObject)trans.GetObject(myOID, OpenMode.ForRead);
            ResultBuffer rb = myObj.GetXDataForApplication(appName); // .XData;
            if (rb != null) // found xdata for the appName
            {
                foreach (TypedValue tv in rb)
                {
                    if (res.Length < 1)
                    {
                        if (tv.TypeCode == 1000)
                        {
                            string tval = tv.Value.ToString();
                            if (tval.ToUpper().StartsWith(fieldName.ToUpper() + "::"))
                            {
                                res = tval.Substring(tval.IndexOf("::") + 2);
                            }
                        }

                    }
                }
            }
            return res;
        }

        public static string GetMyXDValue(
            ObjectId myOID,
            string appName,
            string fieldName)
        {
            // Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = HostApplicationServices.WorkingDatabase;
            // Returns the first value it finds
            // Or empty string if not found
            string res = string.Empty;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                res = GetMyXDValue(db, trans, myOID, appName, fieldName);
            }
            return res;
        }

        public static string GetMyXDValue(
            string regAppName,
            string fieldName)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            // No ObjectID provided - Assumes xdata in ModelSpace
            // Returns the first value it finds
            // Or empty string if not found
            string res = string.Empty;
            ObjectId oid = SymbolUtilityServices.GetBlockModelSpaceId(db);
            res = GetMyXDValue(oid, regAppName, fieldName);
            return res;
        }

        public static bool MyXDFieldExists(
            ObjectId myOID,
            string regAppName,
            string fieldName)
        {
            bool res = false;
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                res = MyXDFieldExists(db, trans, myOID, regAppName, fieldName);
            }
            return res;
        }

        private static bool MyXDFieldExists(
            Database db,
            Transaction trans,
            ObjectId myOID,
            string regAppName,
            string fieldName)
        {
            bool res = false;
            DBObject myObj = (DBObject)trans.GetObject(myOID, OpenMode.ForRead);
            ResultBuffer rb = myObj.GetXDataForApplication(regAppName); // .XData;
            if (rb != null) // found xdata for the appName
            {
                foreach (TypedValue tv in rb)
                {
                    if (tv.TypeCode == 1000)
                    {
                        string tval = tv.Value.ToString();
                        if (tval.StartsWith(fieldName))
                        {
                            res = true;
                            break;
                        }
                    }
                }
            }
            return res;
        }


        private static void AddRegAppTableRecord(string regAppName)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                RegAppTable rat =
                  (RegAppTable)tr.GetObject(
                    db.RegAppTableId,
                    OpenMode.ForRead,
                    false
                  );
                if (!rat.Has(regAppName))
                {
                    rat.UpgradeOpen();
                    RegAppTableRecord ratr =
                      new RegAppTableRecord();
                    ratr.Name = regAppName;
                    rat.Add(ratr);
                    tr.AddNewlyCreatedDBObject(ratr, true);
                }
                tr.Commit();
            }
        }

        public void ListXD()
        {
            // Ask the user to select an entity
            // for which to retrieve XData
            //PromptEntityOptions opt = new PromptEntityOptions("\nSelect entity: ");
            //PromptEntityResult res = ed.GetEntity(opt);
            //if (res.Status == PromptStatus.OK)
            //{
            //    Transaction tr = doc.TransactionManager.StartTransaction();
            //    using (tr)
            //    {
            //        DBObject obj = tr.GetObject(res.ObjectId, OpenMode.ForRead);
            //        ResultBuffer rb = obj.XData;
            //        frmXData frm = new frmXData();
            //        frm.objID = res.ObjectId;
            //        frm.dgvCESList.Rows.Clear();
            //        frm.dgvOther.Rows.Clear();
            //        if (rb == null)
            //        {
            //            frm.dgvCESList.RowCount = 1;
            //        }
            //        else
            //        {
            //            bool isCES = false;
            //            string code = string.Empty;
            //            string value = string.Empty;

            //            int n = 1;
            //            int nCES = 1;
            //            int nOther = 0;
            //            foreach (TypedValue tv in rb)
            //            {
            //                code = tv.TypeCode.ToString();
            //                value = tv.Value.ToString();
            //                if (code == "1001") // application
            //                {
            //                    isCES = value.Equals("CES");
            //                }
            //                if (isCES)
            //                {
            //                    nCES++;
            //                }
            //                else
            //                {
            //                    nOther++;
            //                }
            //                n++;
            //            }
            //            frm.dgvCESList.RowCount = nCES;
            //            frm.dgvOther.RowCount = nOther;
            //            n = 0;
            //            nCES = 0;
            //            nOther = 0;
            //            foreach (TypedValue tv in rb)
            //            {
            //                code = tv.TypeCode.ToString();
            //                value = tv.Value.ToString();
            //                if (code == "1001") // application
            //                {
            //                    isCES = value.Equals("CES");
            //                }
            //                if (isCES)
            //                {
            //                    frm.dgvCESList[0, nCES].Value = XDTools.ParseXDFieldName(value);
            //                    frm.dgvCESList[1, nCES].Value = XDTools.ParseXDValue(value);
            //                    nCES++;
            //                }
            //                else
            //                {
            //                    frm.dgvOther[0, nOther].Value = code;
            //                    frm.dgvOther[1, nOther].Value = value;
            //                    nOther++;
            //                }
            //                n++;
            //            }

            //        }
            //        System.Windows.Forms.DialogResult dres = acadApp.ShowModalDialog(frm);
            //        if (dres == System.Windows.Forms.DialogResult.OK)
            //        {
            //            // This should save xdata added in the dialog but does not
            //            // reason unknown as of 2-26-2010
            //            ObjectId myID = res.ObjectId;
            //            ArrayList al = new ArrayList();
            //            for (int i = 0; i < frm.dgvCESList.RowCount; i++)
            //            {
            //                try
            //                {
            //                    if (frm.dgvCESList[0, i].Value.ToString().Trim().Length > 0)
            //                    {
            //                        al.Add(frm.dgvCESList[0, i].Value.ToString().Trim() + "::" + frm.dgvCESList[1, i].Value.ToString().Trim());
            //                    }

            //                }
            //                catch (System.Exception eexx)
            //                {


            //                }
            //            }
            //            frm.Close();
            //            string str = "";
            //            using (DocumentLock dlock = doc.LockDocument())
            //            {
            //                for (int i = 0; i < al.Count; i++)
            //                {
            //                    str = (string)al[i];
            //                    string fld = ParseXDFieldName(str);
            //                    string val = ParseXDValue(str);
            //                    SetMyXD(myID, "CES", fld, val, true);
            //                }

            //            }

            //            try
            //            {
            //                rb.Dispose();
            //            }
            //            catch (System.Exception)
            //            {
            //            }
            //        }
            //    }
            //}
        }

        private class TypedValueList : List<TypedValue>
        {
            public TypedValueList(params TypedValue[] args)
            {
                AddRange(args);
            }

            // Make it a bit easier to add items:

            public void Add(int typecode, object value)
            {
                base.Add(new TypedValue(typecode, value));
            }

            // Implicit conversion to SelectionFilter
            //public static implicit operator SelectionFilter(TypedValueList src)
            //{
            //    return src != null ? new SelectionFilter(src) : null;
            //}

            // Implicit conversion to ResultBuffer
            public static implicit operator ResultBuffer(TypedValueList src)
            {
                return src != null ? new ResultBuffer(src) : null;
            }

            // Implicit conversion to TypedValue[] 
            public static implicit operator TypedValue[](TypedValueList src)
            {
                return src != null ? src.ToArray() : null;
            }

            // Implicit conversion from TypedValue[] 
            public static implicit operator TypedValueList(TypedValue[] src)
            {
                return src != null ? new TypedValueList(src) : null;
            }

            // Implicit conversion from SelectionFilter
            //public static implicit operator TypedValueList(SelectionFilter src)
            //{
            //    return src != null ? new TypedValueList(src.GetFilter()) : null;
            //}

            // Implicit conversion from ResultBuffer
            public static implicit operator TypedValueList(ResultBuffer src)
            {
                return src != null ? new TypedValueList(src.AsArray()) : null;
            }

        }

        public static string ParseXDFieldName(string xdString)
        {
            string res = string.Empty;
            int sep = xdString.IndexOf("::");
            if (sep > 0)
            {
                res = xdString.Substring(0, sep);
            }
            return res;
        }
        public static string ParseXDValue(string xdString)
        {
            string res = string.Empty;
            int sep = xdString.IndexOf("::");
            if (sep > 0)
            {
                res = xdString.Substring(xdString.IndexOf("::") + 2);
            }
            return res;
        }

    }

}

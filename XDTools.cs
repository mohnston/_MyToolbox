using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using acColors = Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Internal;

namespace CADSYSUtil
{
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
        Database db = HostApplicationServices.WorkingDatabase;
        Autodesk.AutoCAD.ApplicationServices.Document doc = acadApp.DocumentManager.MdiActiveDocument;
        Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;

        public static void SetMyXD(
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
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = HostApplicationServices.WorkingDatabase;
            using (DocumentLock dlock = doc.LockDocument())
            using (Transaction trans = doc.TransactionManager.StartTransaction())
            {
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
                
                trans.Commit();
            }

        }

        public static void SetMyXD(
            string appName, 
            string fieldName, 
            string fieldValue, 
            bool ReplaceIfExists)
        {
            // No ObjectID provided - Assumes xdata in ModelSpace
            Database db = HostApplicationServices.WorkingDatabase;
            ObjectId oid = SymbolUtilityServices.GetBlockModelSpaceId(db);
            SetMyXD(oid, appName, fieldName, fieldValue, ReplaceIfExists);
        }

        public static string GetMyXDValue(
            ObjectId myOID,
            string appName,
            string fieldName)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            // Returns the first value it finds
            // Or empty string if not found
            string res = string.Empty;
            using (Transaction trans = doc.TransactionManager.StartTransaction())
            {
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
                                if (tval.StartsWith(fieldName + "::"))
                                {
                                    res = tval.Substring(tval.IndexOf("::") + 2);
                                }
                            }
                                    
                        }
                    }
                }
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
            using (Transaction trans = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
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
            public static implicit operator SelectionFilter(TypedValueList src)
            {
                return src != null ? new SelectionFilter(src) : null;
            }

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
            public static implicit operator TypedValueList(SelectionFilter src)
            {
                return src != null ? new TypedValueList(src.GetFilter()) : null;
            }

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

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.ApplicationServices;
using CADHelp;

namespace CADHelp
{
    public class SelectionTools
    {
        // Helpers for getting selection sets or ?
        public static int Cabinet = 1;
        public static int Countertop = 2;
        public static int HItem = 4;
        public static int Toekick = 8;
        public static int Detail = 16;
        public static int Metal = 32;

        public static ObjectIdCollection GetItemsForFinishes(bool UserPick, bool IncludeHF)
        {
            SelectionSet resSS = null;
            ObjectIdCollection resOIDs = new ObjectIdCollection();
            ObjectIdCollection oidsPanels = new ObjectIdCollection();
            TypedValue[] tvals = new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") };
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            // user pre selected items if there are any
            if (ed.SelectImplied().Status == PromptStatus.OK)
            {
                resSS = ed.SelectImplied().Value;
            }
            else
            {
                SelectionFilter filter = new SelectionFilter(tvals);
                if (UserPick)
                {
                    Autodesk.AutoCAD.Internal.Utils.GraphScr();
                    PromptSelectionResult psr = ed.GetSelection(filter);
                    if (psr.Status == PromptStatus.OK)
                    {
                        resSS = psr.Value;
                    }
                }
                else
                {
                    resSS = ed.SelectAll(filter).Value;
                }
            }
            if (resSS != null)
            {
                // filter for my blocks
                List<string> blockNames = new List<string>();
                if (IncludeHF)
                {
                    blockNames.Add("SYM_HITEM*");
                    blockNames.Add("SYM_HFITEM*");
                }
                blockNames.Add("SYM_CAB*");
                blockNames.Add("SYM_CTOP*");
                ObjectIdCollection oids = CADHelper.GetDynamicBlocksByName("dynPanel");
                using (Transaction trans = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId oid in oids)
                    {
                        BlockReference bref = trans.GetObject(oid, OpenMode.ForRead) as BlockReference;
                        if (bref != null)
                        {
                            blockNames.Add(bref.Name);
                        }
                    }
                    ObjectId[] objIDs = resSS.GetObjectIds();
                    int objCount = objIDs.GetUpperBound(0) + 1;
                    if (objCount >= 0)
                    {
                        ObjectIdCollection resIDs = new ObjectIdCollection();
                        for (int i = 0; i < objCount; i++)
                        {
                            ObjectId oid = objIDs[i];
                            BlockReference bref = trans.GetObject(oid, OpenMode.ForRead) as BlockReference;
                            if (bref != null)
                            {
                                if (blockNames.Contains(bref.Name))
                                {
                                    resIDs.Add(oid);
                                }
                                else
                                {
                                    foreach (string val in blockNames)
                                    {
                                        if (val.EndsWith("*"))
                                        {
                                            string s = val.Substring(0, val.Length - 1);
                                            if (bref.Name.StartsWith(s))
                                            {
                                                resIDs.Add(oid);
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if (bref.Name.Equals(val))
                                            {
                                                resIDs.Add(oid);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        resOIDs = resIDs;
                    }
                }
            }

            return resOIDs;
        }

        public static SelectionSet GetAllItemTags(int TagVals, bool UserPick)
        {
            SelectionSet oss;
            Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
            if (TagVals > 63)
            {
                System.Windows.Forms.MessageBox.Show("Error in GetAllItemTags.\nTagVals is " + TagVals.ToString() + " but should be less than 64.", "SelectTools Error");
                oss = ed.SelectImplied().Value;
                return oss;
            }
            
            Collection<TypedValue> collTV = new Collection<TypedValue>();

            TypedValue tvCab = new TypedValue((int)DxfCode.BlockName, "SYM_CAB*");
            TypedValue tvCtop = new TypedValue((int)DxfCode.BlockName, "SYM_CTOP*");
            TypedValue tvHItem = new TypedValue((int)DxfCode.BlockName, "SYM_HITEM*");
            TypedValue tvHFItem = new TypedValue((int)DxfCode.BlockName, "SYM_HFITEM*");
            TypedValue tvToe = new TypedValue((int)DxfCode.BlockName, "SYM_TOE*");
            TypedValue tvDetail = new TypedValue((int)DxfCode.BlockName, "SYM_DETAIL*");
            TypedValue tvMetal = new TypedValue((int)DxfCode.BlockName, "SYM_METAL*");

            collTV.Add(new TypedValue((int)DxfCode.Start, "INSERT"));
            collTV.Add(new TypedValue(-4, "<OR"));
            if (TagVals >= Metal)
            {
                collTV.Add(tvMetal);
                TagVals = TagVals - Metal;
            }
            if (TagVals >= Detail)
            {
                collTV.Add(tvDetail);
                TagVals = TagVals - Detail;
            }
            if (TagVals >= Toekick)
            {
                collTV.Add(tvToe);
                TagVals = TagVals - Toekick;
            }
            if (TagVals >= HItem)
            {
                collTV.Add(tvHItem);
                collTV.Add(tvHFItem);
                TagVals = TagVals - HItem;
            }
            if (TagVals >= Countertop)
            {
                collTV.Add(tvCtop);
                TagVals = TagVals - Countertop;
            }
            if (TagVals == Cabinet)
            {
                collTV.Add(tvCab);
                TagVals = TagVals - Cabinet;
            }
                       
            TypedValue[] filter = new TypedValue[collTV.Count + 1];
            for (int i = 0; i < collTV.Count; i++)
			{
                filter[i] = collTV[i];
			}
            filter[filter.GetUpperBound(0)] = new TypedValue(-4, "OR>");
                
            SelectionFilter ssFilt = new SelectionFilter(filter);
            if (UserPick)
            {
                PromptSelectionOptions selOptions = new PromptSelectionOptions();
                selOptions.MessageForAdding = "Select item tags: ";
                selOptions.AllowDuplicates = true;
                Autodesk.AutoCAD.Internal.Utils.GraphScr();
                PromptSelectionResult res = ed.GetSelection(selOptions, ssFilt);
                oss = res.Value;
            }
            else
            {
                oss = ed.SelectAll(ssFilt).Value;
            }
            return oss;
        }

        public static SelectionSet GetAllItemTags(bool UserPick)
        {
            SelectionSet oss = GetAllItemTags(
                Cabinet + Countertop + Toekick + HItem + Detail + Metal,
                UserPick);
            return oss;
        }

        public static SelectionSet GetCESCabinets(bool UserPick)
        {
            SelectionSet oss;
            Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
            
            Collection<TypedValue> collTV = new Collection<TypedValue>();

            collTV.Add(new TypedValue((int)DxfCode.Start, "INSERT"));
            collTV.Add(new TypedValue((int)DxfCode.BlockName, "CES-*"));
            oss = GetSelection(collTV, UserPick, "Select CES Cabinets");
            return oss;
        }

        public static SelectionSet GetSelection(Collection<TypedValue> collTV, bool UserPick, string PromptString)
        {
            SelectionSet oss;
            Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
            
            // Add code for locked layers
            List<string> lln = CADHelper.LockedLayerNames();
            if (lln.Count > 0)
            {
                collTV.Add(new TypedValue(-4, "<AND"));
                collTV.Add(new TypedValue(-4, "<NOT"));
                collTV.Add(new TypedValue(-4, "<OR"));
                for (int i = 0; i < lln.Count; i++)
                {
                    collTV.Add(new TypedValue((int)DxfCode.LayerName, lln[i]));
                }
                collTV.Add(new TypedValue(-4, "OR>"));
                collTV.Add(new TypedValue(-4, "NOT>"));
                collTV.Add(new TypedValue(-4, "AND>"));
            }

            TypedValue[] filter = new TypedValue[collTV.Count];

            for (int i = 0; i < collTV.Count; i++)
            {
                filter[i] = collTV[i];
            }

            SelectionFilter ssFilt = new SelectionFilter(filter);
            if (UserPick)
            {
                PromptSelectionOptions selOptions = new PromptSelectionOptions();
                selOptions.MessageForAdding = PromptString;
                selOptions.AllowDuplicates = true;
                selOptions.RejectObjectsOnLockedLayers = true;
                Autodesk.AutoCAD.Internal.Utils.GraphScr();
                PromptSelectionResult res = ed.GetSelection(selOptions, ssFilt);
                oss = res.Value;
            }
            else
            {
                oss = ed.SelectAll(ssFilt).Value;
            }
            return oss;
        }

        public static SelectionSet GetSelection(string PromptString)
        {
            SelectionSet oss;
            Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
            PromptSelectionOptions selOptions = new PromptSelectionOptions();
            selOptions.MessageForAdding = PromptString;
            selOptions.AllowDuplicates = true;
            Autodesk.AutoCAD.Internal.Utils.GraphScr();
            PromptSelectionResult res = ed.GetSelection(selOptions);
            oss = res.Value;
            return oss;
        }

        public static SelectionSet GetDimensions(bool UserPick, string PromptString)
        {
            SelectionSet oss;
            Collection<TypedValue> collTV = new Collection<TypedValue>();
            collTV.Add(new TypedValue((int)DxfCode.Start, "DIMENSION"));
            oss = GetSelection(collTV, UserPick, PromptString);
            return oss;
        }

        internal static SelectionSet GetText(bool UserPick, string PromptString)
        {
            SelectionSet oss;
            Collection<TypedValue> collTV = new Collection<TypedValue>();
            collTV.Add(new TypedValue((int)DxfCode.Operator, "<OR"));
            collTV.Add(new TypedValue((int)DxfCode.Start, "TEXT"));
            collTV.Add(new TypedValue((int)DxfCode.Start, "MTEXT"));
            collTV.Add(new TypedValue((int)DxfCode.Operator, "OR>"));
            oss = GetSelection(collTV, UserPick, PromptString);
            return oss;
        }

        //public static Table GetTableByName(string TableName)
        //{
        //    Table resTbl = null;
        //    SelectionSet tss = GetTables(false, "");
            
        //    for (int i = 0; i < tss.Count; i++)
        //    {
        //        if (XDTools.GetMyXDValue(tss[i].ObjectId,"CES","TableName") == TableName)
        //        {
        //            //resTbl = tss[i];
        //            break;
        //        } 
        //    }

        //    return resTbl;
        //}

        public static SelectionSet GetTables(bool UserPick, string PromptString)
        {
            SelectionSet oss;
            Collection<TypedValue> collTV = new Collection<TypedValue>();
            collTV.Add(new TypedValue((int)DxfCode.Start, "ACAD_TABLE"));
            oss = GetSelection(collTV, UserPick, PromptString);
            return oss;
        }

        public static IEnumerable<Table> GetTables(bool UserPick, string PromptString, bool InModelSpace)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            IEnumerable<Table> tlist = Autodesk.AutoCAD.Linq.AcDBExtensions.GetAcadTableObjects(db, true);
            return tlist;
        }

        internal static string GetBlockName(ObjectId SpecBoxRefID)
        {
            string res = string.Empty;
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            { 
                BlockReference bref = trans.GetObject(SpecBoxRefID, OpenMode.ForRead) as BlockReference;
                if (bref != null)
                    res = bref.Name;
            }
            return res;
        }
    }

    /// <summary>
    ///
    /// This class merges the functionality of a
    /// collection of DBObjects and a Transaction.
    ///
    /// Practically, the main difference between
    /// using this class and a transaction, is
    /// that repeated calls to GetObject() on this
    /// class with the same ObjectId, returns the
    /// same managed wrapper instance, rather than
    /// a new one on each call as Transaction does.
    ///
    /// Like a Transaction, any DBObject you get
    /// from this class, should not be used after
    /// Commit(), Abort() or Dispose() is called.
    ///
    /// </summary>

    // DBObjectSet.cs  copyright(c) 2007  Tony Tanzillo www.caddzone.com
    //
    // Original source location:
    //
    //  http://www.caddzone.com/DBObjectSet.cs
    //
    // LICENSE:
    //
    // This software may not be published or reproduced in source
    // form without the express written consent of CADDZONE.COM
    //
    // Fair use:
    //
    // Permission to use this software in object code form, for
    // private software development purposes, and without fee is
    // hereby granted, provided that the above copyright notice
    // appears in all copies and that both that copyright notice
    // and limited warranty and restricted rights notice below
    // appear in all supporting documentation.
    //
    // CADDZONE.COM PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
    // CADDZONE.COM SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
    // MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  CADDZONE.COM
    // DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
    // UNINTERRUPTED OR ERROR FREE.
    //
    // Use, duplication, or disclosure by the U.S. Government is subject to
    // restrictions set forth in FAR 52.227-19 (Commercial Computer
    // Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
    // (Rights in Technical Data and Computer Software), as applicable.

    public class DBObjectSet<T> : ICollection<T>, IDisposable where T : DBObject
    {
        private Collection items = new Collection();
        private Transaction trans = null;
        private Database db = null;

        private class Collection : KeyedCollection<ObjectId, T>
        {
            protected override ObjectId GetKeyForItem(T item)
            {
                return item.ObjectId;
            }
        }

        public DBObjectSet(Database database)
        {
            this.db = database;
        }

        public DBObjectSet(IEnumerable<ObjectId> ids, OpenMode mode)
        {
            IEnumerator<ObjectId> enumerator = ids.GetEnumerator();
            if (!enumerator.MoveNext())
                throw new ArgumentException("Empty array or collection");
            ObjectId id = enumerator.Current;
            if (id.IsNull)
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId);
            this.db = id.Database;
            enumerator.Dispose();
            AddRange(ids, mode);
        }

        public DBObjectSet(ObjectIdCollection ids, OpenMode mode)
        {
            if (ids.Count == 0)
                throw new InvalidOperationException("Empty collection");
            if (ids[0].IsNull)
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId);
            this.db = ids[0].Database;
            AddRange(ids, mode);
        }

        private Transaction Transaction
        {
            get
            {
                if (trans == null)
                    trans = db.TransactionManager.StartTransaction();
                return trans;
            }
        }

        public void Commit()
        {
            Clear(true);
        }

        public void Abort()
        {
            Clear(false);
        }

        public T GetObject(ObjectId id, OpenMode mode)
        {
            if (items.Contains(CheckId(id)))
            {
                T item = items[id];
                if (!item.IsWriteEnabled && mode == OpenMode.ForWrite)
                    item.UpgradeOpen();
                return item;
            }
            T obj = Transaction.GetObject(id, mode, mode == OpenMode.ForWrite) as T;
            if (obj != null)
                items.Add(obj);
            return obj;
        }

        public int AddRange(IEnumerable<ObjectId> ids, OpenMode mode)
        {
            int c = this.Count;
            foreach (ObjectId id in ids)
                this.GetObject(id, mode);
            return this.Count - c;
        }

        public int AddRange(ObjectIdCollection ids, OpenMode mode)
        {
            int c = this.Count;
            foreach (ObjectId id in ids)
                this.GetObject(id, mode);
            return this.Count - c;
        }

        #region ICollection<T> Members

        // AddNewlyCreatedDBObject:

        public void Add(T item)
        {
            if (items.Contains(item))
                throw new InvalidOperationException("Item already a member");
            if (!item.IsNewObject)
                throw new InvalidOperationException("Item is already database-resident");
            Transaction.AddNewlyCreatedDBObject(item, true);
            items.Add(item);
        }

        public void Clear(bool commit)
        {
            if (trans != null)
            {
                if (commit)
                    trans.Commit();
                else
                    trans.Abort();
                trans = null;
            }
            items.Clear();
        }

        public void Clear()
        {
            Clear(false);
        }

        private ObjectId CheckId(ObjectId id)
        {
            if (id.IsNull)
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId);
            if (id.Database != this.db)
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.WrongDatabase);
            return id;
        }

        // Operate on a subset of TObject:

        void ForEach<TObject>(Action<TObject> action) where TObject : T
        {
            foreach (T item in this)
            {
                TObject obj = item as TObject;
                if (obj != null)
                    action(obj);
            }
        }

        public IList<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter)
        {
            List<TOutput> list = new List<TOutput>(this.Count);
            foreach (T item in items)
                list.Add(converter(item));
            list.TrimExcess();
            return list;
        }

        public IList<TOutput> ConvertSome<TInput, TOutput>(Converter<TInput, TOutput> converter)
           where TInput : T
        {
            List<TOutput> list = new List<TOutput>(this.Count);
            foreach (T item in items)
            {
                TInput input = item as TInput;
                if (input != null)
                    list.Add(converter(input));
            }
            list.TrimExcess();
            return list;
        }

        public bool Exists(Predicate<T> predicate)
        {
            foreach (T item in this.items)
                if (predicate(item))
                    return true;
            return false;
        }

        public bool TrueForAll(Predicate<T> predicate)
        {
            foreach (T item in this.items)
                if (!predicate(item))
                    return false;
            return true;
        }

        public T Find(Predicate<T> predicate)
        {
            foreach (T item in items)
                if (predicate(item))
                    return item;
            return null;
        }

        public int Sum(Converter<T, int> converter)
        {
            int result = 0;
            foreach (T item in items)
                result += converter(item);
            return result;
        }

        public double Sum(Converter<T, double> converter)
        {
            double result = 0.0;
            foreach (T item in items)
                result += converter(item);
            return result;
        }

        public bool Contains(T item)
        {
            return items.Contains(item);
        }

        public void CopyTo(T[] array, int arrayIndex)
        {
            items.CopyTo(array, arrayIndex);
        }

        public int Count
        {
            get
            {
                return items.Count;
            }
        }

        public bool IsReadOnly
        {
            get
            {
                return false;
            }
        }

        public bool Remove(T item)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region IEnumerable<T> Members

        public IEnumerator<T> GetEnumerator()
        {
            return items.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
        }

        private void Dispose(bool disposing)
        {
            if (disposing && !disposed)
            {
                Clear();
                GC.SuppressFinalize(this);
                disposed = true;
            }
        }

        private bool disposed = false;

        ~DBObjectSet()
        {
            CheckDisposed();
            Dispose(false);
        }

        void CheckDisposed()
        {
            if (!disposed)
            {
                Console.Beep();
                System.Diagnostics.Trace.WriteLine(
                   string.Format(
                      "instances of {0} should always be deterministically disposed",
                      this.GetType().Name), "Warning");
            }
        }

        #endregion
    }

}

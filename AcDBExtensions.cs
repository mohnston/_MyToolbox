using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

using System.IO;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// This code is by James Johnson from his presentation "The Weakest Linq"
// I changed the namespace name because I don't like weakness
namespace Autodesk.AutoCAD.Linq
{
    public static class AcDBExtensions
    {
        /// <summary>
        /// Returns BlocktableRecord of dtatabase using TopTransaction
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static BlockTableRecord GetBlockTableRecord(this Database _db, bool mspace)
        {
            string space = mspace ? BlockTableRecord.ModelSpace : BlockTableRecord.PaperSpace;

            BlockTable bt = (BlockTable)_db.TransactionManager.TopTransaction.GetObject(_db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)_db.TransactionManager.TopTransaction.GetObject(bt[space], OpenMode.ForRead);

            return btr;
        }

        /// <summary>
        /// Returns IEnumerable collection of text objects
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<DBText> GetDBTextObjects(this Database _db, bool mspace)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<DBText> dbTxt = btr.Cast<ObjectId>()
                                           .Where<ObjectId>(oid => (oid.ObjectClass.DxfName == "TEXT"))
                                           .Select<ObjectId, DBText>(oid => (DBText)(oid.GetObject(OpenMode.ForRead, false)))
                                           .ToList();
            return dbTxt;
        }

        /// <summary>
        /// Returns IEnumerable collection of table objects
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<Table> GetAcadTableObjects(this Database _db, bool mspace)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<Table> AcadTable = btr.Cast<ObjectId>()
                                           .Where<ObjectId>(oid => (oid.ObjectClass.DxfName == "ACAD_TABLE"))
                                           .Select<ObjectId, Table>(oid => (Table)(oid.GetObject(OpenMode.ForRead, false)))
                                           .ToList();
            return AcadTable;
        }

       /// <summary>
       /// Get DBText Strings
       /// </summary>
       /// <param name="_db"></param>
       /// <param name="mspace"></param>
       /// <returns></returns>
        public static IEnumerable<string> GetDBTextValues2(this Database _db, bool mspace)
        {

            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<string> dbTxt = btr.Cast<ObjectId>()
                                    .ToList<ObjectId>()
                                    .ConvertAll<DBObject>(oid => { return _db.TransactionManager.TopTransaction.GetObject(oid, OpenMode.ForRead, false); })
                                    .Where<DBObject>(obj => obj.GetType() == typeof(DBText))
                                    .Cast<DBText>()
                                    .ToList()
                                    .ConvertAll<string>(DBTXT => { return DBTXT.TextString; })
                                    .AsEnumerable();

            return dbTxt;
        }

        /// <summary>
        /// Get DBText Strings
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<string> GetDBTextValues(this Database _db, bool mspace)
        {

            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<string> dbTxt = btr.Cast<ObjectId>()
                                    .Select<ObjectId, DBObject>(oid =>  _db.TransactionManager.TopTransaction.GetObject(oid, OpenMode.ForRead, false) )
                                    .Where<DBObject>(obj => obj.GetType() == typeof(DBText))
                                    .Cast<DBText>()
                                    .Select<DBText,string>(dbt => dbt.TextString)
                                    .AsEnumerable();

            return dbTxt;
        }


        /// <summary>
        /// Sorts DBText objects by property
        /// Sample Usage:
        /// List<string> textvalues2 = db.SortDBText("TextString", true, true).ToList().ConvertAll<string>(DBTXT => { return DBTXT.TextString; });
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="propertyToSortBy"></param>
        /// <param name="sortDirection"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<DBText> SortDBText(this Database _db, string propertyToSortBy, bool sortAscending, bool mspace)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<DBText> dbTxt = _db.GetDBTextObjects(mspace);

            var param = Expression.Parameter(typeof(DBText), "item");

            var sortExpression = Expression.Lambda<Func<DBText, object>>
                (Expression.Convert(Expression.Property(param, propertyToSortBy), typeof(object)), param);

            if (sortAscending)
            {
                return dbTxt.AsQueryable<DBText>().OrderBy<DBText, object>(sortExpression);
            }
            else
            {
                return dbTxt.AsQueryable<DBText>().OrderByDescending<DBText, object>(sortExpression);
            }
        }


        /// <summary>
        /// Gets the extents of the layers on specified layer(s)
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="layerNames"></param>
        /// <returns></returns>
        public static Extents3d GetExtentsOfEntitiesOnLayer(this Database _db, params string[] layerNames)
        {
            //Extents3d resultExtents = new Extents3d();
            double MinX = 0, MinY = 0, MinZ = 0;
            double MaxX = 0, MaxY = 0, MaxZ = 0;

            BlockTableRecord btr = _db.GetBlockTableRecord(true);
            List<Entity> objOnLayer = btr.Cast<ObjectId>()
                                           .ToList<ObjectId>()
                                           .ConvertAll<Entity>(oid => { return _db.TransactionManager.TopTransaction.GetObject(oid, OpenMode.ForRead, false) as Entity; })
                                           .Where<Entity>(ent => layerNames.Any<string>(lay => lay == ent.Layer))
                                           .ToList();

            List<Extents3d> objExtents = objOnLayer.ConvertAll<Extents3d>(ent => { return ent.GeometricExtents; });
            MinX = objExtents.Min<Extents3d>(ext => ext.MinPoint.X);
            MinY = objExtents.Min<Extents3d>(ext => ext.MinPoint.Y);
            MinZ = objExtents.Min<Extents3d>(ext => ext.MinPoint.Z);

            MaxX = objExtents.Max<Extents3d>(ext => ext.MaxPoint.X);
            MaxY = objExtents.Max<Extents3d>(ext => ext.MaxPoint.Y);
            MaxZ = objExtents.Max<Extents3d>(ext => ext.MaxPoint.Z);

            return new Extents3d(new Point3d(MinX, MinY, MinZ), new Point3d(MaxX, MaxY, MaxZ));
        }


        /// <summary>
        /// Returns IEnumerable collection of specified DXF objects
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<ObjectId> GetDBdxfObjectIDs(this Database _db, bool mspace, string dxfType)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<ObjectId> dbObjCollection = btr.Cast<ObjectId>()
                                           .Where<ObjectId>(oid => (oid.ObjectClass.DxfName == dxfType))
                                           .ToList();

            return dbObjCollection;
        }


        /// <summary>
        /// Returns IEnumerable collection of specified  objects
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<ObjectId> DBTextObjectIdCollection(this Database _db, bool mspace)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<ObjectId> dbObjCollection = btr.Cast<ObjectId>()
                                           .Where<ObjectId>(oid => (oid.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(DBText)))))
                                          .AsEnumerable();

            return dbObjCollection;
        }

        /// <summary>
        /// Returns IEnumerable collection of specified  objects
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<ObjectId> CircleObjectIdCollection(this Database _db, bool mspace)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            btr.Cast<ObjectId>().Where<ObjectId>(oid => (oid.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Circle)))));

            IEnumerable<ObjectId> dbObjCollection = btr.Cast<ObjectId>()
                                           .Where<ObjectId>(oid => (oid.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Circle)))))
                                          .AsEnumerable();

            return dbObjCollection;
        }

        /// <summary>
        /// Returns IEnumerable collection of specified  objects
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<ObjectId> LineObjectIdCollection(this Database _db, bool mspace)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<ObjectId> dbObjCollection = btr.Cast<ObjectId>()
                                           .Where<ObjectId>(oid => (oid.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Line)))))
                                          .AsEnumerable();

            return dbObjCollection;
        }

        /// <summary>
        /// Returns IEnumerable collection of specified  objects
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static IEnumerable<ObjectId> PolylineObjectIdCollection(this Database _db, bool mspace)
        {
            BlockTableRecord btr = _db.GetBlockTableRecord(mspace);
            IEnumerable<ObjectId> dbObjCollection = btr.Cast<ObjectId>()
                                           .Where<ObjectId>(oid => (oid.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Polyline)))))
                                          .AsEnumerable();

            return dbObjCollection;
        }

        /// <summary>
        /// Gets Last ObjectID in database by model or paper space
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static ObjectId LastObjectID(this Database _db, bool mspace)
        {
            return _db.GetBlockTableRecord(mspace).Cast<ObjectId>().lastObjectID();
        }

        /// <summary>
        /// Gets First ObjectID in database by model or paper space
        /// </summary>
        /// <param name="_db"></param>
        /// <param name="mspace"></param>
        /// <returns></returns>
        public static ObjectId FirstObjectID(this Database _db, bool mspace)
        {
            return _db.GetBlockTableRecord(mspace).Cast<ObjectId>().firstObjectID();
        }

        /// <summary>
        /// Last object created ObjectID in ObjectIDCollection
        /// </summary>
        /// <param name="colObjectID"></param>
        /// <returns></returns>
        public static ObjectId lastObjectID(this ObjectIdCollection colObjectID)
        {
            return colObjectID.Cast<ObjectId>().lastObjectID();
        }

        /// <summary>
        /// First object created ObjectID in ObjectIDCollection
        /// </summary>
        /// <param name="colObjectID"></param>
        /// <returns></returns>
        public static ObjectId firstObjectID(this ObjectIdCollection colObjectID)
        {
            return colObjectID.Cast<ObjectId>().firstObjectID();
        }

        /// <summary>
        /// Gets Last object created in IEnumerable list/array 
        /// </summary>
        /// <param name="oidList"></param>
        /// <returns></returns>
        public static ObjectId lastObjectID(this IEnumerable<ObjectId> oidList)
        {
            if (oidList.Count<ObjectId>() > 0)
            {
                return oidList.OrderBy(oid => (oid.Handle.Value)).Last();
            }
            return ObjectId.Null;
        }

        /// <summary>
        /// Gets Oldest or first object on IEnumerable list/array
        /// </summary>
        /// <param name="oidList"></param>
        /// <returns></returns>
        public static ObjectId firstObjectID(this IEnumerable<ObjectId> oidList)
        {
            if (oidList.Count<ObjectId>() > 0)
            {
                return oidList.OrderBy(oid => (oid.Handle.Value)).First();
            }
            return ObjectId.Null;
        }

        /// <summary>
        /// Change case to title case
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static string ChangeCase(this string source)
        {
            return new System.Globalization.CultureInfo("en-US").TextInfo.ToTitleCase(source);
        }

        public static double Height(this Extents3d source)
        {
            return source.MaxPoint.Y - source.MinPoint.Y;
        }
        public static double Width(this Extents3d source)
        {
            return source.MaxPoint.X - source.MinPoint.X;
        }

    }
}

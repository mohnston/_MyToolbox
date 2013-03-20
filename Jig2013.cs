using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.GraphicsInterface;

// using System.Runtime.InteropServices;

namespace COMJig
{
    // [ComVisible(true)]
    public class Jig // : System.EnterpriseServices.ServicedComponent
    {
        Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Database db = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.WorkingDatabase;
        Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

        public Jig()
        {
            
        }

        public bool DragByHandle(Handle HandleToEntity, Point3d GripPoint)
        {
            string sHandle = HandleToEntity.ToString();
            string sPoint = GripPoint.X.ToString() + ", " + GripPoint.Y.ToString() + ", " + GripPoint.Z.ToString();
            return DragByHandle(sHandle, sPoint);
        }

        public bool DragByHandle(Handle HandleToEntity, string GripPoint)
        {
            string sHandle = HandleToEntity.ToString();
            return DragByHandle(sHandle, GripPoint);
        }

        public bool DragByHandle(string HandleToEntity, string GripPoint)
        {
            bool res = false;
            ObjectId oid = ObjectIDFromHandle(HandleToEntity);
            if (oid == ObjectId.Null) return false;

            string[] comma = { "," };
            string[] xyz = GripPoint.Split(comma, System.StringSplitOptions.RemoveEmptyEntries);
            Point3d mtLoc = Point3d.Origin;

            if (xyz.GetUpperBound(0) == 2)
            {
                double x;
                double y;
                double z;
                if (double.TryParse(xyz[0], out x))
                {
                    if (double.TryParse(xyz[1], out y))
                    {
                        if (double.TryParse(xyz[2], out z))
                        {
                            mtLoc = new Point3d(x, y, z);
                        }
                    }
                }
            }
            try
            {
                using (DocumentLock dlock = doc.LockDocument())
                using (Transaction tr = doc.TransactionManager.StartTransaction()) 
                {
                    if (!oid.IsNull)
                    {
                        ObjectId[] oids = { oid };

                        Entity ent = (Entity)tr.GetObject(oid, OpenMode.ForWrite);

                        SelectionSet ss = Autodesk.AutoCAD.EditorInput.SelectionSet.FromObjectIds(oids);

                        PromptPointResult ppr = ed.Drag(
                            ss,
                            "\nSelect location: ",
                            delegate(Point3d pt, ref Matrix3d mat)
                            {
                                // If no change has been made, say so
                                if (mtLoc == pt)
                                    return SamplerStatus.NoChange;
                                else
                                {
                                    // Otherwise we return the displacement
                                    // matrix for the current position
                                    mat = Matrix3d.Displacement(mtLoc.GetVectorTo(pt));
                                }
                                return SamplerStatus.OK;
                            }
                        );

                        // Assuming it works, transform our thing
                        if (ppr.Status == PromptStatus.OK)
                        {
                            // Get the final translation matrix
                            Matrix3d mat = Matrix3d.Displacement(mtLoc.GetVectorTo(ppr.Value));
                            // Transform our Entity
                            ent.TransformBy(mat);
                            // Finally we commit our transaction
                            tr.Commit();
                            res = true;
                        }
                    }
                } // transaction
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                ed.WriteMessage("\nCommand: ");
            }
            return res;
        }

        public string InsertBlock(string BlockNameOrPath)
        {
            string blockName = BlockNameOrPath;
            string res = string.Empty;
            try
            {
                using (DocumentLock dlock = doc.LockDocument())
                {
                    Transaction tr = doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        // Test to see if block definition exists in drawing
                        if (GetBlockID(blockName).Equals(ObjectId.Null))
                        {
                            // Block definition doesn't exist in drawing
                            // Try to add it
                            blockName = AddBlockDefFromFileOrName(blockName);
                            // How about now? Does block definition exist?
                            if (GetBlockID(blockName).Equals(ObjectId.Null))
                            {
                                // Block definition doesn't exist even after we tried to insert it
                                ed.WriteMessage("\nBlock \"" + blockName + "\" not found.");
                                return string.Empty;
                            }
                        }

                        BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForRead);

                        // Block needs to be inserted to current space before
                        // being able to append attribute to it
                        // Insert the block reference
                        BlockReference br = new BlockReference(new Point3d(), btr.ObjectId);
                        space.AppendEntity(br);
                        tr.AddNewlyCreatedDBObject(br, true);

                        #region Collect attribute data for use later
                        Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();
                        if (btr.HasAttributeDefinitions)
                        {
                            foreach (ObjectId id in btr)
                            {
                                DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                                AttributeDefinition ad = obj as AttributeDefinition;

                                if (ad != null && !ad.Constant)
                                {
                                    AttributeReference ar = new AttributeReference();

                                    ar.SetAttributeFromBlock(ad, br.BlockTransform);
                                    ar.Position = ad.Position.TransformBy(br.BlockTransform);

                                    if (ad.Justify != AttachmentPoint.BaseLeft)
                                    {
                                        ar.AlignmentPoint = ad.AlignmentPoint.TransformBy(br.BlockTransform);
                                    }
                                    if (ar.IsMTextAttribute)
                                    {
                                        ar.UpdateMTextAttribute();
                                    }

                                    ar.TextString = ad.TextString;

                                    ObjectId arId = br.AttributeCollection.AppendAttribute(ar);
                                    tr.AddNewlyCreatedDBObject(ar, true);

                                    // Initialize our dictionary with the ObjectId of
                                    // the attribute reference + attribute definition info
                                    attInfo.Add(arId, new AttInfo(ad.Position, ad.AlignmentPoint, ad.Justify != AttachmentPoint.BaseLeft));
                                }
                            }
                        }

                        #endregion
                        
                        // Run the jig
                        BlockJig myJig = new BlockJig(tr, br, attInfo);
                        if (myJig.Run() != PromptStatus.OK)
                            return string.Empty;
                        else
                            res = br.Handle.ToString();
                        // Commit changes if user accepted, otherwise discard
                        tr.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                ed.WriteMessage("\nCommand: ");
            }
            return res;
        }

        private string AddBlockDefFromFileOrName(string blockName)
        {
            // this is called while doc is locked and there is a transaction running
            // the block name did not exist in the block table
            string res = string.Empty;
            string blkName = blockName;
            string fullPath = string.Empty;
            //
 
            fullPath = SymbolUtilityServices.GetInsertPathNameFromBlockName(blkName);
            if (fullPath.Length > 0) // it was a block name and the blocks path was found
            {
                try
                {
                    fullPath = HostApplicationServices.Current.FindFile(fullPath, db, FindFileHint.Default);

                }
                catch (Exception ex)
                {
                    // res remains empty string
                }

            }
            else // block name was not in search path; probably a full path; fullPath was string empty
            {
                try
                {
                    fullPath = HostApplicationServices.Current.FindFile(blkName, db, FindFileHint.Default);
                }
                catch (Exception ex)
                {
                    // res remains empty string
                }
            }

            if (fullPath.Length > 0)
            {

                blkName = SymbolUtilityServices.GetBlockNameFromInsertPathName(fullPath);

                // Try opening the drawing file
                // It should fail if the file is bad
                try
                {
                    Database bdb = new Database(false, false);
                    bdb.ReadDwgFile(fullPath, System.IO.FileShare.Read, true, string.Empty);
                    db.Insert(blkName, bdb, true);
                    bdb.CloseInput(true);
                    bdb.Dispose();
                    res = blkName;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception acex)
                {
                    // for debugging tap into acex
                    // for production just return an empty string
                    // throw;
                }
            }
            return res;
        }

        private ObjectId GetBlockID(string BlockName)
        {
            // Use this instead of "Has" because the erased objects may give a false positive
            // Database db = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.WorkingDatabase;
            ObjectId resID = ObjectId.Null;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (blocks.Has(BlockName))
                {
                    resID = blocks[BlockName];
                    if (resID.IsErased)
                    {
                        foreach (ObjectId btrID in blocks)
                        {
                            if (!btrID.IsErased)
                            {
                                BlockTableRecord res = (BlockTableRecord)tr.GetObject(btrID, OpenMode.ForRead);
                                if (string.Compare(res.Name, BlockName, true) == 0)
                                {
                                    resID = btrID;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return resID;
        }

        private ObjectId ObjectIDFromHandle(Handle hand)
        {
            ObjectId resoid = ObjectId.Null;
            // Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock dlock = doc.LockDocument())
            {
                try
                {
                    resoid = doc.Database.GetObjectId(false, hand, 0);
                }
                catch (Exception ex)
                {

                }
            }
            return resoid;
        }

        private ObjectId ObjectIDFromHandle(string HandleToEntity)
        {
            ObjectId resoid = ObjectId.Null;
            // Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock dlock = doc.LockDocument())
            {
                try
                {
                    long ln = System.Convert.ToInt64(HandleToEntity, 16);
                    Handle hand = new Handle(ln);
                    resoid = db.GetObjectId(false, hand, 0);
                }
                catch (Exception ex)
                {

                }
            }
            return resoid;
        }

    }

    class AttInfo
    {
        private Point3d _pos;
        private Point3d _aln;
        private bool _aligned;

        public AttInfo(Point3d pos, Point3d aln, bool aligned)
        {
            _pos = pos;
            _aln = aln;
            _aligned = aligned;
        }

        public Point3d Position
        {
            set
            {
                _pos = value;
            }
            get
            {
                return _pos;
            }
        }

        public Point3d Alignment
        {
            set
            {
                _aln = value;
            }
            get
            {
                return _aln;
            }
        }

        public bool IsAligned
        {
            set
            {
                _aligned = value;
            }
            get
            {
                return _aligned;
            }
        }
    }

    class BlockJig : EntityJig
    {
        private Point3d _pos;
        private Dictionary<ObjectId, AttInfo> _attInfo;
        private Transaction _tr;

        public BlockJig(Transaction tr, BlockReference br, Dictionary<ObjectId, AttInfo> attInfo) : base(br)
        {
            _pos = br.Position;
            _attInfo = attInfo;
            _tr = tr;
        }

        protected override bool Update()
        {
            BlockReference br = Entity as BlockReference;

            br.Position = _pos;

            if (br.AttributeCollection.Count != 0)
            {
                foreach (ObjectId id in br.AttributeCollection)
                {
                    DBObject obj = _tr.GetObject(id, OpenMode.ForRead);
                    AttributeReference ar = obj as AttributeReference;

                    // Apply block transform to att def position
                    if (ar != null)
                    {
                        ar.UpgradeOpen();
                        AttInfo ai = _attInfo[ar.ObjectId];
                        ar.Position = ai.Position.TransformBy(br.BlockTransform);
                        if (ai.IsAligned)
                        {
                            ar.AlignmentPoint = ai.Alignment.TransformBy(br.BlockTransform);
                        }
                        if (ar.IsMTextAttribute)
                        {
                            ar.UpdateMTextAttribute();
                        }
                    }
                }
            }
            return true;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions opts = new JigPromptPointOptions("\nSelect insertion point:");
            opts.BasePoint = new Point3d(0, 0, 0);
            opts.UserInputControls = UserInputControls.NoZeroResponseAccepted;

            PromptPointResult ppr = prompts.AcquirePoint(opts);

            if (_pos == ppr.Value)
            {
                return SamplerStatus.NoChange;
            }

            _pos = ppr.Value;

            return SamplerStatus.OK;
        }

        public PromptStatus Run()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptResult promptResult = ed.Drag(this);
            return promptResult.Status;
        }
    }
    
    public class JigHelper
    {
        static Database db = acadApp.DocumentManager.MdiActiveDocument.Database;
        static Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;

        private JigHelper()
        {
        }

        public static void DragBlock(ObjectId BlockDefID)
        {
            ObjectId block = BlockDefID;
            Vector3d Normal = db.Ucsxdir.CrossProduct(db.Ucsydir);
            if (block.IsNull)
            {
                ed.WriteMessage("\nBlock not found.");
                return;
            }
            JigBlock jig = new JigBlock(block, Point3d.Origin, Normal.GetNormal(), new Scale3d(1, 1, 1), false);
            PromptResult res = ed.Drag(jig);
            if (res.Status == PromptStatus.OK)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);
                    btr.AppendEntity(jig.BlockReference);
                    BlockTableRecord dbtr = (BlockTableRecord)tr.GetObject(block, OpenMode.ForWrite);
                    if (dbtr.HasAttributeDefinitions)
                    {
                        foreach (ObjectId boid in dbtr)
                        {
                            Entity ent = (Entity)tr.GetObject(boid, OpenMode.ForWrite);
                            if (ent is AttributeDefinition)
                            {
                                AttributeDefinition adef = (AttributeDefinition)ent;
                                AttributeReference aref = new AttributeReference();

                                aref.SetAttributeFromBlock(adef, jig.BlockReference.BlockTransform);
                                jig.BlockReference.AttributeCollection.AppendAttribute(aref);
                                tr.AddNewlyCreatedDBObject(aref, true);
                            }
                        }
                    }
                    tr.AddNewlyCreatedDBObject(jig.BlockReference, true);
                    tr.Commit();
                }
            }
        }

        public static void DragInsertDemo()
        {
            PromptStringOptions opts = new PromptStringOptions("\nBlock name: ");
            opts.AllowSpaces = true;
            Vector3d Normal = db.Ucsxdir.CrossProduct(db.Ucsydir);
            using (DocumentLock docLock = ed.Document.LockDocument())
            {
                PromptResult res = ed.GetString(opts);
                if (res.Status == PromptStatus.OK && res.StringResult.Trim() != string.Empty)
                {
                    ObjectId block = GetBlockId(db, res.StringResult);
                    if (block.IsNull)
                    {
                        ed.WriteMessage("\nBlock {0} not found.", res.StringResult);
                        return;
                    }
                    JigBlock jig = new JigBlock(block, Point3d.Origin, Normal.GetNormal(), new Scale3d(1, 1, 1), false);
                    res = ed.Drag(jig);
                    if (res.Status == PromptStatus.OK)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);
                            btr.AppendEntity(jig.BlockReference);
                            tr.AddNewlyCreatedDBObject(jig.BlockReference, true);
                            tr.Commit();
                        }
                    }
                }
            }
        }

        private static ObjectId GetBlockId(Database db, string Name)
        {
            ObjectId id = ObjectId.Null;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blocks = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (blocks.Has(Name))
                {
                    id = blocks[Name];
                    if (id.IsErased)
                    {
                        foreach (ObjectId btrId in blocks)
                        {
                            if (!id.IsErased)
                            {
                                BlockTableRecord rec = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                if (string.Compare(rec.Name, Name, true) == 0)
                                {
                                    id = btrId;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return id;
        }
    }

    public class JigBlock : EntityJig
    {
        public JigBlock(ObjectId BlockId, Point3d Position, Vector3d Normal, Scale3d Scale, bool FlipX)
            : base(new BlockReference(Position, BlockId))
        {
            BlockReference.Normal = Normal;
            BlockReference.ScaleFactors = Scale;
            if (FlipX) BlockReference.ScaleFactors = new Scale3d((BlockReference.ScaleFactors.X * -1), BlockReference.ScaleFactors.Y, BlockReference.ScaleFactors.Z);
            position = Position;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jigOpts = new JigPromptPointOptions();
            jigOpts.UserInputControls = UserInputControls.Accept3dCoordinates;
            jigOpts.Message = "\nInsertion point: ";
            PromptPointResult res = prompts.AcquirePoint(jigOpts);
            Point3d curPoint = res.Value;
            if (position.DistanceTo(curPoint) > 1.0e-4)
                position = curPoint;
            else
                return SamplerStatus.NoChange;

            if (res.Status == PromptStatus.Cancel)
                return SamplerStatus.Cancel;
            else
                return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            try
            {
                if (this.BlockReference.Position.DistanceTo(position) > 1.0e-4)
                {
                    this.BlockReference.Position = position;
                    return true;
                }
            }
            catch (System.Exception)
            {
            }
            return false;
        }

        public BlockReference BlockReference
        {
            get
            {
                return base.Entity as BlockReference;
            }
        }

        Point3d position;

    }

    //New class derrived from the DrawJig class
    // Old one left in for existing programs that use it
    public class JigEnt : DrawJig
    {
        #region private member fields

        private Point3d previousCursorPosition;

        private Point3d currentCursorPosition;

        private Entity entityToDrag;

        #endregion

        Autodesk.AutoCAD.ApplicationServices.Document doc =
            acadApp.DocumentManager.MdiActiveDocument;

        public void DragEnt(
            ObjectId EntObjectID,
            Point3d StartPoint)
        {
            try
            {
                using (Transaction trans = doc.TransactionManager.StartTransaction())
                {
                    // Assign the Entity based on the ObjectID provided
                    entityToDrag = (Entity)trans.GetObject(EntObjectID, OpenMode.ForWrite);
                    //Initialize cursor position
                    
                    previousCursorPosition = StartPoint;
                    doc.Editor.WriteMessage("\nPick a point");
                    
                    doc.Editor.Drag(this);

                    trans.Commit();
                }
            }
            catch (System.Exception)
            {

            }

        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            //Get the current cursor position
            PromptPointResult userFeedback = prompts.AcquirePoint();
            currentCursorPosition = userFeedback.Value;

            if (CursorHasMoved())
            {
                //Get the vector of the move
                Vector3d displacementVector = previousCursorPosition.GetVectorTo(currentCursorPosition);

                //Transform the ent to the new location
                entityToDrag.TransformBy(Matrix3d.Displacement(displacementVector));

                //Save the cursor position
                previousCursorPosition = currentCursorPosition;

                return SamplerStatus.OK;
            }
            else
            {
                return SamplerStatus.NoChange;
            }
        }

        protected override bool WorldDraw(WorldDraw draw)
        {
            draw.Geometry.Draw(entityToDrag);
            return true;
        }

        private bool CursorHasMoved()
        {
            return (currentCursorPosition != previousCursorPosition);
        }

    }

    // Use this one - The DragEnt function returns a boolean indicating success
    public class DragEntJig : DrawJig
    {
        #region private member fields

        private Point3d previousCursorPosition;

        private Point3d currentCursorPosition;

        private Entity entityToDrag;

        private bool success = true;

        #endregion

        Autodesk.AutoCAD.ApplicationServices.Document doc =
            acadApp.DocumentManager.MdiActiveDocument;

        public bool DragEnt(
            ObjectId EntObjectID,
            Point3d StartPoint)
        {
            
            try
            {
                using (Transaction trans = doc.TransactionManager.StartTransaction())
                {
                    // Assign the Entity based on the ObjectID provided
                    entityToDrag = (Entity)trans.GetObject(EntObjectID, OpenMode.ForWrite);
                    //Initialize cursor position

                    previousCursorPosition = StartPoint;
                    doc.Editor.WriteMessage("\nPick a point");

                    doc.Editor.Drag(this);

                    trans.Commit();
                }
            }
            catch (System.Exception)
            {
                return false;
            }
            return success;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            //Get the current cursor position
            PromptPointResult userFeedback = prompts.AcquirePoint();
            currentCursorPosition = userFeedback.Value;
            if (userFeedback.Status == PromptStatus.Cancel || userFeedback.Status == PromptStatus.Error)
            {
                success = false;
                return SamplerStatus.Cancel;
            }
            else
            {
                if (CursorHasMoved())
                {

                    //Get the vector of the move
                    Vector3d displacementVector = previousCursorPosition.GetVectorTo(currentCursorPosition);

                    //Transform the ent to the new location
                    entityToDrag.TransformBy(Matrix3d.Displacement(displacementVector));

                    //Save the cursor position
                    previousCursorPosition = currentCursorPosition;

                    return SamplerStatus.OK;
                }
                else
                {
                    return SamplerStatus.NoChange;
                }
            }
        }

        protected override bool WorldDraw(WorldDraw draw)
        {
            draw.Geometry.Draw(entityToDrag);
            return true;
        }

        private bool CursorHasMoved()
        {
            return (currentCursorPosition != previousCursorPosition);
        }

    }
}
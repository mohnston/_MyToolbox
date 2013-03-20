using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using acColors = Autodesk.AutoCAD.Colors;
//using Autodesk.AutoCAD.Customization;

namespace CADHelp
{
	public class CADHelper
	{
		public CADHelper()
		{
			Initialize();
		}

		Database db = HostApplicationServices.WorkingDatabase;
		Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
		Document doc = acadApp.DocumentManager.MdiActiveDocument;
		ObjectId ModelSpaceID = SymbolUtilityServices.GetBlockModelSpaceId(HostApplicationServices.WorkingDatabase);
		//static string xdAppName = "NSSM";
		//static string xdPartName = "PartName";
		//static string xdBlockType = "BlockType";
		//static string xdBlockTypeTitle = "Title";
		//static string xdBlockTypePanel = "Panel";

        #region Color constants
        public static acColors.Color clrMagenta = acColors.Color.FromColorIndex(acColors.ColorMethod.ByLayer, COLOR_WHITE); // acColors.Color.FromColor(System.Drawing.Color.Magenta);
        public static acColors.Color clrGreen = acColors.Color.FromColorIndex(acColors.ColorMethod.ByLayer, COLOR_GREEN);
        public static acColors.Color clrCyan = acColors.Color.FromColorIndex(acColors.ColorMethod.ByLayer, COLOR_CYAN);
        public static acColors.Color clrYellow = acColors.Color.FromColorIndex(acColors.ColorMethod.ByLayer, COLOR_YELLOW);
        public static acColors.Color clrWhite = acColors.Color.FromColorIndex(acColors.ColorMethod.ByLayer, COLOR_WHITE);
        public static acColors.Color clrBlue = acColors.Color.FromColorIndex(acColors.ColorMethod.ByLayer, COLOR_BLUE);
        public static acColors.Color clrRed = acColors.Color.FromColorIndex(acColors.ColorMethod.ByLayer, COLOR_RED);

        private static short COLOR_MAGENTA = 6;
        private static short COLOR_GREEN = 3;
        private static short COLOR_CYAN = 4;
        private static short COLOR_YELLOW = 2;
        private static short COLOR_WHITE = 7;
        private static short COLOR_BLUE = 5;
        private static short COLOR_RED = 1; 
        #endregion

        #region Layer constants
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
        #endregion

        #region Linetype
        const string LINETYPE_NAME_CONTINUOUS = "CONTINUOUS";
        const string LINETYPE_NAME_HIDDEN = "HIDDEN";
        const string LINETYPE_NAME_LOWSWING = "DASHDOT2";
        const string LINETYPE_NAME_TALLSWING = "DASHDOT"; 
        #endregion

		public const string SETTINGS_BLOCK_PATH = @"\\cad_svr\catalog\CABWEST\W_R2K_Set.dwg";

        #region Text style
        public ObjectId TextStyleBoldID = ObjectId.Null;
        public ObjectId TextStyleWingDingID = ObjectId.Null;
        public ObjectId TextStyleStdID = HostApplicationServices.WorkingDatabase.Textstyle;
        public static string TextStyleStandard = "WM_Standard";
        public static string TextStyleBold = "WM_Bold";
        public static string TextStyleWingDing = "WM_Symbol";

        public static string FontNameTrebuchetMS = "";

        int PitchNFamilyTreb = 34;
        int PitchNFamilyOptima = 2;
        int PitchNFamilyWingDing = 32; // ???

        private const string wmStd = "WM_Standard"; 
        #endregion

		private void Initialize()
		{
			
		}

		private const string alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

		[DllImport("acCore.dll",
			CallingConvention = CallingConvention.Cdecl,
			EntryPoint = "acedTrans")]

		static extern int acedTrans(
			double[] point,
			IntPtr fromRb,
			IntPtr toRb,
			int disp,
			double[] result);

		public static void AddSettingsBlock()
		{
			string setBlk = SETTINGS_BLOCK_PATH;
			string bName = SymbolUtilityServices.GetBlockNameFromInsertPathName(setBlk);
			try
			{
				CADHelper ch = new CADHelper();
				if (!ch.BlockDefExists(bName))
				{
					ch.InsertBlockDef(setBlk);
				}
				else
				{
					// REDefine block
					ch.DeleteBlockByName(bName, true);
					ch.InsertBlockDef(setBlk);
				}
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show("AddSettingsBlock error", ex.Message);
			}
		}

		public static Point3d UCSPointToDCS(Point3d pointUCS)
		{
			Point3d res = new Point3d();
			// Transform from UCS to DCS
			ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)),
				rbTo = new ResultBuffer(new TypedValue(5003, 2));

			// double[] firres = new double[] { 0, 0, 0 };
			// double[] secres = new double[] { 0, 0, 0 };
			double[] dres = new double[] { 0, 0, 0 };
			
			// Transform the point...
			acedTrans(pointUCS.ToArray(),
				rbFrom.UnmanagedObject,
				rbTo.UnmanagedObject,
				0,
				dres);
			res = new Point3d(dres[0], dres[1], 0);

			//acedTrans(
			//    start.ToArray(),
			//    rbFrom.UnmanagedObject,
			//    rbTo.UnmanagedObject,
			//    0,
			//    firres);

			//// ... and the second
			//acedTrans(
			//    end.ToArray(),
			//    rbFrom.UnmanagedObject,
			//    rbTo.UnmanagedObject,
			//    0,
			//    secres);
			return res;
		}

		public static System.Drawing.Bitmap Thumbnail()
		{
			System.Drawing.Bitmap bmp = HostApplicationServices.WorkingDatabase.ThumbnailBitmap;
			System.Drawing.Bitmap cpy = bmp.Clone() as System.Drawing.Bitmap;
			if (cpy != null)
				return cpy;
			else
				return bmp;
		}
		public static System.Drawing.Bitmap Thumbnail(string FullFilePath)
		{
			System.Drawing.Bitmap cpy = null;
			Database db = new Database();
			try
			{
				db.ReadDwgFile(FullFilePath, FileOpenMode.OpenTryForReadShare, false, string.Empty);
				System.Drawing.Bitmap bmp = db.ThumbnailBitmap;
				cpy = bmp.Clone() as System.Drawing.Bitmap;
			}
			catch
			{

			}
			return cpy;
		}

		public void SetWestmarkLayers()
		{
			// Cabinet, Border, Green, Text, Hidden, Cyan, LowSwing, TallSwing

			SetupLayer("Cabinet", LINETYPE_NAME_CONTINUOUS, clrYellow, true);
			SetupLayer("Border", LINETYPE_NAME_CONTINUOUS, clrMagenta, true);
			SetupLayer("Green", LINETYPE_NAME_CONTINUOUS, clrGreen, true);
			SetupLayer("Text", LINETYPE_NAME_CONTINUOUS, clrWhite, true);
			SetupLayer("Hidden", LINETYPE_NAME_HIDDEN, clrBlue, true);
			SetupLayer("Cyan", LINETYPE_NAME_CONTINUOUS, clrCyan, true);
			SetupLayer("LowSwing", LINETYPE_NAME_LOWSWING, clrWhite, true);
			SetupLayer("TallSwing", LINETYPE_NAME_TALLSWING, clrWhite, true);
		}

		public void SetWestmark3DLayers()
		{
			// 
			SetupLayer(layerName_2D_Cabinet, LINETYPE_NAME_CONTINUOUS, clrYellow, true);
			SetupLayer(layerName_2D_Border, LINETYPE_NAME_CONTINUOUS, clrMagenta, true);
			SetupLayer(layerName_2D_Hidden, LINETYPE_NAME_HIDDEN, clrBlue, true);
			SetupLayer(layerName_2D_Green, LINETYPE_NAME_CONTINUOUS, clrGreen, true);
			SetupLayer(layerName_2D_Text, LINETYPE_NAME_CONTINUOUS, clrWhite, true);
			SetupLayer(layerName_2D_Cyan, LINETYPE_NAME_CONTINUOUS, clrCyan, true);


			SetupLayer(layerName_3D_Cabinet, LINETYPE_NAME_CONTINUOUS, clrYellow, true);
			SetupLayer(layerName_3D_Cabinet_Base, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.DarkGoldenrod), true);
			SetupLayer(layerName_3D_Cabinet_Upper, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Peru), true);
			SetupLayer(layerName_3D_Cabinet_Panel, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.MediumAquamarine), true);

			SetupLayer(layerName_3D_Countertop, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.CornflowerBlue), true);
			SetupLayer(layerName_3D_Detail_Top_Lower, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.MediumBlue), true);
			SetupLayer(layerName_3D_Detail_Top_Upper, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Olive), true);

			SetupLayer(layerName_3D_Toekick, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Black), true);
			SetupLayer(layerName_3D_Wall, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Ivory), true);
			SetupLayer(layerName_3D_Floor, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.LightGray), true);

			SetupLayer(layerName_3D_Detail_TransactionTop, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.DodgerBlue), true);
			SetupLayer(layerName_3D_Detail_WorkTop, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.SteelBlue), true);
			SetupLayer(layerName_3D_Detail_Panel, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.SeaGreen), true);

			SetupLayer(layerName_3D_Panel, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.DarkSeaGreen), true);
			SetupLayer(layerName_3D_Panel_Removable, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.DarkSeaGreen), true);

			SetupLayer(layerName_3D_Framework, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Tan), true);
			SetupLayer(layerName_3D_Framework_Studs, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Brown), true);
			SetupLayer(layerName_3D_Framework_Plate, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.BurlyWood), true);

			SetupLayer(layerName_3D_Face, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.SteelBlue), true);
			SetupLayer(layerName_3D_Face_Panel, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.SteelBlue), true);
			SetupLayer(layerName_3D_Face_Reveal, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.MidnightBlue), true);

			SetupLayer(layerName_3D_Hardware, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Gray), true);
			SetupLayer(layerName_3D_Hardware_Bracket, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.Silver), true);
			SetupLayer(layerName_3D_Hardware_Standoff, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.SlateGray), true);
			SetupLayer(layerName_3D_Hardware_Fastener, LINETYPE_NAME_CONTINUOUS, acColors.Color.FromColor(System.Drawing.Color.LightGray), true);



		}

		//public void SetupLayer(string LayerName, string LineTypeName, Autodesk.AutoCAD.Colors.Color acColor, bool OverrideIfExists)
		//{
		//    // Autodesk.AutoCAD.ApplicationServices.Document thisDrawing = acadApp.Application.DocumentManager.MdiActiveDocument;
		//    // Database docDB = doc.Database;
		//    using (Transaction trans = db.TransactionManager.StartTransaction())
		//    {
		//        LayerTable ltab = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
		//        LayerTableRecord ltr;
		//        try
		//        {
		//            ObjectId lid = GetTableRecordId(db.LayerTableId, LayerName);
		//            // ObjectId lid = ltab[LayerName]; // if layer doesn't exist this fails
		//            if (lid == ObjectId.Null)
		//            {
		//                // Layer with that name does not exist
		//                try
		//                {
		//                    ltr = new LayerTableRecord();
		//                    ltr.Name = LayerName;
		//                    ltr.Color = acColor;
		//                    // set the linetype
		//                    SetLayerLineType(ref ltr, LineTypeName.ToUpper());
		//                    ltab.Add(ltr);
		//                    trans.AddNewlyCreatedDBObject(ltr, true);
		//                    trans.Commit();
		//                }
		//                catch (System.Exception ex)
		//                {
		//                    System.Windows.Forms.MessageBox.Show("Error adding new layer.\n" + ex.Message);
		//                }
		//            }
		//            else
		//            {
		//                // A Layer with that name already exists
		//                ltr = (LayerTableRecord)trans.GetObject(lid, OpenMode.ForRead);
		//                if (OverrideIfExists)
		//                {
		//                    try
		//                    {
		//                        ltr.UpgradeOpen();
		//                        ltr.Color = acColor;
		//                        SetLayerLineType(ref ltr, LineTypeName.ToUpper());
		//                        trans.Commit();
		//                    }
		//                    catch (System.Exception ex)
		//                    {
		//                        System.Windows.Forms.MessageBox.Show("Error updating existing layer.\n" + ex.Message);
		//                    }
		//                }

		//            }
		//        }
		//        catch (System.Exception ex)
		//        {
		//            System.Windows.Forms.MessageBox.Show("Error setting up layer.\n" + ex.Message);
		//        }
		//    } // end of "Using"
		//}

		//private void SetLayerLineType(ref LayerTableRecord ltr, string LineTypeName)
		//{
		//    ObjectId ltID = GetTableRecordId(db.LinetypeTableId, LineTypeName);
		//    if (ltID != ObjectId.Null)
		//        ltr.LinetypeObjectId = ltID;
		//    else
		//        db.LoadLineTypeFile(LineTypeName, "acad.lin");
		//    ltID = GetTableRecordId(db.LinetypeTableId, LineTypeName);
		//    if (ltID != ObjectId.Null)
		//        ltr.LinetypeObjectId = ltID;
		//}

		public static void LayerShow(string LayerName, bool ShowLayer)
		{
			using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
			{
				LayerTable ltab = (LayerTable)tr.GetObject(HostApplicationServices.WorkingDatabase.LayerTableId, OpenMode.ForWrite);
				try
				{
					LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(ltab[LayerName], OpenMode.ForWrite);
					ltr.IsOff = !ShowLayer;
				}
				catch (System.Exception)
				{
				}
				tr.Commit();
			}
		}

		public static List<string> LockedLayerNames()
		{
			List<string> lln = new List<string>();
			using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
			{
				LayerTable ltab = (LayerTable)tr.GetObject(HostApplicationServices.WorkingDatabase.LayerTableId, OpenMode.ForRead);
				try
				{
					foreach (ObjectId ltrID in ltab)
					{
						LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(ltrID, OpenMode.ForRead);
						if (ltr.IsLocked)
							lln.Add(ltr.Name);
					}
				}
				catch (System.Exception)
				{
				}
				tr.Commit();
			}
			return lln;
		}

		public void StripBlockByName(string bName, string xdSafeKey)
		{
			
			// Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
			// Database docDB = doc.Database;
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);
				try
				{
					// Get the block definition (if doesn't exist goes to "catch")
					BlockTableRecord btr = (BlockTableRecord)bt[bName].GetObject(OpenMode.ForWrite);
					BlockTableRecordEnumerator btren = btr.GetEnumerator();
					while (btren.MoveNext())
					{
						DBObject dobj = trans.GetObject(btren.Current, OpenMode.ForWrite);
						ResultBuffer rbuf = dobj.XData;
						ResultBufferEnumerator rbufEnum = rbuf.GetEnumerator();
						bool hasSafeKey = false;
						while (rbufEnum.MoveNext())
						{
							if (rbufEnum.Current.TypeCode == (short)TypeCode.String)
							{
								if (rbufEnum.Current.Value.ToString() == xdSafeKey)
								{
									hasSafeKey = true;
								}
							}
						}
						if (!hasSafeKey) dobj.Erase();
					}
					trans.Commit();
				}
				catch (System.Exception)
				{
					// eKeyNotFound - Block definition does not exist
					// throw;
				}
			}
		}

		public void StripBlockByName(string bName)
		{
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);
				try
				{
					// Get the block definition (if doesn't exist goes to "catch")
					BlockTableRecord btr = (BlockTableRecord)bt[bName].GetObject(OpenMode.ForWrite);
					BlockTableRecordEnumerator btren = btr.GetEnumerator();
					while (btren.MoveNext())
					{
						trans.GetObject(btren.Current, OpenMode.ForWrite).Erase();
					}
					trans.Commit();
				}
				catch (System.Exception)
				{
					// eKeyNotFound - Block definition does not exist
					// throw;
				}
			}
		}

		public void DeleteBlockByName(string bName, bool EraseDefinition)
		{
			// First delete any instances of the block
			// Then the definition
			// Autodesk.AutoCAD.ApplicationServices.Document doc = acadApp.DocumentManager.MdiActiveDocument;
			// Database docDB = doc.Database;
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);
				try
				{
					// Get the block definition (if doesn't exist goes to "catch")
					BlockTableRecord btr = (BlockTableRecord)bt[bName].GetObject(OpenMode.ForWrite);
					// Get any references
					ObjectIdCollection bRefs = btr.GetBlockReferenceIds(true, true);
					foreach (ObjectId refID in bRefs)
					{
						if (refID.IsErased == false)
						{
							// erase the reference
							Entity ent = (Entity)trans.GetObject(refID, OpenMode.ForWrite);
							ent.Erase();
						}
					}
					// Erase the block definition if directed
					if (EraseDefinition) btr.Erase();
				}
				catch (System.Exception)
				{
					// eKeyNotFound - Block definition does not exist
				}
				trans.Commit();
			}
		}

		public int CountBlocksByName(string bName)
		{
			int res = 0;
			// acadApp.Document doc = acadApp.Application.DocumentManager.MdiActiveDocument;
			// Database db = doc.Database;
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
				try
				{
					// Get the block definition (if doesn't exist goes to "catch")
					BlockTableRecord btr = (BlockTableRecord)bt[bName].GetObject(OpenMode.ForRead);
					
					// Get any references
					ObjectIdCollection bRefs = btr.GetBlockReferenceIds(true,true);
					res = bRefs.Count;
					//trans.Commit();
				}
				catch (System.Exception)
				{
					// eKeyNotFound - Block definition does not exist
					// throw;
				}
			}
			return res;
		}
		
		public bool BlockDefExists(string bName)
		{
			bool res = false;
			
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
				res = bt.Has(bName);
			}
			return res;
		}

		public ObjectId GetBlockDef(string bName, bool UniformScale) // Creates a block table record - If already exists returns ObjectID
		{
			ObjectId res = ObjectId.Null;
			// acadApp.Document doc = acadApp.Application.DocumentManager.MdiActiveDocument;
			// Database docDB = doc.Database;
			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);
				try
				{
					// Get the block definition (if doesn't exist goes to "catch")
					BlockTableRecord btr = (BlockTableRecord)bt[bName].GetObject(OpenMode.ForRead);
					if (UniformScale)
					{
						btr.UpgradeOpen();
						btr.BlockScaling = BlockScaling.Uniform;
						btr.Units = UnitsValue.Inches;
					}
					res = btr.ObjectId;
				}
				catch (System.Exception)
				{
					// eKeyNotFound - Block definition does not exist
				}
				if (res == ObjectId.Null)
				{

					BlockTableRecord newbtr = new BlockTableRecord();
					newbtr.Name = bName;
					newbtr.Units = UnitsValue.Inches;
					if (UniformScale)
						newbtr.BlockScaling = BlockScaling.Uniform;
					else
						newbtr.BlockScaling = BlockScaling.Any;
					try
					{
						res = bt.Add(newbtr);
						trans.AddNewlyCreatedDBObject(newbtr, true);
						//trans.Commit();
					}
					catch (System.Exception)
					{
						if (bt[bName].IsErased)
						{
							System.Windows.Forms.MessageBox.Show("The block table record " + bName + " has been erased.", "Block erased");
							res = ObjectId.Null;
						}
					}
				}
				trans.Commit();
			}
			return res;

		}

		public ObjectId GetBlockDef(string bName) // Creates a block table record - If already exists returns ObjectID
		{
			ObjectId res = ObjectId.Null;
			try
			{
				res = GetBlockDef(bName, false);
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show("Exception in GetBlockDef" + Environment.NewLine + ex.Message);
				
			}
			return res;
		}

		public ObjectId InsertBlock(string FullBlockPath)
		{
			ObjectId resid = ObjectId.Null;
			try
			{
				ObjectId bdid = InsertBlockDef(FullBlockPath);
				if (bdid.IsValid)
				{
					resid = InsertBlockRef(bdid);
				}
				//else
				//{
				//    System.Windows.Forms.MessageBox.Show("InsertBlock broke");
				//}
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message);
				throw;
			} 
			return resid;
		}

		public ObjectId InsertBlock(string BlockNameOrPath, Point3d InsertionPoint)
		{
			string blockName = BlockNameOrPath;
			//string res = string.Empty;
			ObjectId oidRes = ObjectId.Null;
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
								return oidRes; // string.Empty;
							}
						}

						BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

						BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForRead);
						
						ObjectContextManager ocm = db.ObjectContextManager;
						ObjectContextCollection occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");

						// Block needs to be inserted to current space before
						// being able to append attribute to it
						// Insert the block reference
						BlockReference br = new BlockReference(InsertionPoint, btr.ObjectId);

						space.AppendEntity(br);
						tr.AddNewlyCreatedDBObject(br, true);
						oidRes = br.ObjectId;
						if (btr.Annotative == AnnotativeStates.True)
						{

							DBObject obj = tr.GetObject(oidRes, OpenMode.ForRead);
							if (obj != null)
							{
								//obj.AddContext(occ.CurrentContext);
								ObjectContexts.AddContext(obj, occ.CurrentContext);
							}
						}

						if (btr.HasAttributeDefinitions)
						{
							foreach (ObjectId objId in btr)
							{
								AttributeDefinition AttDef = tr.GetObject(objId, OpenMode.ForRead) as AttributeDefinition;
								if (AttDef != null)
								{
									AttributeReference AttRef = new AttributeReference();
									AttRef.SetAttributeFromBlock(AttDef, br.BlockTransform);
									ObjectId attObj = br.AttributeCollection.AppendAttribute(AttRef);
									tr.AddNewlyCreatedDBObject(AttRef, true);

									if (btr.Annotative == AnnotativeStates.True)
									{
										DBObject aobj = tr.GetObject(attObj, OpenMode.ForRead);
										if (aobj != null)
										{
											// aobj.AddContext(occ.CurrentContext);
											ObjectContexts.AddContext(aobj, occ.CurrentContext);
										}
									}
								}
							}
						}

						//// Commit changes if user accepted, otherwise discard
						tr.Commit();
					}
				}
			}
			catch (Autodesk.AutoCAD.Runtime.Exception ex)
			{
				//System.Windows.Forms.MessageBox.Show(ex.Message);
			}
			finally
			{
				ed.WriteMessage("\nCommand: ");
			}
			return oidRes;

		}
		public static ObjectId InsertBlock(Database db, string loName, string BlkPath, string blkName, Point3d insPt)
		{
			ObjectId RtnObjId = ObjectId.Null;
			using (Transaction Trans = db.TransactionManager.StartTransaction())
			{
				DBDictionary LoDict = Trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
				foreach (DictionaryEntry de in LoDict)
				{
					if (string.Compare((string)de.Key, loName, true).Equals(0))
					{
						Layout Lo = Trans.GetObject((ObjectId)de.Value, OpenMode.ForWrite) as Layout;
						BlockTable BlkTbl = Trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
						BlockTableRecord LoRec = Trans.GetObject(Lo.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
						ObjectId BlkTblRecId = GetTableRecordId(BlkTbl.Id, blkName);
						if (BlkTblRecId.IsNull)
						{
							BlkPath = HostApplicationServices.Current.FindFile(BlkPath + blkName + ".dwg", db, FindFileHint.Default);
							if (string.IsNullOrEmpty(BlkPath))
								return RtnObjId;
							BlkTbl.UpgradeOpen();
							using (Database tempDb = new Database(false, true))
							{
								tempDb.ReadDwgFile(BlkPath, FileShare.Read, true, null);
								db.Insert(blkName, tempDb, false);
							}
							BlkTblRecId = GetTableRecordId(BlkTbl.Id, blkName);
						}
						LoRec.UpgradeOpen();
						BlockReference BlkRef = new BlockReference(insPt, BlkTblRecId);

						ObjectId BlkRefID = LoRec.AppendEntity(BlkRef);
						Trans.AddNewlyCreatedDBObject(BlkRef, true);

						BlockTableRecord BlkTblRec = Trans.GetObject(BlkTblRecId, OpenMode.ForWrite) as BlockTableRecord;

						ObjectContextManager ocm = db.ObjectContextManager;
						ObjectContextCollection occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");

						if (BlkTblRec.Annotative == AnnotativeStates.True)
						{

							DBObject obj = Trans.GetObject(BlkRefID, OpenMode.ForRead);
							if (obj != null)
							{
								// obj.AddContext(occ.CurrentContext);
								ObjectContexts.AddContext(obj, occ.CurrentContext);
							}
						}

						if (BlkTblRec.HasAttributeDefinitions)
						{
							foreach (ObjectId objId in BlkTblRec)
							{
								AttributeDefinition AttDef = Trans.GetObject(objId, OpenMode.ForRead) as AttributeDefinition;
								if (AttDef != null)
								{
									AttributeReference AttRef = new AttributeReference();
									AttRef.SetAttributeFromBlock(AttDef, BlkRef.BlockTransform);
									ObjectId attObj = BlkRef.AttributeCollection.AppendAttribute(AttRef);
									Trans.AddNewlyCreatedDBObject(AttRef, true);

									if (BlkTblRec.Annotative == AnnotativeStates.True)
									{
										DBObject aobj = Trans.GetObject(attObj, OpenMode.ForRead);
										if (aobj != null)
										{
											// aobj.AddContext(occ.CurrentContext);
											ObjectContexts.AddContext(aobj, occ.CurrentContext);
										}
									}
								}
							}
						}
						Trans.Commit();
						//return RtnObjId;
						return BlkRef.Id;
					}
				}
			}
			return RtnObjId;
		}

		public static object GetDynamicPropertyValue(ObjectId BlockRefOID, string PropName)
		{
			object res = null;
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			using (DocumentLock dl = doc.LockDocument())
			{
				using (Transaction tr = doc.TransactionManager.StartTransaction())
				{
					try
					{
						BlockReference br = tr.GetObject(BlockRefOID, OpenMode.ForRead) as BlockReference;
						if (br.IsDynamicBlock)
						{
							DynamicBlockReferencePropertyCollection pc1 =
								br.DynamicBlockReferencePropertyCollection;
							for (int i = 0; i < pc1.Count; i++)
							{
								DynamicBlockReferenceProperty dprop = pc1[i];
								// does propName match prop from incoming
								string propName = dprop.PropertyName;
								if (PropName.Equals(propName))
								{
									res = dprop.Value;
									break;
								}
							}
						}

					}
					catch (Autodesk.AutoCAD.Runtime.Exception ex)
					{
						System.Windows.Forms.MessageBox.Show(ex.Message);
					}
				}
			}
			return res;
		}
		
		static public string EffectiveName(BlockReference blkref)
		{
			if (blkref.IsDynamicBlock)
			{
				using (BlockTableRecord obj = (BlockTableRecord)
				blkref.DynamicBlockTableRecord.GetObject(OpenMode.ForRead))
					return obj.Name;
			}
			return blkref.Name;
		}
		
		public static bool AttSync(ObjectId BlockRefOID)
		{
			bool res = false;
			Database db = HostApplicationServices.WorkingDatabase;
			Document doc = acadApp.DocumentManager.MdiActiveDocument;

			try
			{
				using (DocumentLock dlock = doc.LockDocument())
				using (Transaction tr = db.TransactionManager.StartTransaction())
				{
					BlockReference bref = tr.GetObject(BlockRefOID, OpenMode.ForWrite) as BlockReference;
					
					if (bref != null)
					{
						Database wdb = HostApplicationServices.WorkingDatabase;
						HostApplicationServices.WorkingDatabase = db;
						AttributeCollection attCol = bref.AttributeCollection;
						for (int i = 0; i < attCol.Count; i++)
						{
							AttributeReference attRef = tr.GetObject(attCol[i], OpenMode.ForWrite) as AttributeReference;
							
							attRef.AdjustAlignment(BlockRefOID.Database);
						}
						HostApplicationServices.WorkingDatabase = wdb;
						res = true;
					}
					tr.Commit();
				}
			}
			catch (Autodesk.AutoCAD.Runtime.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message);
			}
			return res;
		}

		public static bool SetDynamicProperty(ObjectId BlockRefOID, string PropName, object value)
		{
			bool res = false;
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			string info = string.Empty;

			using (DocumentLock dl = doc.LockDocument())
			{
				using (Transaction tr = doc.TransactionManager.StartTransaction())
				{
					try
					{
						BlockReference br = tr.GetObject(BlockRefOID, OpenMode.ForRead) as BlockReference;
						if (br.IsDynamicBlock)
						{
							DynamicBlockReferencePropertyCollection pc1 =
								br.DynamicBlockReferencePropertyCollection;
							for (int i = 0; i < pc1.Count; i++)
							{
								DynamicBlockReferenceProperty dprop = pc1[i];
								// does propName match prop from incoming hashtable
								//info = info + "Dynamic Property" + Environment.NewLine +
								//    "PropertyName = " + dprop.PropertyName + Environment.NewLine +
								//    "PropertyTypeCode = " + dprop.PropertyTypeCode.ToString() + Environment.NewLine +
								//    "Allowed values = " + dprop.GetAllowedValues().ToString() + Environment.NewLine +
								//    "Units type = " + dprop.UnitsType.ToString() + Environment.NewLine +
								//    "--------------------------------" + Environment.NewLine;
								string propName = dprop.PropertyName;
								
								if (PropName.Equals(propName))
								{
									try
									{
										dprop.Value = Convert.ToDouble(value);
										res = true;
									}
									catch
									{
										try
										{
											dprop.Value = value.ToString();
											res = true;
										}
										catch
										{

										}
									}
									break;
								}
							}
							tr.Commit();
						}
						// Console.WriteLine(info);
					}
					catch (Autodesk.AutoCAD.Runtime.Exception ex)
					{
						res = false;
					}
				}
			}
			// System.Windows.Forms.MessageBox.Show(info);
			return res;
		}

		public static bool SetDynamicProperties(ObjectId BlockRefOID, Hashtable PropValue)
		{
			bool res = false;
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			using (DocumentLock dl = doc.LockDocument())
			{
				using (Transaction tr = doc.TransactionManager.StartTransaction())
				{
					BlockReference br = tr.GetObject(BlockRefOID, OpenMode.ForRead) as BlockReference;
					if (br.IsDynamicBlock)
					{
						DynamicBlockReferencePropertyCollection pc1 =
							br.DynamicBlockReferencePropertyCollection;
						for (int i = 0; i < pc1.Count; i++)
						{
							DynamicBlockReferenceProperty dprop = pc1[i];
							// does propName match prop from incoming hashtable
							string propName = dprop.PropertyName;
							// System.Windows.Forms.MessageBox.Show(propName);
							foreach (DictionaryEntry de in PropValue)
							{
								if (de.Key.ToString().Equals(propName))
								{
									if (de.Value.ToString().Trim().Length > 0)
									{
										try
										{
											dprop.Value = Convert.ToDouble(de.Value);
										}
										catch
										{
											try
											{
												dprop.Value = de.Value.ToString();
											}
											catch
											{

											}
										}
									}
									break;
								}
							}
						}
						tr.Commit();
					}
				}
			}
			return res;
		}

		//public ObjectId InsertBlock(
		//    string FullBlockPath,
		//    Point3d InsertionPoint)
		//{
		//    ObjectId resid = ObjectId.Null;
		//    ObjectId bdid = InsertBlockDef(FullBlockPath);

		//    resid = InsertBlockRef(bdid, InsertionPoint);
		//    return resid;
		//}

		private ObjectId InsertBlockRef(
			ObjectId BlockDefObjectID)
		{
			// Inserts at 0,0,0
			Point3d insPnt = new Point3d(0, 0, 0);
			return InsertBlockRef(BlockDefObjectID, insPnt);
		}
		
		public ObjectId InsertBlockDef(string FullBlockPath)
		{
			ObjectId resID = ObjectId.Null;
			if (FullBlockPath.Length < 8) return resID;
			FullBlockPath = FullBlockPath.ToUpper();
			if (File.Exists(FullBlockPath) == false) return resID;
			if (FullBlockPath.EndsWith(".DWG") == false) return resID;
			
			string bName = SymbolUtilityServices.GetBlockNameFromInsertPathName(FullBlockPath);
			if (bName.Length > 0)
			{
				try
				{
					if (BlockDefExists(bName)) // check
					{
						resID = GetBlockDef(bName); // check
					}
					else
					{
						using (Transaction tr = doc.TransactionManager.StartTransaction())
						{
							try
							{
								DocumentLock dlock = doc.LockDocument();
								Database newdb = new Database(false, false);
								newdb.ReadDwgFile(FullBlockPath, System.IO.FileShare.Read, true, "");
								ObjectId BlkId = doc.Database.Insert(bName, newdb, false);
								newdb.CloseInput(true);
								newdb.Dispose();
								dlock.Dispose();
								resID = BlkId;
							}
							catch (Autodesk.AutoCAD.Runtime.Exception aex)
							{
								System.Windows.Forms.MessageBox.Show("AutoCAD runtime exception\n" + aex.Message);
							}
							catch (System.Exception ex)
							{
								System.Windows.Forms.MessageBox.Show("System exception\n" + ex.Message);
							}
							finally
							{
								tr.Commit();
							}
						}
					}

				}
				catch (Autodesk.AutoCAD.Runtime.Exception aex2)
				{
					System.Windows.Forms.MessageBox.Show("AutoCAD runtime exception\n" + aex2.Message);
				}
				catch (System.Exception ex2)
				{
					System.Windows.Forms.MessageBox.Show("System exception\n" + ex2.Message);
				}
			}
			return resID;
		}

		public ObjectId InsertBlockRef(
			ObjectId BlockDefObjectID, 
			Point3d InsertionPoint)
		{
			ObjectId resID = ObjectId.Null;
			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction tr = db.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead, true);
				BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
				BlockReference bref = new BlockReference(InsertionPoint, BlockDefObjectID);
				bref.RecordGraphicsModified(true);
				resID = btr.AppendEntity(bref);
				BlockTableRecord dbtr = (BlockTableRecord)tr.GetObject(BlockDefObjectID, OpenMode.ForWrite);
				if (dbtr.HasAttributeDefinitions)
				{
					foreach (ObjectId boid in dbtr)
					{
						Entity ent = (Entity)tr.GetObject(boid, OpenMode.ForWrite);
						if (ent is AttributeDefinition)
						{
							AttributeDefinition adef = (AttributeDefinition)ent;
							AttributeReference aref = new AttributeReference();

							aref.SetAttributeFromBlock(adef, bref.BlockTransform);
							bref.AttributeCollection.AppendAttribute(aref);
							tr.AddNewlyCreatedDBObject(aref, true);
						}
					}
					
				}
				tr.AddNewlyCreatedDBObject(bref, true);
				// bref.ExplodeToOwnerSpace();
				// bref.Erase();
				tr.Commit();
			}
			return resID;
		}

		public void SetupLayer(string LayerName, string LineTypeName, Autodesk.AutoCAD.Colors.Color acColor, bool OverrideIfExists)
		{
			// Autodesk.AutoCAD.ApplicationServices.Document thisDrawing = acadApp.Application.DocumentManager.MdiActiveDocument;
			// Database docDB = doc.Database;
			using (Transaction trans = db.TransactionManager.StartTransaction())
			{
				LayerTable ltab = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
				LayerTableRecord ltr;
				try
				{
					ObjectId lid = GetTableRecordId(db.LayerTableId, LayerName);
					// ObjectId lid = ltab[LayerName]; // if layer doesn't exist this fails
					if (lid == ObjectId.Null)
					{
						// Layer with that name does not exist
						try
						{
							ltr = new LayerTableRecord();
							ltr.Name = LayerName;
							ltr.Color = acColor;
							// set the linetype
							SetLayerLineType(ref ltr, LineTypeName.ToUpper());
							ltab.Add(ltr);
							trans.AddNewlyCreatedDBObject(ltr, true);
							trans.Commit();
						}
						catch (System.Exception ex)
						{
							System.Windows.Forms.MessageBox.Show("Error adding new layer.\n" + ex.Message);
						}
					}
					else
					{
						// A Layer with that name already exists
						ltr = (LayerTableRecord)trans.GetObject(lid, OpenMode.ForRead);
						if (OverrideIfExists)
						{
							try
							{
								ltr.UpgradeOpen();
								ltr.Color = acColor;
								SetLayerLineType(ref ltr, LineTypeName.ToUpper());
								trans.Commit();
							}
							catch (System.Exception ex)
							{
								System.Windows.Forms.MessageBox.Show("Error updating existing layer.\n" + ex.Message);
							}
						}

					}
				}
				catch (System.Exception ex)
				{
					System.Windows.Forms.MessageBox.Show("Error setting up layer.\n" + ex.Message);
				}
			} // end of "Using"
		}

		private void SetLayerLineType(ref LayerTableRecord ltr, string LineTypeName)
		{
			ObjectId ltID = GetTableRecordId(db.LinetypeTableId, LineTypeName);
			if (ltID != ObjectId.Null)
				ltr.LinetypeObjectId = ltID;
			else
				db.LoadLineTypeFile(LineTypeName, "acad.lin");
			ltID = GetTableRecordId(db.LinetypeTableId, LineTypeName);
			if (ltID != ObjectId.Null)
				ltr.LinetypeObjectId = ltID;
		}

		public ObjectId GetLineTypeIDByName(string LineTypeName)
		{
			// Autodesk.AutoCAD.ApplicationServices.Document thisDrawing = acadApp.Application.DocumentManager.MdiActiveDocument;
			// Database docDB = doc.Database;
			ObjectId resID = ObjectId.Null;
			resID = GetTableRecordId(db.LinetypeTableId, LineTypeName.ToUpper());
			return resID;
		}

		public ObjectId GetTextStyleIDByName(string TextStyleName)
		{
			ObjectId res = ObjectId.Null;
			// Autodesk.AutoCAD.ApplicationServices.Document thisDrawing = acadApp.Application.DocumentManager.MdiActiveDocument;
			// Database docDB = doc.Database;
			using (Transaction trans = db.TransactionManager.StartTransaction())
			{
				TextStyleTable tst = (TextStyleTable)trans.GetObject(db.TextStyleTableId, OpenMode.ForRead);
				try
				{
					res = tst[TextStyleName];
				}
				catch (System.Exception ex)
				{
					// didn't find a text style with that name
				}
			}
			return res;
		}

		public static TextStyleTableRecord GetTextStyleTable(string StyleName)
		{
			TextStyleTableRecord resTbl = null;
			Database db = HostApplicationServices.WorkingDatabase;
			using (Transaction trans = db.TransactionManager.StartTransaction())
			{
				TextStyleTable tst = (TextStyleTable)trans.GetObject(db.TextStyleTableId, OpenMode.ForRead);
				try
				{
					if (tst.Has(StyleName))
					{
						TextStyleTableRecord tstr = trans.GetObject(tst[StyleName], OpenMode.ForRead) as TextStyleTableRecord;
						resTbl = tstr;
					}
					else
					{
						TextStyleTableRecord tstr = trans.GetObject(tst["Standard"], OpenMode.ForRead) as TextStyleTableRecord;
						resTbl = tstr;
					}
				}
				catch (System.Exception ex)
				{
					TextStyleTableRecord tsr = trans.GetObject(tst["Standard"], OpenMode.ForRead) as TextStyleTableRecord;
					resTbl = tsr;
					// didn't find a text style with that name
				}
				trans.Commit();
			}
			return resTbl;
		}

		private ObjectId CreateTextStyle(string StyleName, string FontName, int PitchNFamily)
		{
			// Autodesk.AutoCAD.ApplicationServices.Document thisDrawing = acadApp.Application.DocumentManager.MdiActiveDocument;
			// Database docDB = doc.Database;
			
			ObjectId resID = ObjectId.Null;
			using (Transaction trans = db.TransactionManager.StartTransaction())
			{
				TextStyleTable tst = (TextStyleTable)trans.GetObject(db.TextStyleTableId, OpenMode.ForRead);
				foreach (ObjectId tsID in tst)
				{
					TextStyleTableRecord tstr = (TextStyleTableRecord)trans.GetObject(tsID, OpenMode.ForRead);
					if (tstr.Name.ToUpper() == StyleName.ToUpper()) resID = tstr.ObjectId;
				}
				if (resID == ObjectId.Null)
				{
					tst.UpgradeOpen();
					TextStyleTableRecord newTSTR = new TextStyleTableRecord();
					newTSTR.Font = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(FontName, false, false, 0, PitchNFamily);
					newTSTR.TextSize = 0;
					newTSTR.Name = StyleName;
					tst.Add(newTSTR);
					trans.AddNewlyCreatedDBObject(newTSTR, true);
					resID = newTSTR.ObjectId;
					trans.Commit();
				}
			}
			return resID;
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
				catch (Autodesk.AutoCAD.Runtime.Exception ex)
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
				catch (Autodesk.AutoCAD.Runtime.Exception ex)
				{
					// res remains empty string
				}
			}

			if (fullPath.Length > 0)
			{

				blkName = SymbolUtilityServices.GetBlockNameFromInsertPathName(fullPath);

				// Try opening the drawing file
				// It should fail if the file is bad
				Database bdb = new Database(false, false);
				bdb.ReadDwgFile(fullPath, System.IO.FileShare.Read, true, string.Empty);
				db.Insert(blkName, bdb, true);
				bdb.CloseInput(true);
				bdb.Dispose();
				res = blkName;
			}

			return res;
		}

		private ObjectId GetBlockID(string BlockName)
		{
			// Use this instead of "Has" because the erased objects may give a false positive
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


		/// <summary>
		/// Consumption example:
		///
		///   Remember that the generic argument determines
		///   what kind of objects may exist in the set.
		///   Attempting to add inelegible types will not
		///   result in an error, and they will not be
		///   added to the collection.
		///
		///   This code takes a user selection, and then
		///   modifies the color of all polylnes in the
		///   selection (ignoring other type of objects).
		/// </summary>
		/// 

		public static ObjectId ObjectIdFromHandle(Database db, string strHandle)
		{
			ObjectId resID = ObjectId.Null;
			if (strHandle.Length == 0)
				return resID;
			Int32 nHandle = Int32.Parse(strHandle, NumberStyles.AllowHexSpecifier);
			Handle handle = new Handle(nHandle);
			return db.GetObjectId(false, handle, 0);
		}
		public static ObjectId ObjectIdFromHandle(string strHandle)
		{
			Int32 nHandle = Int32.Parse(strHandle, NumberStyles.AllowHexSpecifier);
			Handle handle = new Handle(nHandle);
			return HostApplicationServices.WorkingDatabase.GetObjectId(false, handle, 0);
		}

		public static void SetAttributeValue(
			ObjectId BlockRefID, 
			string AttTagName, 
			string AttValue)
		{
			AttTagName = AttTagName.ToUpper();
			Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
			DocumentLock dl = doc.LockDocument();
			
			using (Transaction tr = doc.TransactionManager.StartTransaction())
			{
				BlockReference bref = (BlockReference)tr.GetObject(BlockRefID, OpenMode.ForRead);
				bref.RecordGraphicsModified(true);
				foreach (ObjectId attid in bref.AttributeCollection)
				{
					AttributeReference attref = (AttributeReference)tr.GetObject(attid, OpenMode.ForRead);
					if (attref.Tag == AttTagName)
					{
						bref.UpgradeOpen();
						attref.UpgradeOpen();
						Database wdb = HostApplicationServices.WorkingDatabase;
						try
						{
							
							Database adb = attref.Database;
							attref.TextString = AttValue;
							attref.AdjustAlignment(adb);
						}
						finally
						{
							HostApplicationServices.WorkingDatabase = wdb;
						}                        
						// attref.Draw();
						break;
					}
					attref.Dispose();
				}
				tr.Commit();
			}
			dl.Dispose();

		}

		public static void SetAttributeValues(
			ObjectId BlockRefID,
			Hashtable htAttsNVals)
		{
			// DNA.Tools.UpdateAttributesInBlock(BlockRefID, htAttsNVals);
			ArrayList keys = new ArrayList(htAttsNVals.Keys);
			ArrayList vals = new ArrayList(htAttsNVals.Values);
			for (int i = 0; i < htAttsNVals.Count; i++)
			{
				SetAttributeValue(BlockRefID, keys[i].ToString(), vals[i].ToString());
			}
		}



		public static string GetAttributeValue(
			ObjectId BlockRefID, 
			string AttTagName)
		{
			string res = string.Empty;
			AttTagName = AttTagName.ToUpper();
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction tr = doc.TransactionManager.StartTransaction())
			{
				try
				{
					BlockReference bref = (BlockReference)tr.GetObject(BlockRefID, OpenMode.ForWrite);
					foreach (ObjectId attid in bref.AttributeCollection)
					{
						AttributeReference attref = (AttributeReference)tr.GetObject(attid, OpenMode.ForWrite);
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
					System.Windows.Forms.MessageBox.Show(ex.Message);
				} 
				tr.Commit();
			}
			return res;
		}

		public static bool MakeDwgActive(
			string drawingName,
			bool readOnly)
		{
			// drawingName may be full path or just file name
			bool res = false;
			drawingName = drawingName.ToUpper();
			if (drawingName == acadApp.DocumentManager.MdiActiveDocument.Name.ToUpper())
			{
				res = true;
			}
			else
			{
				object currDoc = null;
				string shortName = drawingName.Substring(drawingName.LastIndexOf(@"\") + 1);
				foreach (Document doc in acadApp.DocumentManager)
				{
					if (doc.Name.ToUpper().EndsWith(shortName))
					{
						currDoc = doc;
					}
				}
				// If drawing is open but NOT the active drawing
				// this will make it the active document
				if (currDoc != null)
				{
					if (acadApp.DocumentManager.MdiActiveDocument.Name.ToUpper().EndsWith(shortName) == false)
					{
						acadApp.DocumentManager.MdiActiveDocument = (Document)currDoc;
					}
					res = true;
				}
				// If drawing was NOT already open then 
				// open it and it will automatically become the active document
				else
				{
					try
					{
						acadApp.DocumentManager.Open(drawingName, readOnly);
						res = true;
					}
					catch
					{
					}
				}
			}
			return res;
		}

		public static bool DwgIsOpen(string drawingName)
		{
			bool res = false;
			if (drawingName.Contains(@"\"))
			{
				string shortName = drawingName.Substring(drawingName.LastIndexOf(@"\") + 1);
				shortName = shortName.ToUpper();
				foreach (Document doc in acadApp.DocumentManager)
				{
					if (doc.Name.ToUpper().EndsWith(shortName))
					{
						res = true;
					}
				}
			}
			else
			{
				drawingName = drawingName.ToUpper();
				foreach (Document doc in acadApp.DocumentManager)
				{
					if (doc.Name.ToUpper().EndsWith(drawingName))
					{
						res = true;
					}
				}
			}
			return res;
		}

		public static ObjectIdCollection GetDynamicBlocksByName(string BlkName)
		{
			ObjectIdCollection res = new ObjectIdCollection();
			Database db = Application.DocumentManager.MdiActiveDocument.Database;
			using (Transaction trans = db.TransactionManager.StartTransaction())
			{
				//get the blockTable and iterate through all block Defs
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
				foreach (ObjectId btrId in bt)
				{
					//get the block Def and see if it is anonymous
					BlockTableRecord btr = (BlockTableRecord)trans.GetObject(btrId, OpenMode.ForRead);
					if (btr.IsDynamicBlock && btr.Name.ToUpper().Equals(BlkName.ToUpper()))
					{
						//get all anonymous blocks from this dynamic block
						ObjectIdCollection anonymousIds = btr.GetAnonymousBlockIds();
						ObjectIdCollection dynBlockRefs = new ObjectIdCollection();
						foreach (ObjectId anonymousBtrId in anonymousIds)
						{
							//get the anonymous block
							BlockTableRecord anonymousBtr = (BlockTableRecord)trans.GetObject(anonymousBtrId, OpenMode.ForRead);
							//and all references to this block
							ObjectIdCollection blockRefIds = anonymousBtr.GetBlockReferenceIds(true, true);
							foreach (ObjectId id in blockRefIds) dynBlockRefs.Add(id);
						}
						res = dynBlockRefs;
						break;
					}
				}
				trans.Commit();
			}
			return res;
		}

		public ObjectIdCollection GetBlocksByName(string BlkName)
		{
			ObjectIdCollection res = new ObjectIdCollection();
			// acadApp.Document doc = acadApp.Application.DocumentManager.MdiActiveDocument;
			// Database db = doc.Database;
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
				try
				{
					// Get the block definition (if doesn't exist goes to "catch")
					BlockTableRecord btr = (BlockTableRecord)bt[BlkName].GetObject(OpenMode.ForRead);
					// Get any references
					res = btr.GetBlockReferenceIds(true, true);
					//trans.Commit();
				}
				catch (System.Exception)
				{
					// eKeyNotFound - Block definition does not exist
					// throw;
				}
			}
			return res;
		}

		public static void Mover(ObjectId id, Point3d fromPoint, Point3d toPoint)
		{
			Vector3d transVec = new Vector3d();
			transVec = toPoint - fromPoint;
			try
			{
				Document doc = acadApp.DocumentManager.MdiActiveDocument;
				Matrix3d m = Matrix3d.Displacement(Db.UcsToWcs(transVec));
				using (DocumentLock dlock = doc.LockDocument())
				using (Transaction tr = doc.TransactionManager.StartTransaction())
				{
					Entity ent = (Entity)tr.GetObject(id, OpenMode.ForWrite, false);
					ent.TransformBy(m);
					tr.Commit();
					acadApp.UpdateScreen();
				}
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show("Error in CADHelp.CADHelper.Mover(objectID)" + Environment.NewLine + ex.Message);
				// throw;
			}
		}

		public static void Mover(ref Entity ent, Point3d fromPoint, Point3d toPoint)
		{
			Vector3d transVec = new Vector3d();
			transVec = toPoint - fromPoint;
			try
			{
				Document doc = acadApp.DocumentManager.MdiActiveDocument;
				Matrix3d m = Matrix3d.Displacement(Db.UcsToWcs(transVec));
				using (DocumentLock dlock = doc.LockDocument())
				using (Transaction tr = doc.TransactionManager.StartTransaction())
				{
					ent.TransformBy(m);
					tr.Commit();
					acadApp.UpdateScreen();
				}

			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show("Error in CADHelp.CADHelper.Mover(Entity)" + Environment.NewLine + ex.Message);
			}
		}

		public static string GetUniqueBlockName(string baseName)
		{
            return CADHelp.Block.GetUniqueBlockName(baseName);
            #region old
            //CADHelper ch = new CADHelper();
            //string res = string.Empty;
            //int inc = 1;
            //bool nameExists = true;
            //string bname = string.Empty;
            //while (nameExists == true)
            //{
            //    bname = baseName + "-" + inc.ToString();
            //    nameExists = ch.BlockDefExists(bname);
            //    inc++;
            //}
            //if (bname != string.Empty) res = bname;
            //return res; 
            #endregion
		}

		public static void LayerFreeze(string LayerName, bool FreezeLayer)
		{
			using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
			{
				LayerTable ltab = (LayerTable)tr.GetObject(HostApplicationServices.WorkingDatabase.LayerTableId, OpenMode.ForWrite);
				try
				{
					LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(ltab[LayerName], OpenMode.ForWrite);
					ltr.IsFrozen = FreezeLayer;
				}
				catch (System.Exception)
				{
				}
				tr.Commit();
			}
		}

		public static ObjectId GetTableRecordId(ObjectId TableId, string Name)
		{
			ObjectId id = ObjectId.Null;
			using (Transaction tr = TableId.Database.TransactionManager.StartTransaction())
			{
				SymbolTable table = (SymbolTable)tr.GetObject(TableId, OpenMode.ForRead);
				if (table.Has(Name))
				{
					id = table[Name];
					if (!id.IsErased)
						return id;
					foreach (ObjectId recId in table)
					{
						if (!recId.IsErased)
						{
							SymbolTableRecord rec = (SymbolTableRecord)tr.GetObject(recId, OpenMode.ForRead);
							if (string.Compare(rec.Name, Name, true) == 0)
								return recId;
						}
					}
				}
			}
			return id;
		}

		public static int GetQtyFromRange(string val)
		{
			int res = 1;
			if (val.Length == 3)
			{
				if (val.StartsWith("A-"))
				{
					int ub = alpha.IndexOf(val.Substring(2));
					if (ub > 1)
					{
						res = ub + 1;
					}
				}
			}            
			return res;
		}

		internal static void AddLayout(string ItemID)
		{
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Database db = doc.Database;
			using (DocumentLock dlock = doc.LockDocument())
			{
				using (Transaction tr = doc.TransactionManager.StartTransaction())
				{
					DBDictionary LOD = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite);
					if (!LOD.Contains(ItemID))
					{
						LayoutManager lm = LayoutManager.Current;
						lm.CopyLayout("Layout1", ItemID);
						tr.Commit();
					}
				}
			}
		}

		internal static void AddLayout(string ItemID, Point2d ViewCenter)
		{
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Database db = doc.Database;
			using (DocumentLock dlock  = doc.LockDocument())
			{
				using (Transaction tr = doc.TransactionManager.StartTransaction())
				{
					DBDictionary LOD = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite);
					if (!LOD.Contains(ItemID))
					{
						LayoutManager lm = LayoutManager.Current;
						lm.CopyLayout("Layout1", ItemID);
						//DBDictionary lays = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
						//foreach (DBDictionaryEntry lay in lays)
						//{
						//    if (lay.Key.Equals(ItemID))
						//    {
						//        Layout alay = tr.GetObject(lay.Value, OpenMode.ForWrite) as Layout;
						//        alay.Initialize();
						//        ObjectIdCollection oids = alay.GetViewports();
						//        if (oids.Count > 1)
						//        {
						//            Viewport vp = tr.GetObject(oids[1], OpenMode.ForWrite) as Viewport;
						//            vp.ViewCenter = ViewCenter;
						//            vp.UpdateDisplay();
						//        }
						//        break;
						//    }
						//}
						tr.Commit();
					}
					
				}
			}
			
		}

		internal static Extents3d GetBoundingBox(ObjectIdCollection idCol)
		{
			Database db = HostApplicationServices.WorkingDatabase;
			Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
			
			Extents3d resExt = new Extents3d(new Point3d(0,0,0),new Point3d(0,0,0));
			using (Transaction tr = db.TransactionManager.StartTransaction())
			using (ViewTableRecord view = ed.GetCurrentView())
			{
				Matrix3d WCS2DCS = Matrix3d.PlaneToWorld(view.ViewDirection);
				WCS2DCS = Matrix3d.Displacement(view.Target - Point3d.Origin) * WCS2DCS;
				WCS2DCS = Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * WCS2DCS;
				WCS2DCS = WCS2DCS.Inverse();
				Entity ent = (Entity)tr.GetObject(idCol[0], OpenMode.ForRead);
				Extents3d ext = ent.GeometricExtents;
				ext.TransformBy(WCS2DCS);
				for (int i = 1; i < idCol.Count; i++)
				{
					ent = (Entity)tr.GetObject(idCol[i], OpenMode.ForRead);
					Extents3d tmp = ent.GeometricExtents;
					tmp.TransformBy(WCS2DCS);
					ext.AddExtents(tmp);
				}
				resExt = ext;
			}
			return resExt;
		}

		private void ZoomObjects(ObjectIdCollection idCol)
		{
			Document doc = Application.DocumentManager.MdiActiveDocument;
			Database db = doc.Database;
			Editor ed = doc.Editor;
			using (Transaction tr = db.TransactionManager.StartTransaction())
			using (ViewTableRecord view = ed.GetCurrentView())
			{
				Matrix3d WCS2DCS = Matrix3d.PlaneToWorld(view.ViewDirection);
				WCS2DCS = Matrix3d.Displacement(view.Target - Point3d.Origin) * WCS2DCS;
				WCS2DCS = Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * WCS2DCS;
				WCS2DCS = WCS2DCS.Inverse();
				Entity ent = (Entity)tr.GetObject(idCol[0], OpenMode.ForRead);
				Extents3d ext = ent.GeometricExtents;
				ext.TransformBy(WCS2DCS);
				for (int i = 1; i < idCol.Count; i++)
				{
					ent = (Entity)tr.GetObject(idCol[i], OpenMode.ForRead);
					Extents3d tmp = ent.GeometricExtents;
					tmp.TransformBy(WCS2DCS);
					ext.AddExtents(tmp);
				}
				double ratio = view.Width / view.Height;
				double width = ext.MaxPoint.X - ext.MinPoint.X;
				double height = ext.MaxPoint.Y - ext.MinPoint.Y;
				if (width > (height * ratio))
					height = width / ratio;
				Point2d center =
					new Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2.0, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2.0);
				view.Height = height;
				view.Width = width;
				view.CenterPoint = center;
				ed.SetCurrentView(view);
				tr.Commit();
			}
		}


		internal static void DeleteByObjectID(ObjectId oid)
		{
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Database db = doc.Database;

			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				DBObject dbo = trans.GetObject(oid, OpenMode.ForWrite);
				dbo.Erase();
				trans.Commit();
				doc.Editor.UpdateScreen();
			}
		}

		internal static double GetDistance()
		{
			double dres = 0;
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Editor ed = doc.Editor;
			PromptDistanceOptions pdo = new PromptDistanceOptions("Pick points:");
			pdo.AllowNone = false;
			pdo.AllowZero = false;
			pdo.AllowArbitraryInput = false;
			pdo.UseDashedLine = true;
			PromptDoubleResult pdr = ed.GetDistance(pdo);
			if (pdr.Status == PromptStatus.OK)
			{
				dres = pdr.Value;
				if (dres < 0) dres = (dres * -1);
			}
			return dres;
		}
		internal static Extents3d GetRectangle()
		{
			Extents3d xres = new Extents3d(new Point3d(), new Point3d());
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Editor ed = doc.Editor;
			PromptPointOptions ppo = new PromptPointOptions("Pick one corner:");
			ppo.AllowNone = false;
			PromptPointResult ppr = ed.GetPoint(ppo);
			if (ppr.Status == PromptStatus.OK)
			{
				PromptCornerOptions pco = new PromptCornerOptions("Pick other corner:", ppr.Value);
				PromptPointResult ppr2 = ed.GetCorner(pco);
				if (ppr2.Status == PromptStatus.OK)
				{
					double minX = 0;
					double minY = 0;
					double maxX = 0;
					double maxY = 0;
					if (ppr.Value.X < ppr2.Value.X)
					{
						minX = ppr.Value.X;
						maxX = ppr2.Value.X;
					}
					else
					{
						minX = ppr2.Value.X;
						maxX = ppr.Value.X;
					}
					if (ppr.Value.Y < ppr2.Value.Y)
					{
						minY = ppr.Value.Y;
						maxY = ppr2.Value.Y;
					}
					else
					{
						minY = ppr2.Value.Y;
						maxY = ppr.Value.Y;
					}
					xres = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
				}
			}
			return xres;
		}
		internal static void SetBlockRotation(ObjectId oid, double rot, bool relative)
		{
			using (DocumentLock dlock = acadApp.DocumentManager.MdiActiveDocument.LockDocument())
			using (Transaction trans = acadApp.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction())
			{
				try
				{
					BlockReference bref = trans.GetObject(oid, OpenMode.ForWrite) as BlockReference;
					Matrix3d mxRot; //, mxTrl;//                     
					if (relative)
					{
						mxRot = Matrix3d.Rotation(
							CADHelp.Ge.DegreesToRadians(rot),
							bref.BlockTransform.CoordinateSystem3d.Zaxis,
							bref.BlockTransform.CoordinateSystem3d.Origin);
					}
					else // absolute
					{
						// get the current rotation
						double preRot = bref.Rotation;
						double newRot = CADHelp.Ge.DegreesToRadians(rot);

						// calculate the revised rotation
						mxRot = Matrix3d.Rotation(
							newRot - preRot,
							bref.BlockTransform.CoordinateSystem3d.Zaxis,
							bref.BlockTransform.CoordinateSystem3d.Origin);
					}                    
					bref.TransformBy(mxRot);
					trans.Commit();
				}
				catch (System.Exception ex)
				{
					System.Windows.Forms.MessageBox.Show(ex.Message);
				}
			}

		}
		internal static void SetBlockRotation(ObjectId oid, double rot)
		{
			SetBlockRotation(oid, rot, false);
		}

		internal static double GetBlockRotation(ObjectId objectId)
		{
			double res = 0;
			using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
			{
				try
				{
					BlockReference bref = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
					res = CADHelp.Ge.RadiansToDegrees(bref.Rotation);
				}
				catch (System.Exception)
				{
				}
				tr.Commit();
			}
			return res;
		}

		internal static string DoubleToString(double dbl)
		{
			string res = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(dbl);
			return res;
		}
		internal static string DoubleToString(double dbl, DistanceUnitFormat DistFormat, int precision)
		{
			string res = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(dbl, DistFormat, precision);
			return res;
		}


		internal static double DecimalDouble(double d, int NumberOfPlaces)
		{
			double scale = Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(d))) + 1);
			return scale * Math.Round(d / scale, NumberOfPlaces);
		}

		internal static Point3d GetBlockLoction(ObjectId oid)
		{
            return CADHelp.Block.GetBlockLoction(oid);
            #region moved
            //Point3d resP = new Point3d();
            //using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            //{
            //    try
            //    {
            //        BlockReference bref = tr.GetObject(oid, OpenMode.ForRead) as BlockReference;
            //        resP = bref.Position;
            //    }
            //    catch (System.Exception)
            //    {
            //    }
            //    tr.Commit();
            //}
            //return resP; 
            #endregion
		}

		internal static Extents3d GetBlockBounds(ObjectId oid)
		{
			Extents3d rese = new Extents3d();
			using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
			{
				try
				{
					BlockReference bref = tr.GetObject(oid, OpenMode.ForRead) as BlockReference;
					rese = (Extents3d)bref.Bounds;
				}
				catch (System.Exception)
				{
				}
				tr.Commit();
			}
			return rese;
		}

		internal static ObjectId InsertImage(string path, string name, Point3d insPoint, double scale)
		{
			return InsertImage(path, name, insPoint, scale, new Point2dCollection());
		}
		internal static ObjectId InsertImage(string path, string name, Point3d insPoint)
		{
			return InsertImage(path, name, insPoint, 1, new Point2dCollection());
		}
		internal static ObjectId InsertImage(string path, string name, Point3d insPoint, double scale, Point2dCollection clipPoints)
		{
			ObjectId oidRes = ObjectId.Null;
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Database db = doc.Database;
			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				ObjectId dictId = RasterImageDef.GetImageDictionary(db);

				if (dictId.IsNull)
				{
					// Image dictionary doesn't exist, create new
					dictId = RasterImageDef.CreateImageDictionary(db);
				}

				// Open the image dictionary
				DBDictionary dict = (DBDictionary)trans.GetObject(dictId, OpenMode.ForRead);
				RasterImageDef rid = new RasterImageDef();
				ObjectId defID = ObjectId.Null;
				if (!dict.Contains(name))
				{
					rid.SourceFileName = path;
					rid.Load();
					dict.UpgradeOpen();
					defID = dict.SetAt(name, rid);
					trans.AddNewlyCreatedDBObject(rid, true);
				}
				else
				{
					defID = dict.GetAt(name);
					rid = trans.GetObject(defID, OpenMode.ForWrite) as RasterImageDef;
				}

				RasterImage ri = new RasterImage();
				ri.ImageDefId = defID;
				ri.ShowImage = true;
				ri.Orientation = new CoordinateSystem3d(
					insPoint,
					new Vector3d(scale,0,0),
					new Vector3d(0,scale,0));
				if (clipPoints.Count > 2)
				{
					Matrix3d mat = ri.PixelToModelTransform;
					mat = mat.Inverse();

					for (int cnt = 0; cnt < clipPoints.Count; cnt++)
					{
						Point3d pt = new Point3d(clipPoints[cnt].X, clipPoints[cnt].Y, 0.0);
						pt = pt.TransformBy(mat);
						clipPoints[cnt] = new Point2d(pt.X, pt.Y);
					}
					
					ri.IsClipped = true;
					// If the next line of code runs the raster image disappears from drawing
					ri.SetClipBoundary(ClipBoundaryType.Poly, clipPoints);
				}
				BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
				BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
				
				btr.AppendEntity(ri);
				trans.AddNewlyCreatedDBObject(ri, true);
				RasterImage.EnableReactors(true);
				ri.AssociateRasterDef(rid);
				trans.Commit();
				db.ResolveXrefs(false, true);
				oidRes = ri.ObjectId;
			}
			return oidRes;
		}

		internal static void ClipImage(ObjectId imgRefID, Point2dCollection clipPoints)
		{
			if (clipPoints.Count > 2)
			{
				Document doc = acadApp.DocumentManager.MdiActiveDocument;
				Database db = doc.Database;
				using (DocumentLock dlock = doc.LockDocument())
				using (Transaction trans = doc.TransactionManager.StartTransaction())
				{
					RasterImage ri = trans.GetObject(imgRefID, OpenMode.ForWrite) as RasterImage;
					if (ri != null)
					{
						Matrix3d mat = ri.PixelToModelTransform;
						mat = mat.Inverse();

						for (int cnt = 0; cnt < clipPoints.Count; cnt++)
						{
							Point3d pt = new Point3d(clipPoints[cnt].X, clipPoints[cnt].Y, 0.0);
							pt = pt.TransformBy(mat);
							clipPoints[cnt] = new Point2d(pt.X, pt.Y);
						}

						ri.IsClipped = true;
						ri.SetClipBoundary(ClipBoundaryType.Poly, clipPoints);
					}
					trans.Commit();
				}
			}
		}

		internal static void SendToBack(ObjectIdCollection oids)
		{
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Database db = doc.Database;
			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = trans.GetObject(doc.Database.BlockTableId, OpenMode.ForWrite) as BlockTable;
				if (bt != null)
				{
					BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
					if (btr != null)
					{
						DrawOrderTable dot = trans.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
						if (dot != null)
						{
							dot.MoveToBottom(oids);
						}
					}
				}
				trans.Commit();
			}
		}
		internal static void BringToFront(ObjectIdCollection oids)
		{
			Document doc = acadApp.DocumentManager.MdiActiveDocument;
			Database db = doc.Database;
			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				BlockTable bt = trans.GetObject(doc.Database.BlockTableId, OpenMode.ForWrite) as BlockTable;
				if (bt != null)
				{
					BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
					if (btr != null)
					{
						DrawOrderTable dot = trans.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
						if (dot != null)
						{
							dot.MoveToTop(oids);
						}
					}
				}
				trans.Commit();
			}
		}
		
		//internal static bool RibbonTabExists(string TabName)
		//{
		//    bool res = false;
		//    string menu = acadApp.GetSystemVariable("MENUNAME").ToString();
		//    menu += ".cuix";
		//    CustomizationSection cs;
		//    cs = new CustomizationSection(menu);
		//    RibbonRoot ribRoot = cs.MenuGroup.RibbonRoot;
		//    RibbonTabSourceCollection tabs = null;
		//    tabs = ribRoot.RibbonTabSources;
		//    foreach (RibbonTabSource rts in tabs)
		//    {
		//        if (rts.Name != null)
		//        {
		//            if (rts.Name.Equals(TabName))
		//            {
		//                res = true;
		//                break;
		//            }
		//        }
		//    }
		//    return res;
		//}
		//internal static RibbonTabSource GetRibbonTab(string TabName)
		//{
		//    RibbonTabSource res = null;
		//    string menu = acadApp.GetSystemVariable("MENUNAME").ToString();
		//    menu += ".cuix";
		//    CustomizationSection cs;
		//    cs = new CustomizationSection(menu);
		//    RibbonRoot ribRoot = cs.MenuGroup.RibbonRoot;
		//    RibbonTabSourceCollection tabs = null;
		//    tabs = ribRoot.RibbonTabSources;
		//    foreach (RibbonTabSource rts in tabs)
		//    {
		//        if (rts.Name != null)
		//        {
		//            if (rts.Name.Equals(TabName))
		//            {
		//                res = rts;
		//                break;
		//            }
		//        }
		//    }
		//    return res;
		//}

		//internal static RibbonPanelSource GetRibbonPanel(RibbonTabSource RibbonTab, string PanelName)
		//{
		//    RibbonPanelSource ribPanel = null;
		//    CustomizationSection cs = RibbonTab.CustomizationSection;
			
		//    // Panel pnl = PanelCollection
			

		//    return ribPanel;
		//}

		//internal static bool RibbonPanelExists(string TabName, string PanelName)
		//{
		//    bool res = false;
		//    if (RibbonTabExists(TabName))
		//    {
				

		//    }
		//    return res;
		//}

		internal static void Save(Document acDoc)
		{
			Save(acDoc, false);
		}

		internal static void Save(Document acDoc, bool CloseDrawing)
		{
			// Save the active drawing

			//Autodesk.AutoCAD.Interop.AcadDocument adoc =
			//    acDoc.GetAcadDocument() as Autodesk.AutoCAD.Interop.AcadDocument;

			//AutoCAD.AcadDocument adoc =
			//    acDoc.GetAcadDocument() as AutoCAD.AcadDocument;
			//if (adoc == null)
			//{
			//    System.Windows.Forms.MessageBox.Show("Drawing " + acDoc.Name + " was not saved.\n" +
			//        "WPICAD2013.CADHelper2013.Save was called.","Drawing may not be saved");
			//}
			//else
			//{
			//    adoc.Save();
			//    if (CloseDrawing)
			//    {
			//        acDoc.CloseAndDiscard();
			//    }
			//}
		}

		public static void attachXref(Document doc, string FullPath, Point3d InsertionPoint)
		{
			Database db = doc.Database;
			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				try
				{
					ObjectId xrefId = db.AttachXref(FullPath, Path.GetFileNameWithoutExtension(FullPath));
					// db.BindXrefs(xrefId, true);
					BlockReference blockRef = new BlockReference(InsertionPoint, xrefId);
					BlockTableRecord layoutBlock = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
					blockRef.ScaleFactors = new Scale3d(1, 1, 1);
					blockRef.Layer = "0";
					layoutBlock.AppendEntity(blockRef);
					trans.AddNewlyCreatedDBObject(blockRef, true);
				}
				catch
				{

				}
				trans.Commit();
			}
		}
		
		internal static BlockTableRecord ModelSpaceBlockTableRecord(Document doc, OpenMode mode)
		{
			ObjectId msID = SymbolUtilityServices.GetBlockModelSpaceId(doc.Database);
			BlockTableRecord btrRes = null;
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				btrRes = trans.GetObject(msID, mode) as BlockTableRecord;
				trans.Commit();
			}
			return btrRes;
		}

		internal static void BindAllXRefs(string FullPathToDwg)
		{
			// Binds all xrefs
			Document doc = acadApp.DocumentManager.Open(FullPathToDwg, false);
			acadApp.DocumentManager.MdiActiveDocument = doc;
			Editor ed = doc.Editor;

			TypedValue[] filterList = { new TypedValue(0, "INSERT") };

			PromptSelectionResult result = ed.SelectAll(new SelectionFilter(filterList));
			if (result.Status != PromptStatus.OK) return;

			ObjectId[] ids = result.Value.GetObjectIds();

			using (DocumentLock dlock = doc.LockDocument())
			using (Transaction trans = doc.TransactionManager.StartTransaction())
			{
				ObjectIdCollection xrefIds = new ObjectIdCollection();
				foreach (ObjectId id in ids)
				{
					BlockReference blockRef = (BlockReference)trans.GetObject(id, OpenMode.ForRead, false, true);
					ObjectId bId = blockRef.BlockTableRecord;
					if (!xrefIds.Contains(bId))
					{
						// xrefIds.Add(bId);
						BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bId, OpenMode.ForRead);
						if (btr.IsFromExternalReference)
							xrefIds.Add(bId); // process(btr);
					}
				}
				if (xrefIds.Count > 0)
				{
					doc.Database.BindXrefs(xrefIds, true);
				}
				trans.Commit();
			}
			Save(doc, true);
		}

		internal static bool LayerExists(string LayerName)
		{
			bool res = false;
			using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
			{
				LayerTable ltab = (LayerTable)tr.GetObject(HostApplicationServices.WorkingDatabase.LayerTableId, OpenMode.ForRead);
				res = ltab.Has(LayerName);
				////try
				//{
				//    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(ltab[LayerName], OpenMode.ForRead);
				//    res = true;
				//}
				//catch (System.Exception)
				//{
				//}
				tr.Commit();
			}
			return res;
		}

		internal static void DeleteReferenceLeadersByValue(string refLtr)
		{
			SelectionSet oss;
			Collection<TypedValue> collTV = new Collection<TypedValue>();
			collTV.Add(new TypedValue((int)DxfCode.Start, "MULTILEADER"));
			oss = CADHelp.SelectionTools.GetSelection(collTV, false, "");
			if (oss != null)
			{
				ObjectId[] oids = oss.GetObjectIds();
				if (oids.GetUpperBound(0) > -1)
				{
					Document doc = acadApp.DocumentManager.MdiActiveDocument;
					using (DocumentLock dlock = doc.LockDocument())
					using (Transaction trans = doc.TransactionManager.StartOpenCloseTransaction())
					{
						for (int i = 0; i <= oids.GetUpperBound(0); i++)
						{
							MLeader mldr = trans.GetObject(oids[i], OpenMode.ForRead) as MLeader;
							if (mldr != null)
							{
								if (mldr.ContentType == ContentType.BlockContent)
								{
									BlockTableRecord btr = trans.GetObject(mldr.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
									if (btr != null)
									{
										if (btr.HasAttributeDefinitions)
										{
											foreach (ObjectId oid in btr)
											{
												AttributeDefinition adef = trans.GetObject(oid, OpenMode.ForRead) as AttributeDefinition;
												if (adef != null)
												{
													AttributeReference aref = mldr.GetBlockAttribute(oid);
													if (aref.TextString == refLtr)
													{
														mldr.UpgradeOpen();
														// mldr.RecordGraphicsModified(true);
														mldr.Erase();
														break;
													}
												}
											}
										}
									}

								}
							}
						}
						trans.Commit();
					}
				}
			}
		}
	}

    public class Lines
    {
        public Lines()
        { }
        public static bool AddLine2d(ObjectId BlockTableRecordID, Line aLine)
        {
            bool res = false;
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {
                BlockTableRecord btr = trans.GetObject(BlockTableRecordID, OpenMode.ForWrite) as BlockTableRecord;
                if (btr != null)
                {
                    ObjectId lineOid = btr.AppendEntity(aLine);
                    trans.AddNewlyCreatedDBObject(aLine, true);
                    trans.Commit();
                    res = true;
                }
            }
            return res;
        }
    }
    public class Dimension
    {
        public Dimension()
        { }

        public static string DIM_STYLE_DIM_1 = "DIM1";
        public static string DIM_STYLE_DIM_2 = "DIM2";
        public static string DIM_STYLE_DIM_4 = "DIM4";
        public static string DIM_STYLE_DIM_8 = "DIM8";
        public static string DIM_STYLE_DIM_16 = "DIM16";
        public static string DIM_STYLE_DIM_FULL = "DIMFULL";
        public static string DIM_STYLE_DIM_1FC = "DIM1FC";
        public static string DIM_STYLE_DIM_1SCR = "DIM1SCR";
        public static string DIM_STYLE_BACKING = "Backing";

        public static string GetActiveDimStyle()
        {
            string res = string.Empty;
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {
                DimStyleTable dst = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                if (dst != null)
                {
                    DimStyleTableRecord dstr = trans.GetObject(db.Dimstyle, OpenMode.ForRead) as DimStyleTableRecord;
                    if (dstr != null)
                    {
                        res = dstr.Name;
                    }
                }
                trans.Commit();
            }
            return res;
        }

        public static bool SetActiveDimStyle(string DimStyleName)
        {
            bool res = false;
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {

                DimStyleTable dst = trans.GetObject(db.DimStyleTableId, OpenMode.ForWrite) as DimStyleTable;
                if (dst != null)
                {
                    if (!dst.Has(DimStyleName))
                    {
                        CADHelp.CADHelper.AddSettingsBlock();
                    }
                    if (dst.Has(DimStyleName))
                    {
                        db.Dimstyle = dst[DimStyleName];
                        DimStyleTableRecord dstr = trans.GetObject(dst[DimStyleName], OpenMode.ForWrite) as DimStyleTableRecord;
                        if (dstr != null)
                        {
                            db.SetDimstyleData(dstr);
                            res = true;
                        }
                    }
                    else
                    {
                        // Couldn't find the dim style
                    }
                }
                trans.Commit();
            }
            return res;
        }

        internal static ObjectId GetActiveDimStyleID()
        {
            ObjectId res = ObjectId.Null;
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {
                DimStyleTable dst = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                if (dst != null)
                {
                    res = db.Dimstyle;
                }
                trans.Commit();
            }
            return res;
        }
    }
    public class Block
    {
        public Block()
        { }
        internal static bool AddToBlockDef(ObjectId BlockTableRecordId, Collection<object> entObjects)
        {
            bool res=false;
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {
                BlockTableRecord btr = trans.GetObject(BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                if (btr != null)
                {
                    for (int i = 0; i < entObjects.Count; i++)
                    {
                        try
                        {
                            Entity ent = entObjects[i] as Entity;
                            if (ent != null)
                            {
                                btr.AppendEntity((Entity)entObjects[i]);
                                trans.AddNewlyCreatedDBObject((Entity)entObjects[i], true); 
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                            // throw;
                        }
                    }
                }
                trans.Commit();
            }
            return res;
        }

        internal static ObjectId GetBlockDef(string BlockDefName)
        {
            ObjectId resOID = ObjectId.Null;
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (bt != null)
                {
                    if (bt.Has(BlockDefName))
                    {
                        resOID = bt[BlockDefName];
                    }
                    else
                    {
                        BlockTableRecord btr = new BlockTableRecord();
                        btr.Name = BlockDefName;
                        bt.UpgradeOpen();
                        resOID = bt.Add(btr);
                        trans.AddNewlyCreatedDBObject(btr, true);
                        trans.Commit();
                    }
                }
            }
            return resOID;
        }

        internal static Point3d GetBlockLoction(ObjectId oid)
        {
            Point3d resP = new Point3d();
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockReference bref = tr.GetObject(oid, OpenMode.ForRead) as BlockReference;
                    resP = bref.Position;
                }
                catch (System.Exception)
                {
                }
                tr.Commit();
            }
            return resP;
        }
        internal static string GetUniqueBlockName(string baseName)
        {
            CADHelper ch = new CADHelper();
            string res = string.Empty;
            int inc = 1;
            bool nameExists = true;
            string bname = string.Empty;
            while (nameExists == true)
            {
                bname = baseName + "-" + inc.ToString();
                nameExists = ch.BlockDefExists(bname);
                inc++;
            }
            if (bname != string.Empty) res = bname;
            return res;
        }

        internal static void AddBlockToBlock(ObjectId TargetBlockDefID, Point3d SourceBlockLocation, string SourcePath, Point3d BlockScale, int Rotation)
        {
            using (Database db = HostApplicationServices.WorkingDatabase)
            using (Transaction trans = db.TransactionManager.StartOpenCloseTransaction())
            {
                BlockTableRecord btr = trans.GetObject(TargetBlockDefID, OpenMode.ForRead) as BlockTableRecord;
                if (btr != null)
                {
                    CADHelper ch = new CADHelper();
                    ObjectId sId = ch.InsertBlockDef(SourcePath);
                    using (BlockReference br = new BlockReference(SourceBlockLocation, sId))
                    {
                        btr.UpgradeOpen();
                        ObjectId newID = btr.AppendEntity(br);
                        trans.AddNewlyCreatedDBObject(br, true);
                    }
                }
                trans.Commit();
            }
        }
    }
	public class AUi
	{
		// User Interface Helper Constructor       
		AUi()
		{
		}
		//----------------------------         
		public static void
		PrintToCmdLine(string str)
		{
			Autodesk.AutoCAD.EditorInput.Editor
				ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
			ed.WriteMessage(str);
		}

		public static string
		PointToString(Point3d pt, DistanceUnitFormat unitType, int prec)
		{
			string x = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(pt.X, unitType, prec);
			string y = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(pt.Y, unitType, prec);
			string z = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(pt.Z, unitType, prec);

			return string.Format("({0}, {1}, {2})", x, y, z);
		}

		public static string
		PointToString(Point3d pt)
		{
			return PointToString(pt, Autodesk.AutoCAD.Runtime.DistanceUnitFormat.Current, -1);
		}
	}
	//------------------------------------------------------------------------

	public class Db
	{
		// DataBase Helper Constructor
		public Db()
		{
		}
		//---------------------------- 
		//// Transform from UCS to DCS
		[DllImport("acad.exe",
				  CallingConvention = CallingConvention.Cdecl,
				  EntryPoint = "acedTrans")
		]
		static extern int acedTrans(
		  double[] point,
		  IntPtr fromRb,
		  IntPtr toRb,
		  int disp,
		  double[] result
		);

		public static Extents2d UCStoDCS(Point3d first, Point3d second)
		{
			Extents2d resE = new Extents2d(0, 0, 0, 0);
			ResultBuffer rbFrom =
				new ResultBuffer(new TypedValue(5003, 1));
			ResultBuffer rbTo =
				new ResultBuffer(new TypedValue(5003, 2));
			double[] firres = new double[] { 0, 0, 0 };
			double[] secres = new double[] { 0, 0, 0 };
			// Transform the first point...

			acedTrans(
			first.ToArray(),
			rbFrom.UnmanagedObject,
			rbTo.UnmanagedObject,
			0,
			firres
			);

			// ... and the second

			acedTrans(
			second.ToArray(),
			rbFrom.UnmanagedObject,
			rbTo.UnmanagedObject,
			0,
			secres
			);

			// We can safely drop the Z-coord at this stage

			resE = new Extents2d(
			  firres[0],
			  firres[1],
			  secres[0],
			  secres[1]
			);

			return resE;
		}

		//ResultBuffer rbFrom =
		//new ResultBuffer(new TypedValue(5003, 1)),
		//          rbTo =
		//new ResultBuffer(new TypedValue(5003, 2));

		//double[] firres = new double[] { 0, 0, 0 };
		//double[] secres = new double[] { 0, 0, 0 };

		//// Transform the first point...

		//acedTrans(
		//first.ToArray(),
		//rbFrom.UnmanagedObject,
		//rbTo.UnmanagedObject,
		//0,
		//firres
		//);

		//// ... and the second

		//acedTrans(
		//second.ToArray(),
		//rbFrom.UnmanagedObject,
		//rbTo.UnmanagedObject,
		//0,
		//secres
		//);

		//// We can safely drop the Z-coord at this stage

		//Extents2d window =
		//new Extents2d(
		//  firres[0],
		//  firres[1],
		//  secres[0],
		//  secres[1]
		//);
		//---------------------------- 


		public static Database
		GetCurDatabase()
		{
			Database
				db = acadApp.DocumentManager.MdiActiveDocument.Database;
			return db;
		}

		public static Autodesk.AutoCAD.DatabaseServices.TransactionManager
		GetTransactionManager(Database db)
		{
			Autodesk.AutoCAD.DatabaseServices.TransactionManager
				tm = db.TransactionManager;
			return tm;
		}
		//----------------------------

		public static bool
		IsPaperSpace(Database db)
		{
			Debug.Assert(db != null);

			if (db.TileMode)
				return false;

			Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
			if (db.PaperSpaceVportId == ed.CurrentViewportObjectId)
				return true;

			return false;
		}

		public static Matrix3d
		GetUcsMatrix(Database db)
		{
			Debug.Assert(db != null);

			Point3d origin;
			Vector3d xAxis, yAxis, zAxis;

			if (IsPaperSpace(db))
			{
				origin = db.Pucsorg;
				xAxis = db.Pucsxdir;
				yAxis = db.Pucsydir;
			}
			else
			{
				origin = db.Ucsorg;
				xAxis = db.Ucsxdir;
				yAxis = db.Ucsydir;
			}

			zAxis = xAxis.CrossProduct(yAxis);

			return Matrix3d.AlignCoordinateSystem(Ge.kOrigin,
				Ge.kXAxis, Ge.kYAxis, Ge.kZAxis,
				origin, xAxis, yAxis, zAxis);
		}

		public static Point3d
		UcsToWcs(Point3d pt)
		{
			Matrix3d m = GetUcsMatrix(GetCurDatabase());

			return pt.TransformBy(m);
		}

		public static Vector3d
		UcsToWcs(Vector3d vec)
		{
			Matrix3d m = GetUcsMatrix(GetCurDatabase());

			return vec.TransformBy(m);
		}

		public static Point3d
		WcsToUcs(Point3d pt)
		{
			Matrix3d m = GetUcsMatrix(GetCurDatabase());

			return pt.TransformBy(m.Inverse());
		}
	}
	//------------------------------------------------------------------------
	
	public class Ge
	{
		// predefined constants for common angles
		public const double kPi = 3.14159265358979323846;
		public const double kHalfPi = 3.14159265358979323846 * 0.50;
		public const double kTwoPi = 3.14159265358979323846 * 2.00;

		public const double kRad0 = 0.0;
		public const double kRad45 = 3.14159265358979323846 * 0.25;
		public const double kRad90 = 3.14159265358979323846 * 0.50;
		public const double kRad135 = 3.14159265358979323846 * 0.75;
		public const double kRad180 = 3.14159265358979323846;
		public const double kRad270 = 3.14159265358979323846 * 1.5;
		public const double kRad360 = 3.14159265358979323846 * 2.0;

		// predefined values for common Points and Vectors
		public static readonly Point3d kOrigin = new Point3d(0.0, 0.0, 0.0);
		public static readonly Vector3d kXAxis = new Vector3d(1.0, 0.0, 0.0);
		public static readonly Vector3d kYAxis = new Vector3d(0.0, 1.0, 0.0);
		public static readonly Vector3d kZAxis = new Vector3d(0.0, 0.0, 1.0);

		// Geometry Helper Constructor
		public Ge()
		{
		}

		public static double
		RadiansToDegrees(double rads)
		{
			return rads * (180.0 / kPi);
		}

		public static double
		DegreesToRadians(double degrees)
		{
			return degrees * (kPi / 180.0);
		}
	}
}

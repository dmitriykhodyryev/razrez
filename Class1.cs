using System;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System.Xml;
using MgdAcApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Text.RegularExpressions;
using autowin = Autodesk.AutoCAD.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;


namespace kolonki
{
    public static class StringExtensions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }
    }
    //ися9 как по базе с скв
    //  string title = "STRING";
    //  bool contains = title.Contains("string", StringComparison.OrdinalIgnoreCase);
    public class Commands
    {
        static int _index = 1;
        public static void hole(double x, double y1, double y2, string pattiern)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 5.25, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 5.25, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) { acHatch.Visible = false; pattiern = "SOLID"; }
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                Circle acCirc = new Circle();
                acCirc.Center = new Point3d(x + 2.625, (y1 + y2)/2, 0);
                acCirc.Radius = 2.5;
                //acCirc.ColorIndex = 5;
                acBlkTblRec.AppendEntity(acCirc);
                acTrans.AddNewlyCreatedDBObject(acCirc, true);
                acHatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { acCirc.ObjectId });

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public static void hole(double x, double y1, double y2, string pattiern, int color)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 5.25, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 5.25, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;
                acPoly.ColorIndex = color;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) { acHatch.Visible = false; pattiern = "SOLID"; }
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                Circle acCirc = new Circle();
                acCirc.Center = new Point3d(x + 2.625, (y1 + y2) / 2, 0);
                acCirc.Radius = 2.5;
                acCirc.ColorIndex = color;
                //acCirc.ColorIndex = 5;
                acBlkTblRec.AppendEntity(acCirc);
                acTrans.AddNewlyCreatedDBObject(acCirc, true);
                acHatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { acCirc.ObjectId });

                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acHatch.ColorIndex = color;

                acTrans.Commit();
            }
        }
        public static void realhole(double x, double y1, double y2 )
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Add the inner boundary
                Circle acCirc = new Circle();
                acCirc.Center = new Point3d(x + 2.625, (y1 + y2) / 2, 0);
                acCirc.Radius = 2.5;
                //acCirc.ColorIndex = 5;
                acBlkTblRec.AppendEntity(acCirc);
                acTrans.AddNewlyCreatedDBObject(acCirc, true);

                acTrans.Commit();
            }
        }
        public static void igeincircle(double x, double y1, double y2,string ige ,double igesize)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Add the inner boundary
                Circle acCirc = new Circle();
                acCirc.Center = new Point3d(x+3 , (y1 + y2) / 2, 0);
                acCirc.Radius = 2.5;
                //acCirc.ColorIndex = 5;
                acBlkTblRec.AppendEntity(acCirc);
                acTrans.AddNewlyCreatedDBObject(acCirc, true);

                // Create a multiline text object
                MText acMText = new MText();
                    acMText.Location = new Point3d(x+3, (y1 + y2) / 2 + 1, 0);
                    acMText.Width = 84;
                    acMText.Contents = ige;
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    acMText.TextHeight = igesize;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);

                    //nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
      
                acTrans.Commit();

            }
        }
        public static void nohole(double x, double y1, double y2, string pattiern)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x , y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 5.25, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 5.25, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) { acHatch.Visible = false; pattiern = "SOLID"; }
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public static void nohole(double x, double y1, double y2, string pattiern, int color)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 5.25, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 5.25, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;
                acPoly.ColorIndex = color;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) { acHatch.Visible = false; pattiern = "SOLID"; }
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary

                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acHatch.ColorIndex = color;

                acTrans.Commit();
            }
        }
        public static void noholekonsist(double x, double y1, double y2, string pattiern)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 2.35, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 2.35, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) acHatch.Visible = false;
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public static void reznoholekonsist(double x, double y1, double y2, string pattiern)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x-2.35, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 2.35, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 2.35, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x-2.35, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) acHatch.Visible = false;
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public static void varnoholekonsist(double x, double y1, double y2, string pattiern)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x - 1, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x - 1, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) acHatch.Visible = false;
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public static void varnoholekonsist(double x, double y1, double y2, string pattiern, double varpipe)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x - varpipe, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + varpipe, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + varpipe, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x - varpipe, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                if ((pattiern == "") || (pattiern == null)) acHatch.Visible = false;
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public static void emptykonsist(double x, double y1, double y2)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x - 2.35, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 2.35, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 2.35, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x - 2.35, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);


                acTrans.Commit();
            }
        }
        public static void varemptykonsist(double x, double y1, double y2)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x - 1, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x - 1, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);


                acTrans.Commit();
            }
        }
        public static void varemptykonsist(double x, double y1, double y2,double varpipe)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x - varpipe, y1);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + varpipe, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + varpipe, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x - varpipe, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);


                acTrans.Commit();
            }
        }
        public static void leftline(double x, double y1, double y2)//as a tradition y2 goes slightly up тут колонки и линии что между слоями
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())//line closest to soil hole from left
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x, y1);
                var acPoly = new Polyline(2);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 43.25, y1), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 46.25, y2), 0.0, -1.0, -1.0);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }

            using (Transaction acTrans = db.TransactionManager.StartTransaction())//line under all texts at left
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint2 = new Point2d(x, 221);
                var acPoly2 = new Polyline(1);
                acPoly2.Normal = Vector3d.ZAxis;
                acPoly2.AddVertexAt(0, acPoint2, 0, -1, -1);
                acPoly2.AddVertexAt(1, new Point2d(x, y1), 0, -1, -1);
                acPoly2.Closed = false;
                acBlkTblRec.AppendEntity(acPoly2);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                acTrans.Commit();
            }

            using (Transaction acTrans = db.TransactionManager.StartTransaction())//vertical bar
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint3 = new Point2d(x+11.1, 221);
                var acPoly3 = new Polyline(1);
                acPoly3.Normal = Vector3d.ZAxis;
                acPoly3.AddVertexAt(0, acPoint3, 0, -1, -1);
                acPoly3.AddVertexAt(1, new Point2d(x+11.1, y1), 0, -1, -1);
                acPoly3.Closed = false;
                acBlkTblRec.AppendEntity(acPoly3);
                acTrans.AddNewlyCreatedDBObject(acPoly3, true);
                                acTrans.Commit();
            }

            using (Transaction acTrans = db.TransactionManager.StartTransaction())//vertical bar
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint4 = new Point2d(x + 22.2, 221);
                var acPoly4 = new Polyline(1);
                acPoly4.Normal = Vector3d.ZAxis;
                acPoly4.AddVertexAt(0, acPoint4, 0, -1, -1);
                acPoly4.AddVertexAt(1, new Point2d(x + 22.2, y1), 0, -1, -1);
                acPoly4.Closed = false;
                acBlkTblRec.AppendEntity(acPoly4);
                acTrans.AddNewlyCreatedDBObject(acPoly4, true);
                                acTrans.Commit();
            }

            using (Transaction acTrans = db.TransactionManager.StartTransaction())//vertical bar
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint5 = new Point2d(x + 33.3, 221);
                var acPoly5 = new Polyline(1);
                acPoly5.Normal = Vector3d.ZAxis;
                acPoly5.AddVertexAt(0, acPoint5, 0, -1, -1);
                acPoly5.AddVertexAt(1, new Point2d(x + 33.3, y1), 0, -1, -1);
                acPoly5.Closed = false;
                acBlkTblRec.AppendEntity(acPoly5);
                acTrans.AddNewlyCreatedDBObject(acPoly5, true);

                acTrans.Commit();
            }
        }
        public static void midline(double x, double y1, double y2)
        {
            double circlecentery = -y1 + y2;
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x, y1);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 3, y2), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 82.2, y2), 0.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(x + 85.2, y1), 0.0, -1.0, -1.0);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 85.2, y1);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 85.2, 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 85.2 + 11, y1);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 85.2 + 11, 221), 0, -1, -1);//11.1
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
             /*fe2 udalil
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 85.2+15, y1);//fe2 2016 22.2
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 85.2+15, 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
                */

           //it is right border of table
           using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 85.2 + 22.2 + 18.5, y1);//85.2 + 22.2+18.5
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 85.2 + 22.2 + 18.5, 221), 0, -1, -1);//85.2 + 22.2+18.5
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        } //and other lines to right - where water levelsтут колонки и линии что между слоями
        public static void dno(double x, double y1)//тут колонки и линии что между слоями
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 3, y1);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x-3, y1), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void upwline(string upwtext, double x) //тут появление воды
        {
            double y1 = double.Parse(upwtext);
            //text((x+167), y1, upwtext);

            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                //double y1 = double.Parse(upwtext);
                var acPoint = new Point2d(x + 164.3, 221 - y1 * 10);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 175.3, 221 - y1 * 10), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void upwrez(double x, double y1)//тут появление воды
        {
            //text((x+167), y1, upwtext);

            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x , y1);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x+1 ,  y1 +1), 0, -1, -1);
                acPoly.Closed = false;
                acPoly.ColorIndex = 5;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void uuwrez(double x, double y1)//тут установивш воды
        {
            //text((x+167), y1, upwtext);

            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x , y1 );
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x-1 ,  y1 -1), 0, -1, -1);
                acPoly.Closed = false;
                acPoly.ColorIndex = 5;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void uuwline(string uuwtext, double x)//тут установивш воды
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                double y1 = double.Parse(uuwtext);
                var acPoint = new Point2d(x + 175.3, 221 - y1 * 10);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 186.5, 221 - y1 * 10), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void palkaatfloor(int floor, double x)//ися7 1ап поменяю цвет чтоб увидеть где она http://cadhouse.narod.ru/articles/acad/acad_color_table.htm
        {//палка на подвале расстояний между скважинами
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, new Point2d(x, -60), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x, -70), 0, -1, -1);
                acPoly.Closed = false;
                //acPoly.ColorIndex = 1;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void palkaatfloor5floorsrazrez(int floor, double x)//ися7 1ап поменяю цвет чтоб увидеть где она http://cadhouse.narod.ru/articles/acad/acad_color_table.htm
        {//палка на подвале расстояний между скважинами
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, new Point2d(x, -30), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x, -40), 0, -1, -1);
                acPoly.Closed = false;
                //acPoly.ColorIndex = 1;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void palkapodkol(double x, double y)//тут колонки и линии что между слоям
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(1);
                acPoly.AddVertexAt(0, new Point2d(x, y), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x, 0), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void lineikafilledblock(double x, double y) //тут колонки
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;


            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y-10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x , y-10), 0, -1, -1);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public static void lineikafilledblock(double x, double y, int counter)//тут колонки
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y - 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, y - 10), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void lineikablock(double x, double y)//тут колонки
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;


            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y - 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, y - 10), 0, -1, -1);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                
                acTrans.Commit();
            }
        }
        public static void rezscalelineikablock(double x, double y, double scale)//тут колонки
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //http://exchange.autodesk.com/autocadmechanical/enu/online-help/AMECH_PP/2012/ENU/pages/WS1a9193826455f5ff2566ffd511ff6f8c7ca-3ee6.htm

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(x , y), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y + scale), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, y + scale), 0, -1, -1);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                acTrans.Commit();
            }
        }
        public static void rezscalelineikablocksolid(double x, double y, double scale)//тут колонки
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //http://exchange.autodesk.com/autocadmechanical/enu/online-help/AMECH_PP/2012/ENU/pages/WS1a9193826455f5ff2566ffd511ff6f8c7ca-3ee6.htm

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(x, y), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y + scale), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, y + scale), 0, -1, -1);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void putfish(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            string fileName = "C:\\AtlPProf\\rezxref";

            if (File.Exists(fileName))
            {
            var acPoint = new Point3d(-60, -60, 0);
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId xrefId = db.AttachXref(fileName, Path.GetFileNameWithoutExtension(fileName));
                BlockReference blockRef = new BlockReference(acPoint, xrefId);
                BlockTableRecord layoutBlock = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                //blockRef.ScaleFactors = new Scale3d(scaleFactor, scaleFactor, scaleFactor);
                blockRef.Layer = "0";
                layoutBlock.AppendEntity(blockRef);
                tr.AddNewlyCreatedDBObject(blockRef, true);
                tr.Commit();
            }}
        }
        public static void filledbot(double x, double y, double bot)//тут колонки
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, bot*10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, bot * 10), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x - 22.1, bot * 10);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 18.6, bot * 10), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void partialkubik(double x, double y, double bot)//кубик пробы
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(x , y), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y + bot ), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, y + bot ), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void vmstoosiY(double y, double bot)// вместо оси игрик
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoly = new Polyline(1);
                acPoly.AddVertexAt(0, new Point2d(0, y), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, y + bot), 0, -1, -1);
                acPoly.Closed = false;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                acTrans.Commit();
            }
        }
        public static void partialkubikempty(double x, double y, double deltay)//кубик пробы
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(x, y), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y + deltay), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, y + deltay), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                acTrans.Commit();
            }
        }
        public static void partialkubikscaled(double x, double y, double bot, double yscale)//кубик пробы
        {
            partialkubik(x, y, bot * yscale);
            vmstoosiY(      y, bot * yscale);
        }
        public static void partialkubikscaledempty(double x, double y, double bot, double yscale)//кубик пробы
        {
            partialkubikempty(x, y, bot * yscale);
            vmstoosiY(           y, bot * yscale);
        }
        public static void emptybot(double x, double y, double bot)//кубик пробы
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, bot * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, bot * 10), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x-22.1, bot*10);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 18.6, bot * 10), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void lineika(double visota, double x)
        {
            double y1 = visota;
            string lastblockcolor = "unknown";
            double remainingvisota = visota;
            double bottomoflinka = 22.1 - visota;
            double currenty = 221;
            int updouncounter = 1;

            bool shouldContinue = true;
            while (shouldContinue)
            {
                if (remainingvisota > 1) //it is not last meter in ruller bllack
                {
                    lineikablock(x, currenty);
                    lastblockcolor = "empty";
                    currenty -= 10;

                    text(x+3, currenty, updouncounter); updouncounter++;
                }
                else { break; shouldContinue = false; }
                remainingvisota -= 1;

                if (remainingvisota > 1) //it is not last meter in ruller white
                {
                    lineikafilledblock(x, currenty);
                    lastblockcolor = "full";
                    currenty -= 10;

                    text(x+3, currenty, updouncounter); updouncounter++;
                }
                else { break; shouldContinue = false; }
                remainingvisota -= 1;
            }

            if (remainingvisota>0) 
            {
                if (lastblockcolor == "empty") filledbot(x, currenty, bottomoflinka); 
                if (lastblockcolor == "full")  emptybot(x, currenty, bottomoflinka);
            }

            if (remainingvisota == 0)
            {
                currenty -= 10;
                text(x + 3, currenty, updouncounter);
            }
        }
        public static void lineikstartstop(double start, double stop)
        {
            double netxttextonlineyka = 1; netxttextonlineyka++;

            double distan = stop - start; if (distan < 0) { distan*=-1;};

            bool shouldContinue = true;
            while (shouldContinue)
            {
                if (distan > 0) 
                {
                    lineikablock(-10, start + 10);
                    start += 10;
                    text(-7, start , netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }
                distan -= 10;

                if (distan > 0) 
                {
                    lineikafilledblock(-10, start + 10);
                    start += 10;
                    text(-7, start , netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }
                distan -= 10;
            }
        }
        public static double lineikaoverpodvaltxt(double visota, double x, double uhorizm)
        {
            double afterdot = uhorizm - Math.Truncate(uhorizm);
            double partialblock = 1 - afterdot;
            filledbot(x, partialblock * 10, 0);
            double part_am = partialblock * 10;
            text(-7, part_am, uhorizm + partialblock);
            return part_am;
        }
        public static void fulllineka(double visota, double x, double uhorizm) 
        {
            double afterdot = uhorizm - Math.Truncate(uhorizm);
            double partialblock = 1 - afterdot;
            filledbot(x, partialblock * 10, 0);
            double part_am = partialblock * 10;
            text(-7, part_am, uhorizm + partialblock);


            double start = part_am; double stop = visota;
            double netxttextonlineyka = uhorizm + partialblock; netxttextonlineyka++;
            double distan = stop - start; if (distan < 0) { distan *= -1; };


            bool shouldContinue = true;
            while (shouldContinue)
            {
                if (distan > 0)
                {
                    lineikablock(-10, start + 10);
                    start += 10;
                    text(-7, start, netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }
                distan -= 10;

                if (distan > 0)
                {
                    lineikafilledblock(-10, start + 10);
                    start += 10;
                    text(-7, start, netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }
                distan -= 10;
            }
        }
        public static void lineikascaled(double visotam, double scale)
        {
            double start = 0;
            bool shouldContinue = true;
            while (shouldContinue)
            {
                if (visotam > 0)
                {
                    rezscalelineikablock(-10, start, scale );
                    start += scale;
                    //text(-7, start, netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }
                visotam -= 1;

                if (visotam > 0)
                {
                    rezscalelineikablocksolid(-10, start, scale);
                    start += scale;
                    //text(-7, start, netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }
                visotam -= 1;
            }
        }
        public void lineikaok26(double maxm, double horizm, double yscale)
        {
            //if not too high and not too low - it is still todo ok26
            double zeropixely = horizm * yscale * -1;
            lineykaydowntill0coordinate(zeropixely, yscale);

            double maxpixely = 0;
            if (maxm > 0) 
            {
                maxpixely = maxm * yscale + zeropixely;
                lineykayupabove0coordinate(maxpixely, zeropixely, yscale);
            }
        }
        public static void lineykaydowntill0coordinate(double visotapixls, double scale)
        {
            double start = visotapixls;
            bool shouldContinue = true;
            int netxttextonlineyka = 0 ;
            while (shouldContinue)
            {
                if (start > 0)
                {
                    rezscalelineikablock(-10, start, scale);
                    if (start < scale) { partialkubik(-10, 0, start); };
                    start -= scale;
                    text(-7, start + scale, netxttextonlineyka); netxttextonlineyka--;
                }
                else { break; }

                if (start > 0)
                {
                    rezscalelineikablocksolid(-10, start, scale);
                    if (start < scale) { partialkubikempty(-10, 0, start); };
                    start -= scale;
                    text(-7, start + scale, netxttextonlineyka); netxttextonlineyka--;
                }
                else { break; }
            }
        }
        public static void lineykayupabove0coordinate(double visotapixls,double zeropixely, double scale)
        {
            double start = zeropixely;
            //double finish = 
            bool shouldContinue = true;
            int netxttextonlineyka = 0;
            while (shouldContinue)
            {
                if (start < visotapixls)
                {
                    rezscalelineikablock(-10, start, scale);
                    start += scale;
                    text(-7, start - scale, netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }

                if (start < visotapixls)
                {
                    rezscalelineikablocksolid(-10, start, scale);
                    start += scale;
                    text(-7, start - scale, netxttextonlineyka); netxttextonlineyka++;
                }
                else { break; }
            }
        }
        private static void text(double x, double y, double x_3)
        {
            Document doc =
        MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Transaction tr =
              doc.TransactionManager.StartTransaction();
            using (tr)
            {
                // We'll add the objects to the model space
                BlockTable bt =
                  (BlockTable)tr.GetObject(
                    doc.Database.BlockTableId,
                    OpenMode.ForRead
                  );

                BlockTableRecord btr =
                  (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite
                  );

                DBText actext = new DBText();
                actext.SetDatabaseDefaults();
                actext.TextString = x_3.ToString();
                actext.Position = new Point3d(x, y, 0);
                actext.Height = 2;
                actext.Thickness = 22;

                btr.AppendEntity(actext);
                tr.AddNewlyCreatedDBObject(actext, true);

                tr.Commit();
            }
        }
        private static void wozrasttext(double x, double y, double x_3)
        {
            Document doc =
        MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Transaction tr =
              doc.TransactionManager.StartTransaction();
            using (tr)
            {
                // We'll add the objects to the model space
                BlockTable bt =
                  (BlockTable)tr.GetObject(
                    doc.Database.BlockTableId,
                    OpenMode.ForRead
                  );

                BlockTableRecord btr =
                  (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite
                  );

                DBText actext = new DBText();
                actext.SetDatabaseDefaults();
                actext.TextString = x_3.ToString();
                actext.Position = new Point3d(x, y, 0);
                actext.Height = 2;
                actext.WidthFactor = 0.6;
                actext.Thickness = 22;

                btr.AppendEntity(actext);
                tr.AddNewlyCreatedDBObject(actext, true);

                tr.Commit();
            }
        }
        public static void probamonbad(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y + 1.85), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x,        y + 1.85), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void probanarbad(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 0.925, y + 1.6), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void probamon(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 0.925, y + 1.6), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void probanar(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 1.85, y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 1.85, y + 1.85), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x, y + 1.85), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void probawod(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point3d(x+1, y, 0);
                //var acPoly = new Polyline(3);

                var cica = new Circle();
                cica.Center = acPoint;
                cica.Diameter = 2;
                //cica.Color = Color.FromDictionaryName();

                cica.Color = Color.FromColorIndex(ColorMethod.ByAci, 5);

                
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(cica);
                acTrans.AddNewlyCreatedDBObject(cica, true);
                // Create the hatch
                var acHatch = new Hatch();
                acHatch.PatternScale = 10;
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.ColorIndex = 5;
                acHatch.Associative = true;
                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { cica.ObjectId });
                // Add the inner boundary
                // Validate the hatch
                acHatch.EvaluateHatch(true);
                acTrans.Commit();
            }
        }
        public static void zagotovka(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x+20, y+256);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 205.1, y+256), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(x + 205.1, y+221), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(x + 20,y +221), 0, -1, -1);
                acPoly.Closed = true;
                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 31.1, y + 256);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 31.1, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 42.2, y + 256);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 42.2, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 53.3, y + 256);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 53.3, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 66.25, y + 256);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 66.25, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 79.1, y + 256);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 79.1, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 164.3, y + 256);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 164.3, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 186.5, y + 256);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 186.5, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 175.3, y + 245);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 175.3, y + 221), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoint = new Point2d(x + 164.3, y + 245);
                var acPoly = new Polyline(1);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x + 186.5, y + 245), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static double readzapas()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
                {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("zapas")) 
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("zapas = ", "");
                return double.Parse(tempzap);
            }

            else 
            {
                return 5;
            }
        }
        public static string readpodvaldistbetweenaccur()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("podvaldistbetweenaccur = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("podvaldistbetweenaccur = ", "");
                return tempzap;
            }

            else
            {
                return "0.00";
            }
        }
        public static string readrighttoholeaccur()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("righttoholeaccur = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("righttoholeaccur = ", "");
                return tempzap;
            }

            else
            {
                return "0.00";
            }
        }
        public static string readlefttoholeaccur()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("lefttoholeaccur = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("lefttoholeaccur = ", "");
                return tempzap;
            }

            else
            {
                return "0.00";
            }
        }
        public static string readpodvaldeepaccur()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("podvaldeepaccur = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("podvaldeepaccur = ", "");
                return tempzap;
            }

            else
            {
                return "0.00";
            }
        }
        public static string readpodvalotmetkaustaccur()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("podvalotmetkaustaccur = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("podvalotmetkaustaccur = ", "");
                return tempzap;
            }

            else
            {
                return "0.00";
            }
        }
        public static double readdiamskwa()
        {
            string diamskwa = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("diamskwa"))
                    {
                        diamskwa = line;
                    }
                }
                diamskwa = diamskwa.Replace("diamskwa = ", "");
                return double.Parse(diamskwa);
            }

            else
            {
                return 2;
            }
        }
        public static string cleansnow(string inshtrihovka)
        {
            inshtrihovka = inshtrihovka.Replace("снег", "");
            return inshtrihovka; 
        }
        public static string readprofshtrihmode()
        {   //ися4
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains("draw_ordinata = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("draw_ordinata = ", "");
                return tempzap;
            }

            else
            {
                return "0.00";
            }
        }
        [CommandMethod("geolog")]
        public void geologsize()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //dynamic acadObject = MgdAcApplication.AcadApplication;    acadObject.ZoomExtents();
            int debuglevel = 0;
            double dbleb = 0.0;
            dbleb = getopissize();
            double egesize = 0.0;
            egesize = Getegesize();
            double textsize = egesize;

            ComplexObjects.ExcelServ xlserv = new ComplexObjects.ExcelServ();
            #region open xl
            int goodskwacount = 0;
            int xlstartline = 5;//todo configit starting line
            int colforstartbur = 15;
            int colforstopbur = 16;
            int colforabsustbur = 17;//Q absolute height of hole
            int colforpodos = 6;
            int colforopisanie = 8;
            int colforvozrast = 18;
            int colforshtrih = 19;
            int colforkonsist = 20;
            int colforige = 11;
            int colforobraznar = 9;//vid mon i nar pereputan
            int colforobrazmon = 10;
            int colforobrazwod = 14;
            int colforupw = 12;//fe2 2016 switch 12 and 13 this line was 12 - no now i decide to change graphics
            int colforuuw = 13;//fe2 2016 switch 12 and 13

            var ofd =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "xls; xlsx",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr = ofd.ShowDialog();

            if (dr != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd.Filename
            );
            xlserv.Open(ofd.Filename);
            #endregion open xl
            skwazya[] skwazi = new skwazya[599];
            #region узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            for (int maxxllines = xlstartline; maxxllines < 3500; maxxllines++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv.UniRead(1, maxxllines, 1, ref cellcontent);//sheet row col
                if ((cellcontent != "") && (cellcontent != null)) //found NEW SKWA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                {
                    #region read first line
                    int maxgr = 0;
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    skwazi[goodskwacount] = new skwazya(goodskwacount);
                    skwazi[goodskwacount].name = cellcontent;
                    string startbur = " "; xlserv.UniRead(1, maxxllines, colforstartbur, ref startbur); if ((startbur != "") && (startbur != null)) { skwazi[goodskwacount].start = startbur; }
                    string stopbur = " "; xlserv.UniRead(1, maxxllines, colforstopbur, ref stopbur); if ((stopbur != "") && (stopbur != null)) { skwazi[goodskwacount].stop = stopbur; }
                    string absustbur = " "; xlserv.UniRead(1, maxxllines, colforabsustbur, ref absustbur); if ((absustbur != "") && (absustbur != null)) { skwazi[goodskwacount].absust = absustbur; } //Q column
                    skwazi[goodskwacount].absust = Regex.Replace(skwazi[goodskwacount].absust,",",".");
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    #endregion read first line

                    #region count Grunts IN SKWA
                    int gruntsinskwa = 0;
                    for (gruntsinskwa = 0; gruntsinskwa < 43; gruntsinskwa++)
                    {
                        string tempmaxlinesinskwa = "";
                        xlserv.UniRead(1, skwazi[goodskwacount].skwastartline + gruntsinskwa, colforopisanie, ref tempmaxlinesinskwa);
                        if ((tempmaxlinesinskwa == "") || (tempmaxlinesinskwa == null)) break;
                    }
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    #endregion count Grunts IN SKWA
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    skwazi[goodskwacount].layerscount = skwazi[goodskwacount].gruntsinskwa;
                    goodskwacount++;
                }
            }
            #endregion узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            #region дозаполнил скважины (а в планах = если есть только описания -и ничего кроме описаний, то нарисовать описания уже. а если описание и мощность, то еще лучше)
            for (int readmoreeveryskwa = 0; readmoreeveryskwa < goodskwacount; readmoreeveryskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[readmoreeveryskwa].gruntsinskwa; gruntinskwa++)
                {
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                    skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva = Regex.Replace(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva,",",".");
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforopisanie, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].opisanie);
                    //xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos + 1, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh);
                    if (gruntinskwa == 0) { skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva; }
                    else
                    {                 
                        // skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva - skwazi[readmoreeveryskwa].gruntiki[gruntinskwa-1].podoshva; }
                        //A__________                                               =  B______                                                  -   C____________
                        double B = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                        double C = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa - 1].podoshva);
                        double A = B - C;
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = A.ToString();
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobraznar, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obraznar);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobrazmon, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obrazmon);

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforige, ref ige);
                    if ((ige != "") && (ige != null)) skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].ige = ige;

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforvozrast, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].vozrast);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforshtrih, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    else 
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforkonsist, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    else 
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }

                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("tg_plast") || skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("TG_PLAST"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].haspalka = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = "";
                    }

                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforupw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].upw);
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforuuw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].uuw);
                }
                if (debuglevel > 1) MgdAcApplication.ShowAlertDialog("из екселя прочитана скважина  : " + readmoreeveryskwa + "по счету от 0. например 0,1,2");
            }
            #endregion

            #region calc all skwa for totalmosh
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                double temptotalmosh = 0.0;
                int templaycoutmax = skwazi[writeskvcnt].layerscount;
                for (int s = 0; s < templaycoutmax; s++)
                {
                    double unparce = 0.0;
                    string tempsting = skwazi[writeskvcnt].gruntiki[s].tolshmosh;
                    double.TryParse(tempsting, out unparce);
                    temptotalmosh = temptotalmosh + unparce;
                }
                //double totalmosh = skwazi[goodskwacount].gettotalmosh();
                skwazi[writeskvcnt].totalmosh = temptotalmosh;
            }
            #endregion calc all skwa for totalmosh
            #region Check all values
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                if ((skwazi[writeskvcnt].absust == "") || (skwazi[writeskvcnt].absust == null)) skwazi[writeskvcnt].absust = "0";
            }
            #endregion Check all values
            #region draw shapki calc FLOOR and MID levels
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region shapki
                zagotovka(210 * writeskvcnt, 0);
                text((22.5 + 210 * writeskvcnt), (265), "Начата   :");
                text((22 + 210 * writeskvcnt), (260), "Окончена  :");
                text((62 + 210 * writeskvcnt), (270), "Наименование : Скважина");
                text((112.5 + 210 * writeskvcnt), (265), "Отметка устья :");
                text((112.5 + 210 * writeskvcnt), (260), "Общая глубина  :");
                text((150 + 210 * writeskvcnt), (275), "Масштаб 1 : 100");
                text((111.3 + 210 * writeskvcnt), (250), "Наименование");
                text((117 + 210 * writeskvcnt), (244), "пород");
                text((120.75 + 210 * writeskvcnt), (237), "и");
                text((120 + 210 * writeskvcnt), (230), "их");
                text((110 + 210 * writeskvcnt), (223), "характеристика");
                text((167.73 + 210 * writeskvcnt), (252), "Сведения");
                text((169.73 + 210 * writeskvcnt), (246.75), "о воде");
                verttext(24 + 210 * writeskvcnt, 226.16, "Геологический");
                verttext(29.55 + 210 * writeskvcnt, 232.8, "индекс");
                verttext(35.1 + 210 * writeskvcnt, 230.8, "Мощность");
                verttext(40.65 + 210 * writeskvcnt, 232.3, "слоя, м");
                verttext(46.2 + 210 * writeskvcnt, 231.8, "Глубина");
                verttext(51.75 + 210 * writeskvcnt, 232.3, "слоя, м");
                verttext(57.3 + 210 * writeskvcnt, 227.5, "Абс. отметка");
                verttext(63.8 + 210 * writeskvcnt, 226.16, "подошвы слоя, м");
                verttext(70.25 + 210 * writeskvcnt, 231, "Геолого-");
                verttext(74.56 + 210 * writeskvcnt, 225.16, "литологический");
                verttext(78.88 + 210 * writeskvcnt, 232.8, "разрез");
                verttext(168.3 + 210 * writeskvcnt, 224.58, "Появление");
                verttext(173.85 + 210 * writeskvcnt, 229.58, "воды");
                verttext(179.4 + 210 * writeskvcnt, 226.25, "Установ.");
                verttext(184.95 + 210 * writeskvcnt, 226.75, "уровень");
                verttext(190.5 + 210 * writeskvcnt, 224.83, "Глубина отбора");
                verttext(199.75 + 210 * writeskvcnt, 230.83, "образцов");

                /*for cgei font
                 *              text((22 + 210 * writeskvcnt), (265), "Начата   :");
                             text((22 + 210 * writeskvcnt), (260), "Окончена :");
                             text((71 + 210 * writeskvcnt), (270), "Наименование :");
                             text((112.5 + 210 * writeskvcnt), (265), "Отметка устья :");
                             text((112.5 + 210 * writeskvcnt), (260), "Общая глубина :");
                             text((150 + 210 * writeskvcnt), (275), "Масштаб 1 : 100");
                             zagotovka(210 * writeskvcnt, 0);
                             text((104 + 210 * writeskvcnt), (250), "Наименование");
                             text((114.7 + 210 * writeskvcnt), (244), "пород");
                             text((120.75 + 210 * writeskvcnt), (237), "и");
                             text((120 + 210 * writeskvcnt), (230), "их");
                             text((101.25 + 210 * writeskvcnt), (223), "характеристика");
            
                             text((167.73 + 210 * writeskvcnt), (252), "Сведения");
                             text((169.73 + 210 * writeskvcnt), (246.75), "о воде");
                                */
                #endregion shapki
                #region calc floor
                //calc floor
                double tempnextceil = 221.0;
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    skwazi[writeskvcnt].gruntiki[readgr].konsiceiling = tempnextceil;
                    skwazi[writeskvcnt].gruntiki[readgr].konsifloor = tempnextceil - 10 * double.Parse(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh);
                    skwazi[writeskvcnt].gruntiki[readgr].circen = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2;
                    tempnextceil = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                }
                #endregion calc floor
            }
            #endregion draw shapki calc FLOOR and MID levels
            #region VALUES of shapki and LINEIKA
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                text((109 + 210 * writeskvcnt), (270), skwazi[writeskvcnt].name);
                text((43 + 210 * writeskvcnt), (265), skwazi[writeskvcnt].start);
                text((43 + 210 * writeskvcnt), (260), skwazi[writeskvcnt].stop);

                string absusttext = skwazi[writeskvcnt].absust;
                absusttext = absusttext.Replace(",", ".");
                double tempabsus = double.Parse(absusttext);
                text((144 + 210 * writeskvcnt), (265), tempabsus.ToString("0.00") + " м");

                text((144 + 210 * writeskvcnt), (260), skwazi[writeskvcnt].totalmosh.ToString("0.00") + " м");

                lineika(skwazi[writeskvcnt].totalmosh, 210 * writeskvcnt + 186.5);
            }
            #endregion
            #region drawall
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                    #region opisanie
                    //opisanie
                    double nextceil = 221.0;
                    for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                    {
                        //au27nextceil = MmText(skwazi[writeskvcnt].getlopisanieat(readgr), (122 + 210 * writeskvcnt), nextceil - 1) - 1;
                        nextceil = MTextbymenu(skwazi[writeskvcnt].getlopisanieat(readgr), (122 + 210 * writeskvcnt), nextceil - 1, dbleb) - 1;
                        double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                        if (floorbydata < nextceil) nextceil = floorbydata;

                        midline((79.1 + 210 * writeskvcnt), floorbydata, nextceil);
                    }
                    #endregion opisanie
                    #region write many to grunt line left
                    nextceil = 221.0;
                    for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                    {
                        skwazi[writeskvcnt].gruntiki[readgr].leftceiling = nextceil;
                        //real abs otmetka glubin podoshvi VOZRAST can be long so - main decesion is done after drawing vozrast
                        if ((skwazi[writeskvcnt].getltpodoshvaat(readgr) != "") && (skwazi[writeskvcnt].getltpodoshvaat(readgr) != null))
                        {
                            double glubinka = double.Parse(skwazi[writeskvcnt].absust) - double.Parse(skwazi[writeskvcnt].getltpodoshvaat(readgr));
                            double fornull1 = Textafterdot(glubinka.ToString(), (60 + 210 * writeskvcnt), nextceil - 1);
                        }

                        double fornull = Textafterdot(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh, (37 + 210 * writeskvcnt), nextceil - 1);
                        fornull = Textafterdot(skwazi[writeskvcnt].gruntiki[readgr].podoshva, (47 + 210 * writeskvcnt), nextceil - 1);

                        nextceil = MmTextvozrast(skwazi[writeskvcnt].gruntiki[readgr].vozrast, (25 + 210 * writeskvcnt), nextceil - 1) - 1;

                        double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                        if (floorbydata < nextceil) nextceil = floorbydata;
                        skwazi[writeskvcnt].gruntiki[readgr].leftfloor = nextceil;//todo 12au if deeper
                        leftline((20 + 210 * writeskvcnt), nextceil, floorbydata);
                    }
                    #endregion vozrast
                    #region tolch
                    //tolsh
                    for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                    {
                        double realcircentre = skwazi[writeskvcnt].gruntiki[readgr].circen;
                        double nullige = igemtext(skwazi[writeskvcnt].gruntiki[readgr].ige, (76.5 + 210 * writeskvcnt), realcircentre + 1.2, textsize);
                    }
                    #endregion
                    #region sneg 
                    for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                    {
                        /*if (skwazi[writeskvcnt].gruntiki[readgr].hassnow == "yes")
                        {
                            double realcircentre = skwazi[writeskvcnt].gruntiki[readgr].circen;
                            //double nul45lige = igemtext(skwazi[writeskvcnt].gruntiki[readgr].ige, (73 + 210 * writeskvcnt), realcircentre + 1.2, textsize);
                            double starcount = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 10;
                            double starcount5 = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 5;
                            double disttoborder = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor - Math.Truncate(starcount5) * 5) / 2;
                            double dlyaumensheniya = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            while (dlyaumensheniya > 1)
                            {
                                double currentvisotasneginki = skwazi[writeskvcnt].gruntiki[readgr].konsifloor + disttoborder + dlyaumensheniya;
                                double nullig2e = snegovik(skwazi[writeskvcnt].gruntiki[readgr].ige, (72.675 + 210 * writeskvcnt), currentvisotasneginki + 1.2);
                                dlyaumensheniya -= 5;
                            }
                        }
                         * */
                        skwazi[writeskvcnt].skwax = 72.675 + 210 * writeskvcnt;

                        if (skwazi[writeskvcnt].gruntiki[readgr].hassnow == "yes")
                        {
                            double realcircentre = skwazi[writeskvcnt].gruntiki[readgr].circen;
                            double starcount = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 5;
                            double dlyadeleniya = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            if (dlyadeleniya > 9)
                            {
                                dlyadeleniya -= 7;
                            }
                                                        
                            if (dlyadeleniya <= 6) { double nullig2e = snegovik("", skwazi[writeskvcnt].skwax, realcircentre); }
                            else
                            {
                                starcount = dlyadeleniya / 5;
                                double ddlyaumensheniya = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                                double dlinaotrezkasnow = ddlyaumensheniya / (Math.Truncate(starcount)+1);
                                ddlyaumensheniya -= dlinaotrezkasnow;

                                 while (ddlyaumensheniya > 3)
                                {
                                    double currentvisotasneginki = skwazi[writeskvcnt].gruntiki[readgr].konsifloor + ddlyaumensheniya;
                                    double nullig2e = snegovik("", skwazi[writeskvcnt].skwax, currentvisotasneginki);
                                    ddlyaumensheniya -= dlinaotrezkasnow;
                                }
                            }
                        }
                    }
                    #endregion sneg
                    #region palka
                    for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                    {
                        skwazi[writeskvcnt].skwax = 72.675 + 210 * writeskvcnt;

                        if (skwazi[writeskvcnt].gruntiki[readgr].haspalka == "yes")
                        {
                            double realcircentre = skwazi[writeskvcnt].gruntiki[readgr].circen;
                            // in snegowik we have starcount fe4 2016
                            palka(skwazi[writeskvcnt].skwax, skwazi[writeskvcnt].gruntiki[readgr].konsiceiling, skwazi[writeskvcnt].gruntiki[readgr].konsifloor);
                        }
                    }
                    #endregion sneg
                    #region shtrih
                    //shtrih
                    for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                    {
                        double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                        double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;


                        //Createchole((70 + 210 * writeskvcnt), realcircentre - circentreabovefloor, realcircentre + circentreabovefloor, skwazi[writeskvcnt].gruntiki[readgr].shtrih);
                        nohole((66.25 + 210 * writeskvcnt), floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih);
                        if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                        {
                            hole((73.85 + 210 * writeskvcnt), floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih);
                        }
                        else nohole((73.85 + 210 * writeskvcnt), floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih);

                        dno(73.85 + 210 * writeskvcnt, floorbydata);
                    }
                    #endregion
                    #region konsist
                    //konsist
                    for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                    {
                        string temp = skwazi[writeskvcnt].gruntiki[readgr].konsist;
                        if ((temp == "") || (temp == null)) continue;

                        double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                        double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                        noholekonsist((71.5 + 210 * writeskvcnt), floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].konsist);
                    }
                    #endregion
            }
            #endregion drawall
            #region draw wodka
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                int ifanywaterhere = 0;

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                    {
                        upwline(skwazi[writeskvcnt].gruntiki[readgr].upw, 210 * writeskvcnt);
                        double upwy = double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw);
                        text((210 * writeskvcnt + 167), 222 - upwy * 10, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw).ToString("0.00"));
                        ifanywaterhere++;
                    };

                    if ((skwazi[writeskvcnt].gruntiki[readgr].uuw != "") && (skwazi[writeskvcnt].gruntiki[readgr].uuw != null))
                    {
                        uuwline(skwazi[writeskvcnt].gruntiki[readgr].uuw, 210 * writeskvcnt);
                        double uuwy = double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw);
                        text((210 * writeskvcnt + 177), 222 - uuwy * 10, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw).ToString("0.00"));
                        ifanywaterhere++;
                    };
                }

                if (ifanywaterhere == 0)
                {
                    double viziblebottomy = skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor;
                    verttext((210 * writeskvcnt + 175), (222 + viziblebottomy) / 2 - 3, "Воды");
                    verttext((210 * writeskvcnt + 178), (222 + viziblebottomy) / 2 - 2, "нет");
                }
            }
            #endregion drawall

            #region XL Read samples of every HOLE of every LAYER
            for (int curskwa = 0; curskwa < goodskwacount; curskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string monstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazmon, ref monstring);
                    monstring = Regex.Replace(monstring, @"\t|\n|\r", ";");
                    string[] split = monstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "mon");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string narstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobraznar, ref narstring);
                    narstring = Regex.Replace(narstring, @"\t|\n|\r", ";");
                    string[] split = narstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "nar");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "nar");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string wodstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazwod, ref wodstring);
                    wodstring = Regex.Replace(wodstring, @"\t|\n|\r", ";");
                    string[] split = wodstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "wod");
                        }
                    }
                }
            }
            #endregion XL Read samples of every HOLE of every LAYER
            #region sort    samples //todo sort (take into consideration level)
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {
                skwazi[writeskvcnt].obriki = sortobriki(skwazi[writeskvcnt].obriki);
                // skwazi[writeskvcnt].obriki = pushl1tol2(skwazi[writeskvcnt].obriki);
                for (int curproba = 1; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    double diff = 0.0; diff = skwazi[writeskvcnt].obriki[curproba].bottom - skwazi[writeskvcnt].obriki[curproba - 1].bottom; if (diff < 0) { diff *= -1; }
                    if (diff < 0.19) skwazi[writeskvcnt].obriki[curproba].level = skwazi[writeskvcnt].obriki[curproba - 1].level + 1;
                }
            }
            #endregion
            #region drawall samples
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {                                                              // colonka by kolonka
                for (int curproba = 0; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "mon") probamon(193.2 + 210 * writeskvcnt + skwazi[writeskvcnt].obriki[curproba].level * 2.5, 221 - skwazi[writeskvcnt].obriki[curproba].bottom * 10);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "nar") probanar(193.2 + 210 * writeskvcnt + skwazi[writeskvcnt].obriki[curproba].level * 2.5, 221 - skwazi[writeskvcnt].obriki[curproba].bottom * 10);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "wod") probawod(193.2 + 210 * writeskvcnt + skwazi[writeskvcnt].obriki[curproba].level * 2.5, 221 - skwazi[writeskvcnt].obriki[curproba].bottom * 10);
                }
            }
            #endregion drawall probs

            #region task wodavpeske
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("PESOK") || skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("pesok"))
                    {
                        //search uuw ??upw?  in ALL GRUNTS or only this Grunt layer?
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double watertop = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                            double waterbot = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            double thislevel = 221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * 10;
                            //ne budu //if ((double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) > skwazi[writeskvcnt].gruntiki[readgr].konsiceiling)&&(double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) < skwazi[writeskvcnt].gruntiki[readgr].konsifloor)){  };
                            noholekonsist(71.5 + 210 * writeskvcnt, thislevel, waterbot, "SOLID");
                        }
                    }
                }
            }
            #endregion task wodavpeske
            xlserv.Close();
        }

        [CommandMethod("hidden_razrez_a")]
        public void razreznohatch()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region ask some
            //dynamic acadObject = MgdAcApplication.AcadApplication;//zoom
            //acadObject.ZoomExtents();//zoom
            double yscale = 10;//calculated in this way 1000/100= 10
            double xscale = 2; //calculated in this way 1000/500= 2
            double uhorizm = 60;
            string autohor = "";
            autohor = getautohor();
            if (autohor == "no") { uhorizm = getuhorizm(); };
            double zapas = 0;
            zapas = readzapas();
            string podvaldistbetweenaccur = readpodvaldistbetweenaccur();
            string lefttoholeaccur = readlefttoholeaccur();
            string righttoholeaccur = readrighttoholeaccur();
            string podvalotmetkaustaccur = readpodvalotmetkaustaccur();
            string podvaldeepaccur = readpodvaldeepaccur();
            double varpipe = readdiamskwa(); varpipe = varpipe / 2;
            double uhoriz = uhorizm * 10;
            double deepesthole = 999;

            double xdevidor = getxscale();
            xscale = 1000 / xdevidor;

            double ydevidor = getyscale();
            yscale = 1000 / ydevidor;

            double geoscale = 10;
            double geodevidor = getgeoscale();
            geoscale = 1000 / geodevidor;
            #endregion ask some

            ComplexObjects.ExcelServ xlserv = new ComplexObjects.ExcelServ();
            #region open xl
            var ofd =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "xls; xlsx",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );
            var dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK)
                return;
            // Display the name of the file and the contained sheets
            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd.Filename
            );
            #endregion open xl
            xlserv.Open(ofd.Filename);
            #region kolonki
            int goodskwacount = 0;
            int xlstartline = 5;//todo configit starting line
            int colforstartbur = 15;
            int colforstopbur = 16;
            int colforabsustbur = 17;//Q absolute height of hole
            int colforpodos = 6;
            int colforopisanie = 8;
            int colforvozrast = 18;
            int colforshtrih = 19;
            int colforkonsist = 20;
            int colforige = 11;
            int colforobrazmon = 10;
            int colforobraznar = 9;
            int colforobrazwod = 14;
            int colforupw = 12;
            int colforuuw = 13;
            int debuglevel = 0;
            int colforpiketplus = 3;
            int colforpiket = 2;
            double dbleb = 0.0;
            //dbleb = getopissize();
            double xbegin = 0.0;
            //xbegin = getxbegin();
            double egesize = 0.0;
            //double replace221= 0.0;
            egesize = Getegesize();
            double textsize = egesize;
            double deepestholepixels = 999;
            double maximmust = -1500;
            double maximgeoline = -1500;

            skwazya[] skwazi = new skwazya[1000];
            #region узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            for (int maxxllines = xlstartline; maxxllines < 5000; maxxllines++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv.UniRead(1, maxxllines, 1, ref cellcontent);//sheet row col
                if ((cellcontent != "") && (cellcontent != null)) //found NEW SKWA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                {
                    #region read first line
                    int maxgr = 0;
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    skwazi[goodskwacount] = new skwazya(goodskwacount);
                    skwazi[goodskwacount].name = cellcontent;
                    skwazi[goodskwacount].replace221 = 0.0;
                    string startbur = " "; xlserv.UniRead(1, maxxllines, colforstartbur, ref startbur); if ((startbur != "") && (startbur != null)) { skwazi[goodskwacount].start = startbur; }
                    string stopbur = " "; xlserv.UniRead(1, maxxllines, colforstopbur, ref stopbur); if ((stopbur != "") && (stopbur != null)) { skwazi[goodskwacount].stop = stopbur; }
                    string absustbur = " "; xlserv.UniRead(1, maxxllines, colforabsustbur, ref absustbur); if ((absustbur != "") && (absustbur != null)) { skwazi[goodskwacount].absust = absustbur;
                    skwazi[goodskwacount].absust = Regex.Replace(skwazi[goodskwacount].absust, ",", ".");} //Q column
                    string piket = " "; xlserv.UniRead(1, maxxllines, colforpiket, ref piket); if ((piket != "") && (piket != null)) { skwazi[goodskwacount].piket = piket; } //Q column
                    string piketplus = " "; xlserv.UniRead(1, maxxllines, colforpiketplus, ref piketplus); if ((piketplus != "") && (piketplus != null)) { skwazi[goodskwacount].piketplus = piketplus; skwazi[goodskwacount].piketplus = Regex.Replace(skwazi[goodskwacount].piketplus, ",", "."); } //Q column
                    skwazi[goodskwacount].skwax = (double.Parse(skwazi[goodskwacount].piket) * 100 + double.Parse(skwazi[goodskwacount].piketplus)) * xscale - xbegin;//todo start of left position at 0 even if starting piket is not 0/but 44meters
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    #endregion read first line

                    #region count Grunts IN SKWA
                    int gruntsinskwa = 0;
                    for (gruntsinskwa = 0; gruntsinskwa < 43; gruntsinskwa++)
                    {
                        string tempmaxlinesinskwa = "";
                        xlserv.UniRead(1, skwazi[goodskwacount].skwastartline + gruntsinskwa, colforopisanie, ref tempmaxlinesinskwa);
                        if ((tempmaxlinesinskwa == "") || (tempmaxlinesinskwa == null)) break;
                    }
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    #endregion count Grunts IN SKWA
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    skwazi[goodskwacount].layerscount = skwazi[goodskwacount].gruntsinskwa;
                    goodskwacount++;
                }
            }
            #endregion узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            #region дозаполнил скважины (а в планах = если есть только описания -и ничего кроме описаний, то нарисовать описания уже. а если описание и мощность, то еще лучше)
            for (int readmoreeveryskwa = 0; readmoreeveryskwa < goodskwacount; readmoreeveryskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[readmoreeveryskwa].gruntsinskwa; gruntinskwa++)
                {
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                    skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva = Regex.Replace(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva, ",", ".");
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforopisanie, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].opisanie);
                    //xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos + 1, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh);
                    if (gruntinskwa == 0) { skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva; }
                    else
                    {
                        // skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva - skwazi[readmoreeveryskwa].gruntiki[gruntinskwa-1].podoshva; }
                        //A__________                                               =  B______                                                  -   C____________
                        double B = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                        double C = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa - 1].podoshva);
                        double A = B - C;
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = A.ToString();
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobraznar, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obraznar);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobrazmon, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obrazmon);

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforige, ref ige);
                    if ((ige != "") && (ige != null)) skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].ige = ige;

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforvozrast, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].vozrast);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforshtrih, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforkonsist, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforupw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].upw);
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforuuw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].uuw);
                }
                if (debuglevel > 1) MgdAcApplication.ShowAlertDialog("из екселя прочитана скважина  : " + readmoreeveryskwa + "по счету от 0. например 0,1,2");
            }
            #endregion

            #region calc all skwa for totalmosh
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                double temptotalmosh = 0.0;
                int templaycoutmax = skwazi[writeskvcnt].layerscount;
                for (int s = 0; s < templaycoutmax; s++)
                {
                    double unparce = 0.0;
                    string tempsting = skwazi[writeskvcnt].gruntiki[s].tolshmosh;
                    double.TryParse(tempsting, out unparce);
                    temptotalmosh = temptotalmosh + unparce;
                }
                //double totalmosh = skwazi[goodskwacount].gettotalmosh();
                skwazi[writeskvcnt].totalmosh = temptotalmosh;
            }
            #endregion calc all skwa for totalmosh
            #region Check all values
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                if ((skwazi[writeskvcnt].absust == "") || (skwazi[writeskvcnt].absust == null)) skwazi[writeskvcnt].absust = "0";
                skwazi[writeskvcnt].absdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) - skwazi[writeskvcnt].totalmosh;
                skwazi[writeskvcnt].visualdeepestpnt = double.Parse(skwazi[writeskvcnt].absust)*yscale - skwazi[writeskvcnt].totalmosh*geoscale;

                if (deepesthole > skwazi[writeskvcnt].absdeepestpnt) { deepesthole = skwazi[writeskvcnt].absdeepestpnt; };
                if (deepestholepixels > skwazi[writeskvcnt].visualdeepestpnt) { deepestholepixels = skwazi[writeskvcnt].visualdeepestpnt; };

                double ustcontent = double.Parse(skwazi[writeskvcnt].absust);
                if (ustcontent > maximmust) { maximmust = ustcontent; };

               // if (double.Parse(cellcontent) > maximgeoline) { maximgeoline = double.Parse(cellcontent); };
            }
            deepestholepixels = deepestholepixels + zapas;
            #endregion Check all values

            if (autohor == "yes") { uhorizm = deepestholepixels / yscale ; uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ }//se29 todo print uhorizm
            else { uhorizm = getuhorizm(); uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ };

            #region calc FLOOR and MID levels
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                skwazi[writeskvcnt].replace221 = (double.Parse(skwazi[writeskvcnt].absust) - uhorizm) * yscale;
                double tempnextceil = skwazi[writeskvcnt].replace221;
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    skwazi[writeskvcnt].gruntiki[readgr].konsiceiling = tempnextceil;
                    skwazi[writeskvcnt].gruntiki[readgr].konsifloor = tempnextceil - geoscale * double.Parse(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh);//10 - > geoscale
                    skwazi[writeskvcnt].gruntiki[readgr].circen = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2;
                    tempnextceil = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                }
            }
            #endregion draw shapki calc FLOOR and MID levels

            #region drawall
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region text ige and other
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);

                    undertext(skwazi[writeskvcnt].skwax , floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur));
                    leftgeotxt(skwazi[writeskvcnt].skwax - varpipe , floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }
                #endregion
                #region konsist
                //konsist
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    string temp = skwazi[writeskvcnt].gruntiki[readgr].konsist;
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((temp == "") || (temp == null))
                    {
                        varemptykonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, varpipe);
                        continue;
                    }
                    varnoholekonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].konsist, varpipe);
                }
                #endregion
                #region sneg
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].hassnow == "yes")
                    {
                        double realcircentre = skwazi[writeskvcnt].gruntiki[readgr].circen;
                        double starcount = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 5;
                        double dlyadeleniya = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                        if (dlyadeleniya > 9)
                        {
                            dlyadeleniya -= 7;
                        }

                        if (dlyadeleniya <= 6) { double nullig2e = snegovik("", skwazi[writeskvcnt].skwax, (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2); }
                        else
                        {
                            starcount = dlyadeleniya / 5;
                            double ddlyaumensheniya = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling - skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                            double dlinaotrezkasnow = ddlyaumensheniya / (Math.Truncate(starcount) + 1);
                            ddlyaumensheniya -= dlinaotrezkasnow;

                            while (ddlyaumensheniya > 3)
                            {
                                double currentvisotasneginki = skwazi[writeskvcnt].gruntiki[readgr].konsifloor + ddlyaumensheniya;
                                double nullig2e = snegovik("", skwazi[writeskvcnt].skwax, currentvisotasneginki);
                                ddlyaumensheniya -= dlinaotrezkasnow;
                            }
                        }
                    }
                }
                #endregion sneg
                #region podval 
                text((skwazi[writeskvcnt].skwax), (-9), skwazi[writeskvcnt].name);
                //isa-1 podval
                string absusttext = skwazi[writeskvcnt].absust;
                absusttext = absusttext.Replace(",", ".");
                double tempabsus = double.Parse(absusttext);
                text((skwazi[writeskvcnt].skwax), (-19), tempabsus.ToString(podvalotmetkaustaccur));

                text((skwazi[writeskvcnt].skwax), (-29), skwazi[writeskvcnt].totalmosh.ToString(podvaldeepaccur));

                /*

                if (writeskvcnt > 0)
                {
                    text((skwazi[writeskvcnt - 1].skwax + (skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / 2), (-39), ((skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwazi[writeskvcnt].skwax / 2), (-39), (skwazi[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }               
                 * 
                 * */

                text((skwazi[writeskvcnt].skwax), (-49), skwazi[writeskvcnt].stop); //дата проходки

                palkaatfloor5floorsrazrez(4, skwazi[writeskvcnt].skwax);
                #endregion podval
            }
            #endregion drawall

            #region  isa-2  draw sorted dist at podval - we can use this sorted as main but it is bad for debugging - all skwaz mixed as the rezult and we have fk if input bad
            skwazya[] skwagood = new skwazya[goodskwacount];
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                skwagood[writeskvcnt] = skwazi[writeskvcnt];
            }

            skwazya[] skwasortedx = sortskwa(skwagood);

            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            { 
                #region podval

                if (writeskvcnt > 0)
                {
                    text((skwasortedx[writeskvcnt - 1].skwax + (skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / 2), (-39), ((skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwasortedx[writeskvcnt].skwax / 2), (-39), (skwasortedx[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }

                #endregion podval
            }

            #endregion draw sorted dist at podval



            #region XL Read samples of every HOLE of every LAYER
            for (int curskwa = 0; curskwa < goodskwacount; curskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string monstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazmon, ref monstring);
                    monstring = Regex.Replace(monstring, @"\t|\n|\r", ";");
                    string[] split = monstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "mon");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string narstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobraznar, ref narstring);
                    narstring = Regex.Replace(narstring, @"\t|\n|\r", ";");
                    string[] split = narstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "nar");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "nar");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string wodstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazwod, ref wodstring);
                    wodstring = Regex.Replace(wodstring, @"\t|\n|\r", ";");
                    string[] split = wodstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "wod");
                        }
                    }
                }
            }
            #endregion XL Read samples of every HOLE of every LAYER
            #region sort    samples //todo sort (take into consideration level)
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka 2015 explained:obraztsi treugolnik and kvadrat chtob ne nalazili odin na drugogo/ so smecheniem vpravo - pri pechati eto vidno po levelu
            {
                skwazi[writeskvcnt].obriki = sortobriki(skwazi[writeskvcnt].obriki);
                // skwazi[writeskvcnt].obriki = pushl1tol2(skwazi[writeskvcnt].obriki);
                for (int curproba = 1; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    double diff = 0.0; diff = skwazi[writeskvcnt].obriki[curproba].bottom - skwazi[writeskvcnt].obriki[curproba - 1].bottom; if (diff < 0) { diff *= -1; }
                    if (diff < 1.9/ geoscale) skwazi[writeskvcnt].obriki[curproba].level = skwazi[writeskvcnt].obriki[curproba - 1].level + 1;//do se25 0.19
                }
            }
            #endregion
            #region drawall samples
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {                                                              // colonka by kolonka
                for (int curproba = 0; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "mon") probamon(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "nar") probanar(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "wod") probawod(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                }
            }
            #endregion drawall probs
            #region draw wodka
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                int ifanywaterhere = 0;

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if ((skwazi[writeskvcnt].gruntiki[readgr].uuw != "") && (skwazi[writeskvcnt].gruntiki[readgr].uuw != null))
                    {
                        double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw) * geoscale;
                        uuwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw).ToString("0.00"), skwazi[writeskvcnt].stop);
                        ifanywaterhere++;
                    };

                    if (skwazi[writeskvcnt].gruntiki[readgr].uuw != skwazi[writeskvcnt].gruntiki[readgr].upw)
                    {
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            upwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw).ToString("0.00"), skwazi[writeskvcnt].start);
                            ifanywaterhere++;
                        };
                    }
                }

                /*if (ifanywaterhere == 0)
                {
                    double viziblebottomy = skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor;
                    verttext((210 * writeskvcnt + 175), (222 + viziblebottomy) / 2 - 3, "Воды");
                    verttext((210 * writeskvcnt + 178), (222 + viziblebottomy) / 2 - 2, "нет");
                }*/
            }
            #endregion drawall
            #region task wodavpeske
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("PESOK") || skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("pesok"))
                    {
                        //search uuw ??upw?  in ALL GRUNTS or only this Grunt layer?
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double watertop = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                            double waterbot = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            //ne budu //if ((double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) > skwazi[writeskvcnt].gruntiki[readgr].konsiceiling)&&(double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) < skwazi[writeskvcnt].gruntiki[readgr].konsifloor)){  };
                            //noholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID");
                            //noholekonsist(skwazi[writeskvcnt].skwax-2.35, thislevel, waterbot, "SOLID");
                            varnoholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID", varpipe);
                        }
                    }
                }
            }
            #endregion task wodavpeske
            #region task palochki pod kolonkoy
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                palkapodkol(skwazi[writeskvcnt].skwax, skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor);
            }
            #endregion task palochki pod kolonkoy
            #endregion kolonki
            xlserv.Close();

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "xls; xlsx",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);
            #region horizworks
            #region horizline
            #region detect xl maximum -----> kartacounter

            string light = "green";
            int kartacounter = 1;
            for (; kartacounter < 5000; kartacounter++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(1, kartacounter, 1, ref cellcontent);//sheet row col
                if (((cellcontent == "") || (cellcontent == null)) && (kartacounter == 1))
                {
                    light = "red"; break; ;
                }
                if ((cellcontent == "") || (cellcontent == null))
                {
                    break; ;
                }
            }
            //kartacounter is maximum lines in exclell
            #endregion detect xl maximum -----> kartacounter
            #region drawline_no_scaling
            double maxyininput = 0;
            double maxy = 0;

            string firstpntx = "0";
            string firstpntxplus = "0";
            string firstpnty = "0";

            string nextx = "0";
            string nextxplus = "0";
            string nexty = "0";
            /*
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            firstpntx = " ";
                            xlserv2.UniRead(1, 1, 1, ref firstpntx);//sheet row col
                             firstpntxplus = " "; xlserv2.UniRead(1, 1, 2, ref firstpntxplus);//sheet row col
                            firstpnty = " ";
                            xlserv2.UniRead(1, 1, 3, ref firstpnty);//sheet row col

                            var acPoint = new Point2d((double.Parse(firstpntx) * 100 + double.Parse(firstpntxplus))*xscale-xbegin, double.Parse(firstpnty) * yscale - uhoriz);

                            if (double.Parse(firstpnty) * yscale - uhoriz > maxy) {maxy = double.Parse(firstpnty) * yscale - uhoriz; }
                            var acPoly = new Polyline(kartacounter-1);
                            acPoly.Normal = Vector3d.ZAxis;
                            acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                            for (int lineinxl = 2; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                            {
                                 nextx = " "; xlserv2.UniRead(1, lineinxl, 1, ref nextx);//sheet row col
                                 nextxplus = " "; xlserv2.UniRead(1, lineinxl, 2, ref nextxplus);//sheet row col
                                nexty = " "; xlserv2.UniRead(1, lineinxl, 3, ref nexty);//sheet row col
                                acPoly.AddVertexAt(lineinxl - 1, new Point2d((double.Parse(nextx) * 100 + double.Parse(nextxplus))*xscale-xbegin, double.Parse(nexty) * yscale - uhoriz), 0, -1, -1);
                                if (double.Parse(nexty) * yscale - uhoriz > maxy) { maxy = double.Parse(nexty) * yscale - uhoriz; }
                            }
                            acPoly.Closed = false;

                            acBlkTblRec.AppendEntity(acPoly);
                            acTrans.AddNewlyCreatedDBObject(acPoly, true);
                            acTrans.Commit();
                        }
                    */
            #endregion drawline_no_scaling
            #region horiz load
            //#region #endregion horiz load
            double[] lineforscalingpikets = new double[5000];
            double[] lineforscalingplusos = new double[5000];
            double[] lineforscalingy = new double[5000];
            double[] lineforscalingx = new double[5000];
            string[] ordinatapart1 = new string[5000];
            string[] ordinatapart2 = new string[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                string cellcontentpk = " "; xlserv2.UniRead(1, lineinxl, 1, ref cellcontentpk);
                lineforscalingpikets[lineinxl] = double.Parse(cellcontentpk);
                string cellcontentps = " "; xlserv2.UniRead(1, lineinxl, 2, ref cellcontentps);
                lineforscalingplusos[lineinxl] = double.Parse(cellcontentps);
                string cellcontenty = " "; xlserv2.UniRead(1, lineinxl, 3, ref cellcontenty);
                lineforscalingy[lineinxl] = double.Parse(cellcontenty);
                lineforscalingx[lineinxl] = lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl];

                string ordinatapart1content = " "; xlserv2.UniRead(1, lineinxl, 4, ref ordinatapart1content);
                ordinatapart1[lineinxl] = ordinatapart1content;
                string ordinatapart2content = " "; xlserv2.UniRead(1, lineinxl, 5, ref ordinatapart2content);
                ordinatapart2[lineinxl] = ordinatapart2content;

                if (double.Parse(cellcontenty) > maximgeoline) { maximgeoline = double.Parse(cellcontenty); };
            }
            #endregion horiz load
            #region horiz scale
            //#region #endregion 
            double[] lineyscaled = new double[5000];
            double[] linexscaled = new double[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                linexscaled[lineinxl] = lineforscalingx[lineinxl] * xscale;
                lineyscaled[lineinxl] = lineforscalingy[lineinxl] * yscale;
            }
            #endregion horiz scale
            #region horiz scale again add USLOV HORIZ -conditional horizon
            //#region #endregion scale again
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                //lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhoriz * yscale;
                lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhorizm * yscale;
            }
            #endregion  scale again add USLOV HORIZ -conditional horizon
            #region horiz draw scaled
            //#region #endregion draw scaled
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var pl = new Polyline();
                for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                {
                    pl.AddVertexAt(lineinxl - 1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, 0, 0);//original here --> pl.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);last here --->pl.AddVertexAt(3, new Point2d(0, 10), 0, 0, 0); pl.Closed = true;
                }

                pl.Closed = false;
                acBlkTblRec.AppendEntity(pl);
                acTrans.AddNewlyCreatedDBObject(pl, true);
                acTrans.Commit();
            }
            #endregion draw scaled
            #region     task ordinata
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if ((ordinatapart2[lineinxl] != null) && (ordinatapart2[lineinxl] != "") && (ordinatapart2[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart2[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
                if ((ordinatapart1[lineinxl] != null) && (ordinatapart1[lineinxl] != "") && (ordinatapart1[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart1[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
            }
            #endregion  task ordinata

            #endregion horizline
            lineikaok26(maximgeoline, uhorizm, yscale);

            #region boxunderground
            double firstx = linexscaled[1];
            double firsty = lineyscaled[1];
            double lastx = linexscaled[kartacounter - 1];
            double lasty = lineyscaled[kartacounter - 1];

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(firstx, firsty), 0, -1, -1);  //do ok27 acPoly.AddVertexAt(0, new Point2d(firstx, firsty)   , 0, -1, -1); 
                acPoly.AddVertexAt(1, new Point2d(0, 0), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, lasty), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion boxunderground
            int podvalcount = 1;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Номеp скважины", -63, -9, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 2;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметка устья, м", -63, -19, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 3;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Глубина, м", -63, -29, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 4;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояние, м", -63, -39, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 5;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Дата пpоходки", -63, -49, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval

            string scalextext = "Масштаб по оси X 1 : " + xdevidor;
            lefttext(scalextext, -65, 12, 2);
            string scaleytext = "Масштаб по оси Y 1 : " + ydevidor;
            lefttext(scaleytext, -65, 7.5, 2);
            string geoscaletext = "Масштаб геол. 1 : " + geodevidor;
            lefttext(geoscaletext, -65, 3, 2);

            #endregion horizworks
            xlserv2.Close();
        }
        [CommandMethod("hidden_razrez_a_sboku_shtrihovka")]
        public void razreznohatchsboku()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region ask some
            //dynamic acadObject = MgdAcApplication.AcadApplication;//zoom
            //acadObject.ZoomExtents();//zoom
            double yscale = 10;//calculated in this way 1000/100= 10
            double xscale = 2; //calculated in this way 1000/500= 2
            double uhorizm = 60;
            string autohor = "";
            autohor = getautohor();
            if (autohor == "no") { uhorizm = getuhorizm(); };
            double zapas = 0;
            zapas = readzapas();
            string podvaldistbetweenaccur = readpodvaldistbetweenaccur();
            string lefttoholeaccur = readlefttoholeaccur();
            string righttoholeaccur = readrighttoholeaccur();
            string podvalotmetkaustaccur = readpodvalotmetkaustaccur();
            string podvaldeepaccur = readpodvaldeepaccur();
            double varpipe = readdiamskwa(); varpipe = varpipe / 2;
            double uhoriz = uhorizm * 10;
            double deepesthole = 999;

            double xdevidor = getxscale();
            xscale = 1000 / xdevidor;

            double ydevidor = getyscale();
            yscale = 1000 / ydevidor;

            double geoscale = 10;
            double geodevidor = getgeoscale();
            geoscale = 1000 / geodevidor;
            double maximgeoline = -1500;
            #endregion ask some

            ComplexObjects.ExcelServ xlserv = new ComplexObjects.ExcelServ();
            #region open xl
            var ofd =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "xls; xlsx",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );
            var dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK)
                return;
            // Display the name of the file and the contained sheets
            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd.Filename
            );
            #endregion open xl
            xlserv.Open(ofd.Filename);
            #region kolonki
            int goodskwacount = 0;
            int xlstartline = 5;//todo configit starting line
            int colforstartbur = 15;
            int colforstopbur = 16;
            int colforabsustbur = 17;//Q absolute height of hole
            int colforpodos = 6;
            int colforopisanie = 8;
            int colforvozrast = 18;
            int colforshtrih = 19;
            int colforkonsist = 20;
            int colforige = 11;
            int colforobraznar = 9;
            int colforobrazmon = 10;
            int colforobrazwod = 14;
            int colforupw = 12;
            int colforuuw = 13;
            int debuglevel = 0;
            int colforpiketplus = 3;
            int colforpiket = 2;
            double dbleb = 0.0;
            //dbleb = getopissize();
            double xbegin = 0.0;
            //xbegin = getxbegin();
            double egesize = 0.0;
            //double replace221= 0.0;
            egesize = Getegesize();
            double textsize = egesize;
            double deepestholepixels = 999;

            skwazya[] skwazi = new skwazya[1000];
            #region узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            for (int maxxllines = xlstartline; maxxllines < 5000; maxxllines++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv.UniRead(1, maxxllines, 1, ref cellcontent);//sheet row col
                if ((cellcontent != "") && (cellcontent != null)) //found NEW SKWA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                {
                    #region read first line
                    int maxgr = 0;
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    skwazi[goodskwacount] = new skwazya(goodskwacount);
                    skwazi[goodskwacount].name = cellcontent;
                    skwazi[goodskwacount].replace221 = 0.0;
                    string startbur = " "; xlserv.UniRead(1, maxxllines, colforstartbur, ref startbur); if ((startbur != "") && (startbur != null)) { skwazi[goodskwacount].start = startbur; }
                    string stopbur = " "; xlserv.UniRead(1, maxxllines, colforstopbur, ref stopbur); if ((stopbur != "") && (stopbur != null)) { skwazi[goodskwacount].stop = stopbur; }
                    string absustbur = " "; xlserv.UniRead(1, maxxllines, colforabsustbur, ref absustbur); if ((absustbur != "") && (absustbur != null)) { skwazi[goodskwacount].absust = absustbur;
                    skwazi[goodskwacount].absust = Regex.Replace(skwazi[goodskwacount].absust, ",", "."); }
                    string piket = " "; xlserv.UniRead(1, maxxllines, colforpiket, ref piket); if ((piket != "") && (piket != null)) { skwazi[goodskwacount].piket = piket; } //Q column
                    string piketplus = " "; xlserv.UniRead(1, maxxllines, colforpiketplus, ref piketplus); if ((piketplus != "") && (piketplus != null)) { skwazi[goodskwacount].piketplus = piketplus; 
                    skwazi[goodskwacount].piketplus = Regex.Replace(skwazi[goodskwacount].piketplus, ",", ".");} 
                    skwazi[goodskwacount].skwax = (double.Parse(skwazi[goodskwacount].piket) * 100 + double.Parse(skwazi[goodskwacount].piketplus)) * xscale - xbegin;//todo start of left position at 0 even if starting piket is not 0/but 44meters
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    #endregion read first line

                    #region count Grunts IN SKWA
                    int gruntsinskwa = 0;
                    for (gruntsinskwa = 0; gruntsinskwa < 43; gruntsinskwa++)
                    {
                        string tempmaxlinesinskwa = "";
                        xlserv.UniRead(1, skwazi[goodskwacount].skwastartline + gruntsinskwa, colforopisanie, ref tempmaxlinesinskwa);
                        if ((tempmaxlinesinskwa == "") || (tempmaxlinesinskwa == null)) break;
                    }
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    #endregion count Grunts IN SKWA
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    skwazi[goodskwacount].layerscount = skwazi[goodskwacount].gruntsinskwa;
                    goodskwacount++;
                }
            }
            #endregion узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            #region дозаполнил скважины (а в планах = если есть только описания -и ничего кроме описаний, то нарисовать описания уже. а если описание и мощность, то еще лучше)
            for (int readmoreeveryskwa = 0; readmoreeveryskwa < goodskwacount; readmoreeveryskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[readmoreeveryskwa].gruntsinskwa; gruntinskwa++)
                {
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                    skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva = Regex.Replace(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva, ",", ".");
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforopisanie, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].opisanie);
                    //xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos + 1, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh);
                    if (gruntinskwa == 0) { skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva; }
                    else
                    {
                        // skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva - skwazi[readmoreeveryskwa].gruntiki[gruntinskwa-1].podoshva; }
                        //A__________                                               =  B______                                                  -   C____________
                        double B = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                        double C = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa - 1].podoshva);
                        double A = B - C;
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = A.ToString();
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobraznar, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obraznar);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobrazmon, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obrazmon);

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforige, ref ige);
                    if ((ige != "") && (ige != null)) skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].ige = ige;

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforvozrast, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].vozrast);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforshtrih, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforkonsist, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforupw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].upw);
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforuuw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].uuw);
                }
                if (debuglevel > 1) MgdAcApplication.ShowAlertDialog("из екселя прочитана скважина  : " + readmoreeveryskwa + "по счету от 0. например 0,1,2");
            }
            #endregion

            #region calc all skwa for totalmosh
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                double temptotalmosh = 0.0;
                int templaycoutmax = skwazi[writeskvcnt].layerscount;
                for (int s = 0; s < templaycoutmax; s++)
                {
                    double unparce = 0.0;
                    string tempsting = skwazi[writeskvcnt].gruntiki[s].tolshmosh;
                    double.TryParse(tempsting, out unparce);
                    temptotalmosh = temptotalmosh + unparce;
                }
                //double totalmosh = skwazi[goodskwacount].gettotalmosh();
                skwazi[writeskvcnt].totalmosh = temptotalmosh;
            }
            #endregion calc all skwa for totalmosh
            #region Check all values
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                if ((skwazi[writeskvcnt].absust == "") || (skwazi[writeskvcnt].absust == null)) skwazi[writeskvcnt].absust = "0";
                skwazi[writeskvcnt].absdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) - skwazi[writeskvcnt].totalmosh;
                skwazi[writeskvcnt].visualdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) * yscale - skwazi[writeskvcnt].totalmosh * geoscale;

                if (deepesthole > skwazi[writeskvcnt].absdeepestpnt) { deepesthole = skwazi[writeskvcnt].absdeepestpnt; };
                if (deepestholepixels > skwazi[writeskvcnt].visualdeepestpnt) { deepestholepixels = skwazi[writeskvcnt].visualdeepestpnt; };
            }
            deepestholepixels = deepestholepixels + zapas;
            #endregion Check all values

            if (autohor == "yes") { uhorizm = deepestholepixels / yscale ; uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ }//se29 todo print uhorizm
            else { uhorizm = getuhorizm(); uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ };

            #region calc FLOOR and MID levels
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region calc floor
                skwazi[writeskvcnt].replace221 = (double.Parse(skwazi[writeskvcnt].absust) - uhorizm) * yscale;
                double tempnextceil = skwazi[writeskvcnt].replace221;
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    skwazi[writeskvcnt].gruntiki[readgr].konsiceiling = tempnextceil;
                    skwazi[writeskvcnt].gruntiki[readgr].konsifloor = tempnextceil - geoscale * double.Parse(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh);//10 - > geoscale
                    skwazi[writeskvcnt].gruntiki[readgr].circen = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2;
                    tempnextceil = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                }
                #endregion calc floor
            }
            #endregion draw shapki calc FLOOR and MID levels

            #region drawall
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region text ige and other
                /*for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);

                    undertext(skwazi[writeskvcnt].skwax, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur), varpipe);
                    leftgeotxt(skwazi[writeskvcnt].skwax - varpipe, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }*/

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }

                #endregion
                #region REZ shtrih - podoshva text - REZ
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    //left part of skwa
                    nohole(skwazi[writeskvcnt].skwax - 5.25 - varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih,93);//left part of skwa

                    //text(skwazi[writeskvcnt].skwax + 17, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax + 4, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur));

                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);
                    ///text(skwazi[writeskvcnt].skwax - 17, floorbydata, (tempabsuseverylayer  - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax - 12, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        hole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih,93);
                    }
                    else nohole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih,93);
                }
                #endregion
                #region konsist
                //konsist
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    string temp = skwazi[writeskvcnt].gruntiki[readgr].konsist;
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((temp == "") || (temp == null))
                    {
                        varemptykonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, varpipe);
                        continue;
                    }
                    varnoholekonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].konsist, varpipe);
                }
                #endregion
                #region podval
                text((skwazi[writeskvcnt].skwax), (-9), skwazi[writeskvcnt].name);

                string absusttext = skwazi[writeskvcnt].absust;
                absusttext = absusttext.Replace(",", ".");
                double tempabsus = double.Parse(absusttext);
                text((skwazi[writeskvcnt].skwax), (-19), tempabsus.ToString(podvalotmetkaustaccur));

                text((skwazi[writeskvcnt].skwax), (-29), skwazi[writeskvcnt].totalmosh.ToString(podvaldeepaccur));

                //fe9 2016 task to do sort before kalkulating distances between/ because holes are at random in source file 
                /* 
                if (writeskvcnt > 0)
                {
                    text((skwazi[writeskvcnt - 1].skwax + (skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / 2), (-39), ((skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwazi[writeskvcnt].skwax / 2), (-39), (skwazi[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }
*/
                text((skwazi[writeskvcnt].skwax), (-49), skwazi[writeskvcnt].stop);

                palkaatfloor5floorsrazrez(4, skwazi[writeskvcnt].skwax);
                #endregion podval
            }
            #endregion drawall

            #region    draw sorted dist at podval - we can use this sorted as main but it is bad for debugging - all skwaz mixed as the rezult and we have fk if input bad
            skwazya[] skwagood = new skwazya[goodskwacount];
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                skwagood[writeskvcnt] = skwazi[writeskvcnt];
            }

            skwazya[] skwasortedx = sortskwa(skwagood);

            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region podval

                if (writeskvcnt > 0)
                {
                    text((skwasortedx[writeskvcnt - 1].skwax + (skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / 2), (-39), ((skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwasortedx[writeskvcnt].skwax / 2), (-39), (skwasortedx[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }

                #endregion podval
            }

            #endregion draw sorted dist at podval


            #region XL Read samples of every HOLE of every LAYER
            for (int curskwa = 0; curskwa < goodskwacount; curskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string monstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazmon, ref monstring);
                    monstring = Regex.Replace(monstring, @"\t|\n|\r", ";");
                    string[] split = monstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "mon");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string narstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobraznar, ref narstring);
                    narstring = Regex.Replace(narstring, @"\t|\n|\r", ";");
                    string[] split = narstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "nar");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "nar");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string wodstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazwod, ref wodstring);
                    wodstring = Regex.Replace(wodstring, @"\t|\n|\r", ";");
                    string[] split = wodstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "wod");
                        }
                    }
                }
            }
            #endregion XL Read samples of every HOLE of every LAYER
            #region sort    samples //todo sort (take into consideration level)
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {
                skwazi[writeskvcnt].obriki = sortobriki(skwazi[writeskvcnt].obriki);
                // skwazi[writeskvcnt].obriki = pushl1tol2(skwazi[writeskvcnt].obriki);
                for (int curproba = 1; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    double diff = 0.0; diff = skwazi[writeskvcnt].obriki[curproba].bottom - skwazi[writeskvcnt].obriki[curproba - 1].bottom; if (diff < 0) { diff *= -1; }
                    if (diff < 1.9 / geoscale) skwazi[writeskvcnt].obriki[curproba].level = skwazi[writeskvcnt].obriki[curproba - 1].level + 1;//do se25 0.19
                }
            }
            #endregion
            #region drawall samples
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {                                                              // colonka by kolonka
                for (int curproba = 0; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "mon") probamon(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "nar") probanar(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "wod") probawod(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                }
            }
            #endregion drawall probs
            #region draw wodka
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                int ifanywaterhere = 0;

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if ((skwazi[writeskvcnt].gruntiki[readgr].uuw != "") && (skwazi[writeskvcnt].gruntiki[readgr].uuw != null))
                    {
                        double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw) * geoscale;
                        uuwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw).ToString("0.00"), skwazi[writeskvcnt].stop);
                        ifanywaterhere++;
                    };

                    if (skwazi[writeskvcnt].gruntiki[readgr].uuw != skwazi[writeskvcnt].gruntiki[readgr].upw)
                    {
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            upwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw).ToString("0.00"), skwazi[writeskvcnt].start);
                            ifanywaterhere++;
                        };
                    }
                }

                /*if (ifanywaterhere == 0)
                {
                    double viziblebottomy = skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor;
                    verttext((210 * writeskvcnt + 175), (222 + viziblebottomy) / 2 - 3, "Воды");
                    verttext((210 * writeskvcnt + 178), (222 + viziblebottomy) / 2 - 2, "нет");
                }*/
            }
            #endregion drawall
            #region task wodavpeske
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("PESOK") || skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("pesok"))
                    {
                        //search uuw ??upw?  in ALL GRUNTS or only this Grunt layer?
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double watertop = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                            double waterbot = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            //ne budu //if ((double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) > skwazi[writeskvcnt].gruntiki[readgr].konsiceiling)&&(double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) < skwazi[writeskvcnt].gruntiki[readgr].konsifloor)){  };
                            //noholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID");
                            //noholekonsist(skwazi[writeskvcnt].skwax-2.35, thislevel, waterbot, "SOLID");
                            varnoholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID", varpipe);
                        }
                    }
                }
            }
            #endregion task wodavpeske
            #region task palochki pod kolonkoy
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                palkapodkol(skwazi[writeskvcnt].skwax, skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor);
            }
            #endregion task palochki pod kolonkoy
            #endregion kolonki
            xlserv.Close();

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "xls; xlsx",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);
            #region horizworks
            #region horizline
            #region detect xl maximum -----> kartacounter
            
            string light = "green";
            int kartacounter = 1;
            for (; kartacounter < 5000; kartacounter++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(1, kartacounter, 1, ref cellcontent);//sheet row col
                if (((cellcontent == "") || (cellcontent == null)) && (kartacounter == 1))
                {
                    light = "red"; break; ;
                }
                if ((cellcontent == "") || (cellcontent == null))
                {
                    break; ;
                }
            }
            //kartacounter is maximum lines in exclell
            #endregion detect xl maximum -----> kartacounter
            #region drawline_no_scaling
            double maxyininput = 0;
            double maxy = 0;

            string firstpntx = "0";
            string firstpntxplus = "0";
            string firstpnty = "0";

            string nextx = "0";
            string nextxplus = "0";
            string nexty = "0";
            /*
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            firstpntx = " ";
                            xlserv2.UniRead(1, 1, 1, ref firstpntx);//sheet row col
                             firstpntxplus = " "; xlserv2.UniRead(1, 1, 2, ref firstpntxplus);//sheet row col
                            firstpnty = " ";
                            xlserv2.UniRead(1, 1, 3, ref firstpnty);//sheet row col

                            var acPoint = new Point2d((double.Parse(firstpntx) * 100 + double.Parse(firstpntxplus))*xscale-xbegin, double.Parse(firstpnty) * yscale - uhoriz);

                            if (double.Parse(firstpnty) * yscale - uhoriz > maxy) {maxy = double.Parse(firstpnty) * yscale - uhoriz; }
                            var acPoly = new Polyline(kartacounter-1);
                            acPoly.Normal = Vector3d.ZAxis;
                            acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                            for (int lineinxl = 2; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                            {
                                 nextx = " "; xlserv2.UniRead(1, lineinxl, 1, ref nextx);//sheet row col
                                 nextxplus = " "; xlserv2.UniRead(1, lineinxl, 2, ref nextxplus);//sheet row col
                                nexty = " "; xlserv2.UniRead(1, lineinxl, 3, ref nexty);//sheet row col
                                acPoly.AddVertexAt(lineinxl - 1, new Point2d((double.Parse(nextx) * 100 + double.Parse(nextxplus))*xscale-xbegin, double.Parse(nexty) * yscale - uhoriz), 0, -1, -1);
                                if (double.Parse(nexty) * yscale - uhoriz > maxy) { maxy = double.Parse(nexty) * yscale - uhoriz; }
                            }
                            acPoly.Closed = false;

                            acBlkTblRec.AppendEntity(acPoly);
                            acTrans.AddNewlyCreatedDBObject(acPoly, true);
                            acTrans.Commit();
                        }
                    */
            #endregion drawline_no_scaling
            #region horiz load
            //#region #endregion horiz load
            double[] lineforscalingpikets = new double[5000];
            double[] lineforscalingplusos = new double[5000];
            double[] lineforscalingy = new double[5000];
            double[] lineforscalingx = new double[5000];
            string[] ordinatapart1 = new string[5000];
            string[] ordinatapart2 = new string[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                string cellcontentpk = " "; xlserv2.UniRead(1, lineinxl, 1, ref cellcontentpk);
                lineforscalingpikets[lineinxl] = double.Parse(cellcontentpk);
                string cellcontentps = " "; xlserv2.UniRead(1, lineinxl, 2, ref cellcontentps);
                lineforscalingplusos[lineinxl] = double.Parse(cellcontentps);
                string cellcontenty = " "; xlserv2.UniRead(1, lineinxl, 3, ref cellcontenty);
                lineforscalingy[lineinxl] = double.Parse(cellcontenty);
                lineforscalingx[lineinxl] = lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl];

                string ordinatapart1content = " "; xlserv2.UniRead(1, lineinxl, 4, ref ordinatapart1content);
                ordinatapart1[lineinxl] = ordinatapart1content;
                string ordinatapart2content = " "; xlserv2.UniRead(1, lineinxl, 5, ref ordinatapart2content);
                ordinatapart2[lineinxl] = ordinatapart2content;

                if (double.Parse(cellcontenty) > maximgeoline) { maximgeoline = double.Parse(cellcontenty); };
            }
            #endregion horiz load
            #region horiz scale
            //#region #endregion 
            double[] lineyscaled = new double[5000];
            double[] linexscaled = new double[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                linexscaled[lineinxl] = lineforscalingx[lineinxl] * xscale;
                lineyscaled[lineinxl] = lineforscalingy[lineinxl] * yscale;
            }
            #endregion horiz scale
            #region horiz scale again add USLOV HORIZ -conditional horizon
            //#region #endregion scale again
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                //lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhoriz * yscale;
                lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhorizm * yscale;
            }
            #endregion  scale again add USLOV HORIZ -conditional horizon
            #region horiz draw scaled
            //#region #endregion draw scaled
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var pl = new Polyline();
                for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                {
                    pl.AddVertexAt(lineinxl - 1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, 0, 0);//original here --> pl.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);last here --->pl.AddVertexAt(3, new Point2d(0, 10), 0, 0, 0); pl.Closed = true;
                }

                pl.Closed = false;
                acBlkTblRec.AppendEntity(pl);
                acTrans.AddNewlyCreatedDBObject(pl, true);
                acTrans.Commit();
            }
            #endregion draw scaled
            #region     task ordinata
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if ((ordinatapart2[lineinxl] != null) && (ordinatapart2[lineinxl] != "") && (ordinatapart2[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart2[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
                if ((ordinatapart1[lineinxl] != null) && (ordinatapart1[lineinxl] != "") && (ordinatapart1[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart1[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
            }
            #endregion  task ordinata

            #endregion horizline
            lineikaok26(maximgeoline, uhorizm, yscale); 

            #region boxunderground
            double firstx = linexscaled[1];
            double firsty = lineyscaled[1];
            double lastx = linexscaled[kartacounter - 1];
            double lasty = lineyscaled[kartacounter - 1];

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(firstx, firsty), 0, -1, -1);  //do ok27 acPoly.AddVertexAt(0, new Point2d(firstx, firsty)   , 0, -1, -1); 
                acPoly.AddVertexAt(1, new Point2d(0, 0), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, lasty), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion boxunderground
            int podvalcount = 1;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Номеp скважины", -63, -9, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 2;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметка устья, м", -63, -19, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 3;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Глубина, м", -63, -29, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 4;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояние, м", -63, -39, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 5;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Дата пpоходки", -63, -49, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-65, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-65, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval

            string scalextext = "Масштаб по оси X 1 : " + xdevidor;
            lefttext(scalextext, -65, 12, 2);
            string scaleytext = "Масштаб по оси Y 1 : " + ydevidor;
            lefttext(scaleytext, -65, 7.5, 2);
            string geoscaletext = "Масштаб геол. 1 : " + geodevidor;
            lefttext(geoscaletext, -65, 3, 2);

            #endregion horizworks
            xlserv2.Close();
        }

        [CommandMethod("razrez_a")]
        public void profvrazrbezbok()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region ask some
            //dynamic acadObject = MgdAcApplication.AcadApplication;//zoom
            //acadObject.ZoomExtents();//zoom
            double yscale = 10;//calculated in this way 1000/100= 10
            double xscale = 2; //calculated in this way 1000/500= 2
            double uhorizm = 60;
            string autohor = "";

            autohor = readconfig(@"autohor = ");//isa18 ap18 task dont ask too many questions
            if((autohor!="no")&&(autohor!="yes")) 
            {                
                autohor = getautohor();
            } 

            if (autohor == "no") { uhorizm = getuhorizm(); };
            double zapas = 0;
            zapas = readzapas();
            string podvaldistbetweenaccur = readpodvaldistbetweenaccur();
            string lefttoholeaccur = readlefttoholeaccur();
            string righttoholeaccur = readrighttoholeaccur();
            string podvalotmetkaustaccur = readpodvalotmetkaustaccur();
            string podvaldeepaccur = readpodvaldeepaccur();
            double varpipe = readdiamskwa(); varpipe = varpipe / 2;
            double uhoriz = uhorizm * 10;
            double deepesthole = 999;

            double xdevidor = getxscale();
            xscale = 1000 / xdevidor;

            double ydevidor = getyscale();
            yscale = 1000 / ydevidor;

            double geoscale = 10;
            double geodevidor = getgeoscale();
            geoscale = 1000 / geodevidor;
            double maximgeoline = -1500;
            #endregion ask some

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);
            #region horizline LOAD
            #region detect xl maximum -----> kartacounter

            string light = "green";
            int kartacounter = 1;
            for (; kartacounter < 5000; kartacounter++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(1, kartacounter, 1, ref cellcontent);//sheet row col
                if (((cellcontent == "") || (cellcontent == null)) && (kartacounter == 1))
                {
                    light = "red"; break; ;
                }
                if ((cellcontent == "") || (cellcontent == null))
                {
                    break; ;
                }
            }
            //kartacounter is maximum lines in exclell
            #endregion detect xl maximum -----> kartacounter
            #region drawline_no_scaling
            double maxyininput = 0;
            double maxy = 0;

            string firstpntx = "0";
            string firstpntxplus = "0";
            string firstpnty = "0";

            string nextx = "0";
            string nextxplus = "0";
            string nexty = "0";
            /*
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            firstpntx = " ";
                            xlserv2.UniRead(1, 1, 1, ref firstpntx);//sheet row col
                             firstpntxplus = " "; xlserv2.UniRead(1, 1, 2, ref firstpntxplus);//sheet row col
                            firstpnty = " ";
                            xlserv2.UniRead(1, 1, 3, ref firstpnty);//sheet row col

                            var acPoint = new Point2d((double.Parse(firstpntx) * 100 + double.Parse(firstpntxplus))*xscale-xbegin, double.Parse(firstpnty) * yscale - uhoriz);

                            if (double.Parse(firstpnty) * yscale - uhoriz > maxy) {maxy = double.Parse(firstpnty) * yscale - uhoriz; }
                            var acPoly = new Polyline(kartacounter-1);
                            acPoly.Normal = Vector3d.ZAxis;
                            acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                            for (int lineinxl = 2; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                            {
                                 nextx = " "; xlserv2.UniRead(1, lineinxl, 1, ref nextx);//sheet row col
                                 nextxplus = " "; xlserv2.UniRead(1, lineinxl, 2, ref nextxplus);//sheet row col
                                nexty = " "; xlserv2.UniRead(1, lineinxl, 3, ref nexty);//sheet row col
                                acPoly.AddVertexAt(lineinxl - 1, new Point2d((double.Parse(nextx) * 100 + double.Parse(nextxplus))*xscale-xbegin, double.Parse(nexty) * yscale - uhoriz), 0, -1, -1);
                                if (double.Parse(nexty) * yscale - uhoriz > maxy) { maxy = double.Parse(nexty) * yscale - uhoriz; }
                            }
                            acPoly.Closed = false;

                            acBlkTblRec.AppendEntity(acPoly);
                            acTrans.AddNewlyCreatedDBObject(acPoly, true);
                            acTrans.Commit();
                        }
                    */
            #endregion drawline_no_scaling
            #region horiz load
            //#region #endregion horiz load
            double[] lineforscalingpikets = new double[5000];
            double[] lineforscalingplusos = new double[5000];
            double[] lineforscalingy = new double[5000];
            double[] lineforscalingx = new double[5000];
            string[] ordinatapart1 = new string[5000];
            string[] ordinatapart2 = new string[5000];
            string[] ordinatabothparts = new string[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                string cellcontentpk = " "; xlserv2.UniRead(1, lineinxl, 1, ref cellcontentpk);
                lineforscalingpikets[lineinxl] = double.Parse(cellcontentpk);
                string cellcontentps = " "; xlserv2.UniRead(1, lineinxl, 2, ref cellcontentps);
                lineforscalingplusos[lineinxl] = double.Parse(cellcontentps);
                string cellcontenty = " "; xlserv2.UniRead(1, lineinxl, 3, ref cellcontenty);
                lineforscalingy[lineinxl] = double.Parse(cellcontenty);
                lineforscalingx[lineinxl] = lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl];

                string ordinatapart1content = " "; xlserv2.UniRead(1, lineinxl, 4, ref ordinatapart1content);
                ordinatapart1[lineinxl] = ordinatapart1content;
                string ordinatapart2content = " "; xlserv2.UniRead(1, lineinxl, 5, ref ordinatapart2content);
                ordinatapart2[lineinxl] = ordinatapart2content;
                ordinatabothparts[lineinxl] = ordinatapart1[lineinxl] + ordinatapart2[lineinxl];

                if (double.Parse(cellcontenty) > maximgeoline) { maximgeoline = double.Parse(cellcontenty); };
            }
            #endregion horiz load

            #region   ися9 как по базе с скв
            //paragraph.ToLower(culture).Contains(word.ToLower(culture)) with CultureInfo.InvariantCulture
            List<string> allowedskwnameslist = new List<string>(999);
            List<double> allowedskw_piket = new List<double>(999);
            List<double> allowedskw_plus = new List<double>(999);
            List<int> allowedskwnumlist = new List<int>(999);
            //if NEED proverka IF mode exeptional skwas not ALL swas as usual
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if (Regex.IsMatch(ordinatabothparts[lineinxl], "скв", RegexOptions.IgnoreCase))
                {
                    allowedskwnameslist.Add(ordinatabothparts[lineinxl]);
                    //allowedskwnameslist.Add("sposob 1 reg");
                    allowedskw_piket.Add(lineforscalingpikets[lineinxl]);
                    allowedskw_plus.Add(lineforscalingplusos[lineinxl]);
                }
            }

            TextWriter nextoutfile = new StreamWriter("C:\\AtlPProf\\rezultat_proverki.txt", true, Encoding.Default);
            nextoutfile.WriteLine(" _ _ _ ");

            int ap5counter = 0;
            foreach (string value in allowedskwnameslist)
            {
                nextoutfile.WriteLine(value);
                Regex regex = new Regex(@"\d+");
                Match match = regex.Match(value);
                if (match.Success)
                {
                    nextoutfile.WriteLine(match.Value + "only digit");
                    allowedskwnumlist.Add(int.Parse(match.Value));
                }
                nextoutfile.WriteLine(match.Value + " piket " + allowedskw_piket[ap5counter] + " plus " + allowedskw_plus[ap5counter]);
                ap5counter++;

                Match match2 = Regex.Match(value.ToLower(), @"скв([.]+)$", RegexOptions.IgnoreCase);
                if (match2.Success)
                {
                    // Finally, we get the Group value and display it.
                    string key = match2.Groups[1].Value;
                    nextoutfile.WriteLine(key);
                }
            };
            nextoutfile.WriteLine(" _ _ _ ");
            nextoutfile.Close();

            /*
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match("Dot 55 Perls");
            if (match.Success)
            {
                Console.WriteLine(match.Value);
            }
            */
            #endregion

            #endregion horizline
            xlserv2.Close();

            ComplexObjects.ExcelServ xlserv = new ComplexObjects.ExcelServ();
            #region open xl
            var ofd =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );
            var dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK)
                return;
            // Display the name of the file and the contained sheets
            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd.Filename
            );
            #endregion open xl
            xlserv.Open(ofd.Filename);
            #region kolonki
            int goodskwacount = 0;
            int xlstartline = 5;//todo configit starting line
            int colforstartbur = 15;
            int colforstopbur = 16;
            int colforabsustbur = 17;//Q absolute height of hole
            int colforpodos = 6;
            int colforopisanie = 8;
            int colforvozrast = 18;
            int colforshtrih = 19;
            int colforkonsist = 20;
            int colforige = 11;
            int colforobraznar = 9;
            int colforobrazmon = 10;
            int colforobrazwod = 14;
            int colforupw = 12;
            int colforuuw = 13;
            int debuglevel = 0;
            int colforpiketplus = 3;
            int colforpiket = 2;
            double dbleb = 0.0;
            //dbleb = getopissize();
            double xbegin = 0.0;
            //xbegin = getxbegin();

            //isa18 ap18 task dont ask too many questions
            string egesizetxt = "2.5";
            double egesize = 2.5;
            //double replace221= 0.0;
            egesizetxt = readconfig(@"egesize = ");
            if ( (egesizetxt!="2")&&(egesizetxt!="2.5")&&(egesizetxt!="3")&&(egesizetxt!="3.5")&&(egesizetxt!="4")&&(egesizetxt!="4.5")&&(egesizetxt!="5")&&(egesizetxt!="6")){
                egesize = Getegesize();
            }
            else
            {
                egesize = double.Parse(egesizetxt);
            }
            double textsize = egesize;
            double deepestholepixels = 999;

            skwazya[] skwazi = new skwazya[1000];
            #region узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            for (int maxxllines = xlstartline; maxxllines < 4000; maxxllines++) //todo configit or autodetect
            {
                double picketallowed = 0;
                double plusallowed = 0;
                // ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка
                #region ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка

                string cellcontent = " "; xlserv.UniRead(1, maxxllines, 1, ref cellcontent);//sheet row col
                //04062016 ap6 check if skvanomer in allow list
                bool allowtofill = false;
                if ((cellcontent != "") && (cellcontent != null))
                {
                    allowtofill = true;

                    int numtosearch = 0;

                    Regex regex = new Regex(@"\d+");
                    Match match = regex.Match(cellcontent);
                    if (match.Success)
                    {
                        numtosearch = int.Parse(match.Value);
                    }
                    else
                    {
                        //Notification   SKWA bez numera
                    }

                    if (allowedskwnumlist.Contains(numtosearch))
                    {
                        allowtofill = true;
                        picketallowed = allowedskw_piket.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                        plusallowed = allowedskw_plus.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                    }
                    else
                    {
                        allowtofill = false;
                    }
                }
                #endregion

                if (allowtofill) //found NEW SKWA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                {
                    #region read first line
                    int maxgr = 0;
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    skwazi[goodskwacount] = new skwazya(goodskwacount);
                    skwazi[goodskwacount].name = cellcontent;
                    skwazi[goodskwacount].replace221 = 0.0;
                    string startbur = " "; xlserv.UniRead(1, maxxllines, colforstartbur, ref startbur); if ((startbur != "") && (startbur != null)) { skwazi[goodskwacount].start = startbur; }
                    string stopbur = " "; xlserv.UniRead(1, maxxllines, colforstopbur, ref stopbur); if ((stopbur != "") && (stopbur != null)) { skwazi[goodskwacount].stop = stopbur; }
                    string absustbur = " "; xlserv.UniRead(1, maxxllines, colforabsustbur, ref absustbur); if ((absustbur != "") && (absustbur != null))
                    {
                        skwazi[goodskwacount].absust = absustbur;
                        skwazi[goodskwacount].absust = Regex.Replace(skwazi[goodskwacount].absust, ",", ".");
                    }

                    /*
                    string piket = " "; xlserv.UniRead(1, maxxllines, colforpiket, ref piket); if ((piket != "") && (piket != null)) { skwazi[goodskwacount].piket = piket; } //Q column
                    string piketplus = " "; xlserv.UniRead(1, maxxllines, colforpiketplus, ref piketplus); if ((piketplus != "") && (piketplus != null))
                    {
                        skwazi[goodskwacount].piketplus = piketplus;
                        skwazi[goodskwacount].piketplus = Regex.Replace(skwazi[goodskwacount].piketplus, ",", ".");
                    }
                    */

                    skwazi[goodskwacount].skwax = (picketallowed * 100 + plusallowed) * xscale - xbegin;//todo start of left position at 0 even if starting piket is not 0/but 44meters
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    #endregion read first line

                    #region count Grunts IN SKWA
                    int gruntsinskwa = 0;
                    for (gruntsinskwa = 0; gruntsinskwa < 43; gruntsinskwa++)
                    {
                        string tempmaxlinesinskwa = "";
                        xlserv.UniRead(1, skwazi[goodskwacount].skwastartline + gruntsinskwa, colforopisanie, ref tempmaxlinesinskwa);
                        if ((tempmaxlinesinskwa == "") || (tempmaxlinesinskwa == null)) break;
                    }
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    #endregion count Grunts IN SKWA
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    skwazi[goodskwacount].layerscount = skwazi[goodskwacount].gruntsinskwa;
                    goodskwacount++;
                }

                picketallowed = 0;
                plusallowed = 0;
            }
            #endregion узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            #region дозаполнил скважины (а в планах = если есть только описания -и ничего кроме описаний, то нарисовать описания уже. а если описание и мощность, то еще лучше)
            for (int readmoreeveryskwa = 0; readmoreeveryskwa < goodskwacount; readmoreeveryskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[readmoreeveryskwa].gruntsinskwa; gruntinskwa++)
                {
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                    skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva = Regex.Replace(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva, ",", ".");
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforopisanie, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].opisanie);
                    //xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos + 1, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh);
                    if (gruntinskwa == 0) { skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva; }
                    else
                    {
                        // skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva - skwazi[readmoreeveryskwa].gruntiki[gruntinskwa-1].podoshva; }
                        //A__________                                               =  B______                                                  -   C____________
                        double B = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                        double C = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa - 1].podoshva);
                        double A = B - C;
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = A.ToString();
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobraznar, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obraznar);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobrazmon, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obrazmon);

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforige, ref ige);
                    if ((ige != "") && (ige != null)) skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].ige = ige;

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforvozrast, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].vozrast);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforshtrih, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforkonsist, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforupw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].upw);
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforuuw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].uuw);
                }
                if (debuglevel > 1) MgdAcApplication.ShowAlertDialog("из екселя прочитана скважина  : " + readmoreeveryskwa + "по счету от 0. например 0,1,2");
            }
            #endregion

            #region calc all skwa for totalmosh
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                double temptotalmosh = 0.0;
                int templaycoutmax = skwazi[writeskvcnt].layerscount;
                for (int s = 0; s < templaycoutmax; s++)
                {
                    double unparce = 0.0;
                    string tempsting = skwazi[writeskvcnt].gruntiki[s].tolshmosh;
                    double.TryParse(tempsting, out unparce);
                    temptotalmosh = temptotalmosh + unparce;
                }
                //double totalmosh = skwazi[goodskwacount].gettotalmosh();
                skwazi[writeskvcnt].totalmosh = temptotalmosh;
            }
            #endregion calc all skwa for totalmosh
            #region Check all values
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                if ((skwazi[writeskvcnt].absust == "") || (skwazi[writeskvcnt].absust == null)) skwazi[writeskvcnt].absust = "0";
                skwazi[writeskvcnt].absdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) - skwazi[writeskvcnt].totalmosh;
                skwazi[writeskvcnt].visualdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) * yscale - skwazi[writeskvcnt].totalmosh * geoscale;

                if (deepesthole > skwazi[writeskvcnt].absdeepestpnt) { deepesthole = skwazi[writeskvcnt].absdeepestpnt; };
                if (deepestholepixels > skwazi[writeskvcnt].visualdeepestpnt) { deepestholepixels = skwazi[writeskvcnt].visualdeepestpnt; };
            }
            deepestholepixels = deepestholepixels + zapas;
            #endregion Check all values

            if (autohor == "yes") { uhorizm = deepestholepixels / yscale; uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ }//se29 todo print uhorizm
            else { uhorizm = getuhorizm(); uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ };

            #region calc FLOOR and MID levels
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region calc floor
                skwazi[writeskvcnt].replace221 = (double.Parse(skwazi[writeskvcnt].absust) - uhorizm) * yscale;
                double tempnextceil = skwazi[writeskvcnt].replace221;
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    skwazi[writeskvcnt].gruntiki[readgr].konsiceiling = tempnextceil;
                    skwazi[writeskvcnt].gruntiki[readgr].konsifloor = tempnextceil - geoscale * double.Parse(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh);//10 - > geoscale
                    skwazi[writeskvcnt].gruntiki[readgr].circen = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2;
                    tempnextceil = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                }
                #endregion calc floor
            }
            #endregion draw shapki calc FLOOR and MID levels

            #region drawall
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region text ige and other
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }

                #endregion
                #region REZ shtrih - podoshva text - REZ
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    //left part of skwa
                    //ися10 без штриховок - другии версии командnohole(skwazi[writeskvcnt].skwax - 5.25 - varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);//left part of skwa

                    //text(skwazi[writeskvcnt].skwax + 17, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax + 4, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur));

                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);
                    //text(skwazi[writeskvcnt].skwax - 17, floorbydata, (tempabsuseverylayer  - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax - 12, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        //ися10 без штриховок - другии версии командhole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                    }
                    //ися10 без штриховок - другии версии командelse nohole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                }
                #endregion
                #region konsist
                //konsist
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    string temp = skwazi[writeskvcnt].gruntiki[readgr].konsist;
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((temp == "") || (temp == null))
                    {
                        varemptykonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, varpipe);
                        continue;
                    }
                    varnoholekonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].konsist, varpipe);
                }
                #endregion
                #region podval isa-1 ися12   переделка профиля в разрез
                text((skwazi[writeskvcnt].skwax), (-9), skwazi[writeskvcnt].name);

                string absusttext = skwazi[writeskvcnt].absust;
                absusttext = absusttext.Replace(",", ".");
                double tempabsus = double.Parse(absusttext);
                text((skwazi[writeskvcnt].skwax), (-19), tempabsus.ToString(podvalotmetkaustaccur));

                text((skwazi[writeskvcnt].skwax), (-29), skwazi[writeskvcnt].totalmosh.ToString(podvaldeepaccur));

                //fe9 2016 task to do sort before kalkulating distances between/ because holes are at random in source file 
                /* 
                if (writeskvcnt > 0)
                {
                    text((skwazi[writeskvcnt - 1].skwax + (skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / 2), (-39), ((skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwazi[writeskvcnt].skwax / 2), (-39), (skwazi[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }
*/
                text((skwazi[writeskvcnt].skwax), (-49), skwazi[writeskvcnt].stop);

                palkaatfloor5floorsrazrez(4, skwazi[writeskvcnt].skwax);
                //palkaatfloor5(4, skwazi[writeskvcnt].skwax);
                #endregion podval
            }
            #endregion drawall

            #region   isa-2 draw sorted dist at podval - we can use this sorted as main but it is bad for debugging - all skwaz mixed as the rezult and we have fk if input bad
            skwazya[] skwagood = new skwazya[goodskwacount];
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                skwagood[writeskvcnt] = skwazi[writeskvcnt];
            }

            skwazya[] skwasortedx = sortskwa(skwagood);

            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region podval

                if (writeskvcnt > 0)
                {
                    text((skwasortedx[writeskvcnt - 1].skwax + (skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / 2), (-39), ((skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwasortedx[writeskvcnt].skwax / 2), (-39), (skwasortedx[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }

                #endregion podval
            }

            #endregion draw sorted dist at podval


            #region XL Read samples of every HOLE of every LAYER
            for (int curskwa = 0; curskwa < goodskwacount; curskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string monstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazmon, ref monstring);
                    monstring = Regex.Replace(monstring, @"\t|\n|\r", ";");
                    string[] split = monstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "mon");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string narstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobraznar, ref narstring);
                    narstring = Regex.Replace(narstring, @"\t|\n|\r", ";");
                    string[] split = narstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "nar");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "nar");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string wodstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazwod, ref wodstring);
                    wodstring = Regex.Replace(wodstring, @"\t|\n|\r", ";");
                    string[] split = wodstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "wod");
                        }
                    }
                }
            }
            #endregion XL Read samples of every HOLE of every LAYER
            #region sort    samples //todo sort (take into consideration level)
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {
                skwazi[writeskvcnt].obriki = sortobriki(skwazi[writeskvcnt].obriki);
                // skwazi[writeskvcnt].obriki = pushl1tol2(skwazi[writeskvcnt].obriki);
                for (int curproba = 1; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    double diff = 0.0; diff = skwazi[writeskvcnt].obriki[curproba].bottom - skwazi[writeskvcnt].obriki[curproba - 1].bottom; if (diff < 0) { diff *= -1; }
                    if (diff < 1.9 / geoscale) skwazi[writeskvcnt].obriki[curproba].level = skwazi[writeskvcnt].obriki[curproba - 1].level + 1;//do se25 0.19
                }
            }
            #endregion
            #region drawall samples
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {                                                              // colonka by kolonka
                for (int curproba = 0; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "mon") probamon(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "nar") probanar(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "wod") probawod(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                }
            }
            #endregion drawall probs
            #region draw wodka
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                int ifanywaterhere = 0;

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if ((skwazi[writeskvcnt].gruntiki[readgr].uuw != "") && (skwazi[writeskvcnt].gruntiki[readgr].uuw != null))
                    {
                        double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw) * geoscale;
                        uuwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw).ToString("0.00"), skwazi[writeskvcnt].stop);
                        ifanywaterhere++;
                    };

                    if (skwazi[writeskvcnt].gruntiki[readgr].uuw != skwazi[writeskvcnt].gruntiki[readgr].upw)
                    {
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            upwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw).ToString("0.00"), skwazi[writeskvcnt].start);
                            ifanywaterhere++;
                        };
                    }
                }

                /*if (ifanywaterhere == 0)
                {
                    double viziblebottomy = skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor;
                    verttext((210 * writeskvcnt + 175), (222 + viziblebottomy) / 2 - 3, "Воды");
                    verttext((210 * writeskvcnt + 178), (222 + viziblebottomy) / 2 - 2, "нет");
                }*/
            }
            #endregion drawall
            #region task wodavpeske
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("PESOK") || skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("pesok"))
                    {
                        //search uuw ??upw?  in ALL GRUNTS or only this Grunt layer?
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double watertop = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                            double waterbot = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            //ne budu //if ((double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) > skwazi[writeskvcnt].gruntiki[readgr].konsiceiling)&&(double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) < skwazi[writeskvcnt].gruntiki[readgr].konsifloor)){  };
                            //noholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID");
                            //noholekonsist(skwazi[writeskvcnt].skwax-2.35, thislevel, waterbot, "SOLID");
                            varnoholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID", varpipe);
                        }
                    }
                }
            }
            #endregion task wodavpeske
            #region task palochki pod kolonkoy
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                palkapodkol(skwazi[writeskvcnt].skwax, skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor);
            }
            #endregion task palochki pod kolonkoy
            #endregion kolonki
            xlserv.Close();

            #region Horiz prodolshzhenie (no loading)
            #region horiz scale
            //#region #endregion 
            double[] lineyscaled = new double[5000];
            double[] linexscaled = new double[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                linexscaled[lineinxl] = lineforscalingx[lineinxl] * xscale;
                lineyscaled[lineinxl] = lineforscalingy[lineinxl] * yscale;
            }
            #endregion horiz scale
            #region horiz scale again add USLOV HORIZ -conditional horizon
            //#region #endregion scale again
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                //lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhoriz * yscale;
                lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhorizm * yscale;
            }
            #endregion  scale again add USLOV HORIZ -conditional horizon
            #region horiz draw scaled
            //#region #endregion draw scaled
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var pl = new Polyline();
                for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                {
                    pl.AddVertexAt(lineinxl - 1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, 0, 0);//original here --> pl.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);last here --->pl.AddVertexAt(3, new Point2d(0, 10), 0, 0, 0); pl.Closed = true;
                }

                pl.Closed = false;
                acBlkTblRec.AppendEntity(pl);
                acTrans.AddNewlyCreatedDBObject(pl, true);
                acTrans.Commit();
            }
            #endregion draw scaled

            #region     task ordinata //ися1 нужны вертикальные подписи и вертикальные палочки на отметки с комментариями даже без комментов
            int isnewpicket = 0;

            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if ((ordinatapart2[lineinxl] != null) && (ordinatapart2[lineinxl] != "") && (ordinatapart2[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart2[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
                if ((ordinatapart1[lineinxl] != null) && (ordinatapart1[lineinxl] != "") && (ordinatapart1[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart1[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

                //ися3 поправки ися1 теперь нужны все верт отметки даже без комментов
                //ися4 if modeconfigordinata >0 than

                if (lineinxl > 0)
                { //try isa6 
                    if (lineforscalingpikets[lineinxl] != lineforscalingpikets[lineinxl - 1]) isnewpicket = 1;
                    else isnewpicket = 0;
                }

                //ися12 тут в профилях каждая отметка рисуется и пикеты желтеньким
            }
            #endregion  task ordinata


            #region boxunderground
            double firstx = linexscaled[1];
            double firsty = lineyscaled[1];
            double lastx = linexscaled[kartacounter - 1];
            double lasty = lineyscaled[kartacounter - 1];

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(firstx, firsty), 0, -1, -1);  //do ok27 acPoly.AddVertexAt(0, new Point2d(firstx, firsty)   , 0, -1, -1); 
                acPoly.AddVertexAt(1, new Point2d(0, 0), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, lasty), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion boxunderground
            int podvalcount = 1;
            int otstup = -65;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Номеp скважины", otstup+2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 2;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметка устья, м", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 3;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Глубина, м", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 4;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояние, м", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 5;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Дата пpоходки", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval

            string scalextext = "Масштаб по оси X 1 : " + xdevidor;
            lefttext(scalextext, -65, 12, 2);
            string scaleytext = "Масштаб по оси Y 1 : " + ydevidor;
            lefttext(scaleytext, -65, 7.5, 2);
            string geoscaletext = "Масштаб геол. 1 : " + geodevidor;
            lefttext(geoscaletext, -65, 3, 2);

            #endregion
            lineikaok26(maximgeoline, uhorizm, yscale);
        }
        [CommandMethod("razrez_a_sboku_shtrihovka")]
        public void profvrazrzbok()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region ask some
            //dynamic acadObject = MgdAcApplication.AcadApplication;//zoom
            //acadObject.ZoomExtents();//zoom
            double yscale = 10;//calculated in this way 1000/100= 10
            double xscale = 2; //calculated in this way 1000/500= 2
            double uhorizm = 60;
            string autohor = "";

            autohor = readconfig(@"autohor = ");//isa18 ap18 task dont ask too many questions
            if ((autohor != "no") && (autohor != "yes"))
            {
                autohor = getautohor();
            } 

            if (autohor == "no") { uhorizm = getuhorizm(); };
            double zapas = 0;
            zapas = readzapas();
            string podvaldistbetweenaccur = readpodvaldistbetweenaccur();
            string lefttoholeaccur = readlefttoholeaccur();
            string righttoholeaccur = readrighttoholeaccur();
            string podvalotmetkaustaccur = readpodvalotmetkaustaccur();
            string podvaldeepaccur = readpodvaldeepaccur();
            double varpipe = readdiamskwa(); varpipe = varpipe / 2;
            double uhoriz = uhorizm * 10;
            double deepesthole = 999;

            double xdevidor = getxscale();
            xscale = 1000 / xdevidor;

            double ydevidor = getyscale();
            yscale = 1000 / ydevidor;

            double geoscale = 10;
            double geodevidor = getgeoscale();
            geoscale = 1000 / geodevidor;
            double maximgeoline = -1500;
            #endregion ask some

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);
            #region horizline LOAD
            #region detect xl maximum -----> kartacounter

            string light = "green";
            int kartacounter = 1;
            for (; kartacounter < 5000; kartacounter++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(1, kartacounter, 1, ref cellcontent);//sheet row col
                if (((cellcontent == "") || (cellcontent == null)) && (kartacounter == 1))
                {
                    light = "red"; break; ;
                }
                if ((cellcontent == "") || (cellcontent == null))
                {
                    break; ;
                }
            }
            //kartacounter is maximum lines in exclell
            #endregion detect xl maximum -----> kartacounter
            #region drawline_no_scaling
            double maxyininput = 0;
            double maxy = 0;

            string firstpntx = "0";
            string firstpntxplus = "0";
            string firstpnty = "0";

            string nextx = "0";
            string nextxplus = "0";
            string nexty = "0";
            /*
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            firstpntx = " ";
                            xlserv2.UniRead(1, 1, 1, ref firstpntx);//sheet row col
                             firstpntxplus = " "; xlserv2.UniRead(1, 1, 2, ref firstpntxplus);//sheet row col
                            firstpnty = " ";
                            xlserv2.UniRead(1, 1, 3, ref firstpnty);//sheet row col

                            var acPoint = new Point2d((double.Parse(firstpntx) * 100 + double.Parse(firstpntxplus))*xscale-xbegin, double.Parse(firstpnty) * yscale - uhoriz);

                            if (double.Parse(firstpnty) * yscale - uhoriz > maxy) {maxy = double.Parse(firstpnty) * yscale - uhoriz; }
                            var acPoly = new Polyline(kartacounter-1);
                            acPoly.Normal = Vector3d.ZAxis;
                            acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                            for (int lineinxl = 2; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                            {
                                 nextx = " "; xlserv2.UniRead(1, lineinxl, 1, ref nextx);//sheet row col
                                 nextxplus = " "; xlserv2.UniRead(1, lineinxl, 2, ref nextxplus);//sheet row col
                                nexty = " "; xlserv2.UniRead(1, lineinxl, 3, ref nexty);//sheet row col
                                acPoly.AddVertexAt(lineinxl - 1, new Point2d((double.Parse(nextx) * 100 + double.Parse(nextxplus))*xscale-xbegin, double.Parse(nexty) * yscale - uhoriz), 0, -1, -1);
                                if (double.Parse(nexty) * yscale - uhoriz > maxy) { maxy = double.Parse(nexty) * yscale - uhoriz; }
                            }
                            acPoly.Closed = false;

                            acBlkTblRec.AppendEntity(acPoly);
                            acTrans.AddNewlyCreatedDBObject(acPoly, true);
                            acTrans.Commit();
                        }
                    */
            #endregion drawline_no_scaling
            #region horiz load
            //#region #endregion horiz load
            double[] lineforscalingpikets = new double[5000];
            double[] lineforscalingplusos = new double[5000];
            double[] lineforscalingy = new double[5000];
            double[] lineforscalingx = new double[5000];
            string[] ordinatapart1 = new string[5000];
            string[] ordinatapart2 = new string[5000];
            string[] ordinatabothparts = new string[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                string cellcontentpk = " "; xlserv2.UniRead(1, lineinxl, 1, ref cellcontentpk);
                lineforscalingpikets[lineinxl] = double.Parse(cellcontentpk);
                string cellcontentps = " "; xlserv2.UniRead(1, lineinxl, 2, ref cellcontentps);
                lineforscalingplusos[lineinxl] = double.Parse(cellcontentps);
                string cellcontenty = " "; xlserv2.UniRead(1, lineinxl, 3, ref cellcontenty);
                lineforscalingy[lineinxl] = double.Parse(cellcontenty);
                lineforscalingx[lineinxl] = lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl];

                string ordinatapart1content = " "; xlserv2.UniRead(1, lineinxl, 4, ref ordinatapart1content);
                ordinatapart1[lineinxl] = ordinatapart1content;
                string ordinatapart2content = " "; xlserv2.UniRead(1, lineinxl, 5, ref ordinatapart2content);
                ordinatapart2[lineinxl] = ordinatapart2content;
                ordinatabothparts[lineinxl] = ordinatapart1[lineinxl] + ordinatapart2[lineinxl];

                if (double.Parse(cellcontenty) > maximgeoline) { maximgeoline = double.Parse(cellcontenty); };
            }
            #endregion horiz load

            #region   ися9 как по базе с скв
            //paragraph.ToLower(culture).Contains(word.ToLower(culture)) with CultureInfo.InvariantCulture
            List<string> allowedskwnameslist = new List<string>(999);
            List<double> allowedskw_piket = new List<double>(999);
            List<double> allowedskw_plus = new List<double>(999);
            List<int> allowedskwnumlist = new List<int>(999);
            //if NEED proverka IF mode exeptional skwas not ALL swas as usual
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if (Regex.IsMatch(ordinatabothparts[lineinxl], "скв", RegexOptions.IgnoreCase))
                {
                    allowedskwnameslist.Add(ordinatabothparts[lineinxl]);
                    //allowedskwnameslist.Add("sposob 1 reg");
                    allowedskw_piket.Add(lineforscalingpikets[lineinxl]);
                    allowedskw_plus.Add(lineforscalingplusos[lineinxl]);
                }
            }

            TextWriter nextoutfile = new StreamWriter("c:\\AtlPProf\\rezultat_proverki.txt", true, Encoding.Default);
            nextoutfile.WriteLine(" _ _ _ ");

            int ap5counter = 0;
            foreach (string value in allowedskwnameslist)
            {
                nextoutfile.WriteLine(value);
                Regex regex = new Regex(@"\d+");
                Match match = regex.Match(value);
                if (match.Success)
                {
                    nextoutfile.WriteLine(match.Value + "only digit");
                    allowedskwnumlist.Add(int.Parse(match.Value));
                }
                nextoutfile.WriteLine(match.Value + " piket " + allowedskw_piket[ap5counter] + " plus " + allowedskw_plus[ap5counter]);
                ap5counter++;

                Match match2 = Regex.Match(value.ToLower(), @"скв([.]+)$", RegexOptions.IgnoreCase);
                if (match2.Success)
                {
                    // Finally, we get the Group value and display it.
                    string key = match2.Groups[1].Value;
                    nextoutfile.WriteLine(key);
                }
            };
            nextoutfile.WriteLine(" _ _ _ ");
            nextoutfile.Close();

            /*
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match("Dot 55 Perls");
            if (match.Success)
            {
                Console.WriteLine(match.Value);
            }
            */
            #endregion

            #endregion horizline
            xlserv2.Close();

            ComplexObjects.ExcelServ xlserv = new ComplexObjects.ExcelServ();
            #region open xl
            var ofd =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );
            var dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK)
                return;
            // Display the name of the file and the contained sheets
            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd.Filename
            );
            #endregion open xl
            xlserv.Open(ofd.Filename);
            #region kolonki
            int goodskwacount = 0;
            int xlstartline = 5;//todo configit starting line
            int colforstartbur = 15;
            int colforstopbur = 16;
            int colforabsustbur = 17;//Q absolute height of hole
            int colforpodos = 6;
            int colforopisanie = 8;
            int colforvozrast = 18;
            int colforshtrih = 19;
            int colforkonsist = 20;
            int colforige = 11;
            int colforobraznar = 9;
            int colforobrazmon = 10;
            int colforobrazwod = 14;
            int colforupw = 12;
            int colforuuw = 13;
            int debuglevel = 0;
            int colforpiketplus = 3;
            int colforpiket = 2;
            double dbleb = 0.0;
            //dbleb = getopissize();
            double xbegin = 0.0;
            //xbegin = getxbegin();

            //isa18 ap18 task dont ask too many questions
            string egesizetxt = "2.5";
            double egesize = 2.5;
            //double replace221= 0.0;
            egesizetxt = readconfig(@"egesize = ");
            if ((egesizetxt != "2") && (egesizetxt != "2.5") && (egesizetxt != "3") && (egesizetxt != "3.5") && (egesizetxt != "4") && (egesizetxt != "4.5") && (egesizetxt != "5") && (egesizetxt != "6"))
            {
                egesize = Getegesize();
            }
            else
            {
                egesize = double.Parse(egesizetxt);
            }
            double textsize = egesize;
            double deepestholepixels = 999;

            skwazya[] skwazi = new skwazya[1000];
            #region узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            for (int maxxllines = xlstartline; maxxllines < 5000; maxxllines++) //todo configit or autodetect
            {
                double picketallowed = 0;
                double plusallowed = 0;
                // ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка
                #region ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка

                string cellcontent = " "; xlserv.UniRead(1, maxxllines, 1, ref cellcontent);//sheet row col
                //04062016 ap6 check if skvanomer in allow list
                bool allowtofill = false;
                if ((cellcontent != "") && (cellcontent != null))
                {
                    allowtofill = true;

                    int numtosearch = 0;

                    Regex regex = new Regex(@"\d+");
                    Match match = regex.Match(cellcontent);
                    if (match.Success)
                    {
                        numtosearch = int.Parse(match.Value);
                    }
                    else
                    {
                        //Notification   SKWA bez numera
                    }

                    if (allowedskwnumlist.Contains(numtosearch))
                    {
                        allowtofill = true;
                        picketallowed = allowedskw_piket.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                        plusallowed = allowedskw_plus.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                    }
                    else
                    {
                        allowtofill = false;
                    }
                }
                #endregion

                if (allowtofill) //found NEW SKWA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                {
                    #region read first line
                    int maxgr = 0;
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    skwazi[goodskwacount] = new skwazya(goodskwacount);
                    skwazi[goodskwacount].name = cellcontent;
                    skwazi[goodskwacount].replace221 = 0.0;
                    string startbur = " "; xlserv.UniRead(1, maxxllines, colforstartbur, ref startbur); if ((startbur != "") && (startbur != null)) { skwazi[goodskwacount].start = startbur; }
                    string stopbur = " "; xlserv.UniRead(1, maxxllines, colforstopbur, ref stopbur); if ((stopbur != "") && (stopbur != null)) { skwazi[goodskwacount].stop = stopbur; }
                    string absustbur = " "; xlserv.UniRead(1, maxxllines, colforabsustbur, ref absustbur); if ((absustbur != "") && (absustbur != null))
                    {
                        skwazi[goodskwacount].absust = absustbur;
                        skwazi[goodskwacount].absust = Regex.Replace(skwazi[goodskwacount].absust, ",", ".");
                    }

                    /*
                    string piket = " "; xlserv.UniRead(1, maxxllines, colforpiket, ref piket); if ((piket != "") && (piket != null)) { skwazi[goodskwacount].piket = piket; } //Q column
                    string piketplus = " "; xlserv.UniRead(1, maxxllines, colforpiketplus, ref piketplus); if ((piketplus != "") && (piketplus != null))
                    {
                        skwazi[goodskwacount].piketplus = piketplus;
                        skwazi[goodskwacount].piketplus = Regex.Replace(skwazi[goodskwacount].piketplus, ",", ".");
                    }
                    */

                    skwazi[goodskwacount].skwax = (picketallowed * 100 + plusallowed) * xscale - xbegin;//todo start of left position at 0 even if starting piket is not 0/but 44meters
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    #endregion read first line

                    #region count Grunts IN SKWA
                    int gruntsinskwa = 0;
                    for (gruntsinskwa = 0; gruntsinskwa < 43; gruntsinskwa++)
                    {
                        string tempmaxlinesinskwa = "";
                        xlserv.UniRead(1, skwazi[goodskwacount].skwastartline + gruntsinskwa, colforopisanie, ref tempmaxlinesinskwa);
                        if ((tempmaxlinesinskwa == "") || (tempmaxlinesinskwa == null)) break;
                    }
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    #endregion count Grunts IN SKWA
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    skwazi[goodskwacount].layerscount = skwazi[goodskwacount].gruntsinskwa;
                    goodskwacount++;
                }

                picketallowed = 0;
                plusallowed = 0;
            }
            #endregion узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            #region дозаполнил скважины (а в планах = если есть только описания -и ничего кроме описаний, то нарисовать описания уже. а если описание и мощность, то еще лучше)
            for (int readmoreeveryskwa = 0; readmoreeveryskwa < goodskwacount; readmoreeveryskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[readmoreeveryskwa].gruntsinskwa; gruntinskwa++)
                {
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                    skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva = Regex.Replace(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva, ",", ".");
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforopisanie, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].opisanie);
                    //xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos + 1, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh);
                    if (gruntinskwa == 0) { skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva; }
                    else
                    {
                        // skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva - skwazi[readmoreeveryskwa].gruntiki[gruntinskwa-1].podoshva; }
                        //A__________                                               =  B______                                                  -   C____________
                        double B = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                        double C = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa - 1].podoshva);
                        double A = B - C;
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = A.ToString();
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobraznar, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obraznar);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobrazmon, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obrazmon);

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforige, ref ige);
                    if ((ige != "") && (ige != null)) skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].ige = ige;

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforvozrast, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].vozrast);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforshtrih, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforkonsist, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforupw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].upw);
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforuuw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].uuw);
                }
                if (debuglevel > 1) MgdAcApplication.ShowAlertDialog("из екселя прочитана скважина  : " + readmoreeveryskwa + "по счету от 0. например 0,1,2");
            }
            #endregion

            #region calc all skwa for totalmosh
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                double temptotalmosh = 0.0;
                int templaycoutmax = skwazi[writeskvcnt].layerscount;
                for (int s = 0; s < templaycoutmax; s++)
                {
                    double unparce = 0.0;
                    string tempsting = skwazi[writeskvcnt].gruntiki[s].tolshmosh;
                    double.TryParse(tempsting, out unparce);
                    temptotalmosh = temptotalmosh + unparce;
                }
                //double totalmosh = skwazi[goodskwacount].gettotalmosh();
                skwazi[writeskvcnt].totalmosh = temptotalmosh;
            }
            #endregion calc all skwa for totalmosh
            #region Check all values
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                if ((skwazi[writeskvcnt].absust == "") || (skwazi[writeskvcnt].absust == null)) skwazi[writeskvcnt].absust = "0";
                skwazi[writeskvcnt].absdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) - skwazi[writeskvcnt].totalmosh;
                skwazi[writeskvcnt].visualdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) * yscale - skwazi[writeskvcnt].totalmosh * geoscale;

                if (deepesthole > skwazi[writeskvcnt].absdeepestpnt) { deepesthole = skwazi[writeskvcnt].absdeepestpnt; };
                if (deepestholepixels > skwazi[writeskvcnt].visualdeepestpnt) { deepestholepixels = skwazi[writeskvcnt].visualdeepestpnt; };
            }
            deepestholepixels = deepestholepixels + zapas;
            #endregion Check all values

            if (autohor == "yes") { uhorizm = deepestholepixels / yscale; uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ }//se29 todo print uhorizm
            else { uhorizm = getuhorizm(); uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ };

            #region calc FLOOR and MID levels
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region calc floor
                skwazi[writeskvcnt].replace221 = (double.Parse(skwazi[writeskvcnt].absust) - uhorizm) * yscale;
                double tempnextceil = skwazi[writeskvcnt].replace221;
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    skwazi[writeskvcnt].gruntiki[readgr].konsiceiling = tempnextceil;
                    skwazi[writeskvcnt].gruntiki[readgr].konsifloor = tempnextceil - geoscale * double.Parse(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh);//10 - > geoscale
                    skwazi[writeskvcnt].gruntiki[readgr].circen = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2;
                    tempnextceil = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                }
                #endregion calc floor
            }
            #endregion draw shapki calc FLOOR and MID levels

            #region drawall
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region text ige and other
                /*for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);

                    undertext(skwazi[writeskvcnt].skwax, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur), varpipe);
                    leftgeotxt(skwazi[writeskvcnt].skwax - varpipe, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }*/

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }

                #endregion
                #region REZ shtrih - podoshva text - REZ
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    //left part of skwa
                    nohole(skwazi[writeskvcnt].skwax - 5.25 - varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);//left part of skwa

                    //text(skwazi[writeskvcnt].skwax + 17, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax + 4, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur));

                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);
                    ///text(skwazi[writeskvcnt].skwax - 17, floorbydata, (tempabsuseverylayer  - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax - 12, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        hole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                    }
                    else nohole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                }
                #endregion
                #region konsist
                //konsist
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    string temp = skwazi[writeskvcnt].gruntiki[readgr].konsist;
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((temp == "") || (temp == null))
                    {
                        varemptykonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, varpipe);
                        continue;
                    }
                    varnoholekonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].konsist, varpipe);
                }
                #endregion
                #region podval isa-1 ися12   переделка профиля в разрез
                text((skwazi[writeskvcnt].skwax), (-9), skwazi[writeskvcnt].name);

                string absusttext = skwazi[writeskvcnt].absust;
                absusttext = absusttext.Replace(",", ".");
                double tempabsus = double.Parse(absusttext);
                text((skwazi[writeskvcnt].skwax), (-19), tempabsus.ToString(podvalotmetkaustaccur));

                text((skwazi[writeskvcnt].skwax), (-29), skwazi[writeskvcnt].totalmosh.ToString(podvaldeepaccur));

                //fe9 2016 task to do sort before kalkulating distances between/ because holes are at random in source file 
                /* 
                if (writeskvcnt > 0)
                {
                    text((skwazi[writeskvcnt - 1].skwax + (skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / 2), (-39), ((skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwazi[writeskvcnt].skwax / 2), (-39), (skwazi[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }
*/
                text((skwazi[writeskvcnt].skwax), (-49), skwazi[writeskvcnt].stop);

                palkaatfloor5floorsrazrez(4, skwazi[writeskvcnt].skwax);
                //ися12   переделка профиля в разрез palkaatfloor(4, skwazi[writeskvcnt].skwax);
                #endregion podval
            }
            #endregion drawall

            #region   isa-2 draw sorted dist at podval - we can use this sorted as main but it is bad for debugging - all skwaz mixed as the rezult and we have fk if input bad
            skwazya[] skwagood = new skwazya[goodskwacount];
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                skwagood[writeskvcnt] = skwazi[writeskvcnt];
            }

            skwazya[] skwasortedx = sortskwa(skwagood);

            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region podval

                if (writeskvcnt > 0)
                {
                    text((skwasortedx[writeskvcnt - 1].skwax + (skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / 2), (-39), ((skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwasortedx[writeskvcnt].skwax / 2), (-39), (skwasortedx[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }

                #endregion podval
            }

            #endregion draw sorted dist at podval


            #region XL Read samples of every HOLE of every LAYER
            for (int curskwa = 0; curskwa < goodskwacount; curskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string monstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazmon, ref monstring);
                    monstring = Regex.Replace(monstring, @"\t|\n|\r", ";");
                    string[] split = monstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "mon");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string narstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobraznar, ref narstring);
                    narstring = Regex.Replace(narstring, @"\t|\n|\r", ";");
                    string[] split = narstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "nar");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "nar");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string wodstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazwod, ref wodstring);
                    wodstring = Regex.Replace(wodstring, @"\t|\n|\r", ";");
                    string[] split = wodstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "wod");
                        }
                    }
                }
            }
            #endregion XL Read samples of every HOLE of every LAYER
            #region sort    samples //todo sort (take into consideration level)
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {
                skwazi[writeskvcnt].obriki = sortobriki(skwazi[writeskvcnt].obriki);
                // skwazi[writeskvcnt].obriki = pushl1tol2(skwazi[writeskvcnt].obriki);
                for (int curproba = 1; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    double diff = 0.0; diff = skwazi[writeskvcnt].obriki[curproba].bottom - skwazi[writeskvcnt].obriki[curproba - 1].bottom; if (diff < 0) { diff *= -1; }
                    if (diff < 1.9 / geoscale) skwazi[writeskvcnt].obriki[curproba].level = skwazi[writeskvcnt].obriki[curproba - 1].level + 1;//do se25 0.19
                }
            }
            #endregion
            #region drawall samples
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {                                                              // colonka by kolonka
                for (int curproba = 0; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "mon") probamon(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "nar") probanar(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "wod") probawod(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                }
            }
            #endregion drawall probs
            #region draw wodka
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                int ifanywaterhere = 0;

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if ((skwazi[writeskvcnt].gruntiki[readgr].uuw != "") && (skwazi[writeskvcnt].gruntiki[readgr].uuw != null))
                    {
                        double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw) * geoscale;
                        uuwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw).ToString("0.00"), skwazi[writeskvcnt].stop);
                        ifanywaterhere++;
                    };

                    if (skwazi[writeskvcnt].gruntiki[readgr].uuw != skwazi[writeskvcnt].gruntiki[readgr].upw)
                    {
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            upwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw).ToString("0.00"), skwazi[writeskvcnt].start);
                            ifanywaterhere++;
                        };
                    }
                }

                /*if (ifanywaterhere == 0)
                {
                    double viziblebottomy = skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor;
                    verttext((210 * writeskvcnt + 175), (222 + viziblebottomy) / 2 - 3, "Воды");
                    verttext((210 * writeskvcnt + 178), (222 + viziblebottomy) / 2 - 2, "нет");
                }*/
            }
            #endregion drawall
            #region task wodavpeske
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("PESOK") || skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("pesok"))
                    {
                        //search uuw ??upw?  in ALL GRUNTS or only this Grunt layer?
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double watertop = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                            double waterbot = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            //ne budu //if ((double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) > skwazi[writeskvcnt].gruntiki[readgr].konsiceiling)&&(double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) < skwazi[writeskvcnt].gruntiki[readgr].konsifloor)){  };
                            //noholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID");
                            //noholekonsist(skwazi[writeskvcnt].skwax-2.35, thislevel, waterbot, "SOLID");
                            varnoholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID", varpipe);
                        }
                    }
                }
            }
            #endregion task wodavpeske
            #region task palochki pod kolonkoy
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                palkapodkol(skwazi[writeskvcnt].skwax, skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor);
            }
            #endregion task palochki pod kolonkoy
            #endregion kolonki
            xlserv.Close();

            #region Horiz prodolshzhenie (no loading)
            #region horiz scale
            //#region #endregion 
            double[] lineyscaled = new double[5000];
            double[] linexscaled = new double[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                linexscaled[lineinxl] = lineforscalingx[lineinxl] * xscale;
                lineyscaled[lineinxl] = lineforscalingy[lineinxl] * yscale;
            }
            #endregion horiz scale
            #region horiz scale again add USLOV HORIZ -conditional horizon
            //#region #endregion scale again
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                //lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhoriz * yscale;
                lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhorizm * yscale;
            }
            #endregion  scale again add USLOV HORIZ -conditional horizon
            #region horiz draw scaled
            //#region #endregion draw scaled
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var pl = new Polyline();
                for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                {
                    pl.AddVertexAt(lineinxl - 1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, 0, 0);//original here --> pl.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);last here --->pl.AddVertexAt(3, new Point2d(0, 10), 0, 0, 0); pl.Closed = true;
                }

                pl.Closed = false;
                acBlkTblRec.AppendEntity(pl);
                acTrans.AddNewlyCreatedDBObject(pl, true);
                acTrans.Commit();
            }
            #endregion draw scaled

            #region     task ordinata //ися12 переделка профиля в разрез = разница в подвале и туче отметок
            int isnewpicket = 0;

            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if ((ordinatapart2[lineinxl] != null) && (ordinatapart2[lineinxl] != "") && (ordinatapart2[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart2[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
                if ((ordinatapart1[lineinxl] != null) && (ordinatapart1[lineinxl] != "") && (ordinatapart1[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart1[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

                //udalenu tut no est v profile
                //ися3 поправки ися1 теперь нужны все верт отметки даже без комментов udalenu tut no est v profile
                //ися4 if modeconfigordinata >0 than
            }
            #endregion  task ordinata


            #region boxunderground
            double firstx = linexscaled[1];
            double firsty = lineyscaled[1];
            double lastx = linexscaled[kartacounter - 1];
            double lasty = lineyscaled[kartacounter - 1];

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(firstx, firsty), 0, -1, -1);  //do ok27 acPoly.AddVertexAt(0, new Point2d(firstx, firsty)   , 0, -1, -1); 
                acPoly.AddVertexAt(1, new Point2d(0, 0), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, lasty), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion boxunderground
            int podvalcount = 1;
            int otstup = -65;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Номеp скважины", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 2;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметка устья, м", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 3;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Глубина, м", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 4;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояние, м", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 5;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Дата пpоходки", otstup + 2, podvalcount * -10 - 1, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(otstup, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(otstup, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval

            string scalextext = "Масштаб по оси X 1 : " + xdevidor;
            lefttext(scalextext, -65, 12, 2);
            string scaleytext = "Масштаб по оси Y 1 : " + ydevidor;
            lefttext(scaleytext, -65, 7.5, 2);
            string geoscaletext = "Масштаб геол. 1 : " + geodevidor;
            lefttext(geoscaletext, -65, 3, 2);

            #endregion
            lineikaok26(maximgeoline, uhorizm, yscale);
        }

        [CommandMethod("profile_a")]
        public void profileasbokubezshtrihovok()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region ask some
            //dynamic acadObject = MgdAcApplication.AcadApplication;//zoom
            //acadObject.ZoomExtents();//zoom
            double yscale = 10;//calculated in this way 1000/100= 10
            double xscale = 2; //calculated in this way 1000/500= 2
            double uhorizm = 60;
            string autohor = "";
            autohor = readconfig(@"autohor = ");//isa18 ap18 task dont ask too many questions
            if ((autohor != "no") && (autohor != "yes"))
            {
                autohor = getautohor();
            } 

            if (autohor == "no") { uhorizm = getuhorizm(); };
            double zapas = 0;
            zapas = readzapas();
            string podvaldistbetweenaccur = readpodvaldistbetweenaccur();
            string lefttoholeaccur = readlefttoholeaccur();
            string righttoholeaccur = readrighttoholeaccur();
            string podvalotmetkaustaccur = readpodvalotmetkaustaccur();
            string podvaldeepaccur = readpodvaldeepaccur();
            double varpipe = readdiamskwa(); varpipe = varpipe / 2;
            double uhoriz = uhorizm * 10;
            double deepesthole = 999;

            double xdevidor = getxscale();
            xscale = 1000 / xdevidor;

            double ydevidor = getyscale();
            yscale = 1000 / ydevidor;

            double geoscale = 10;
            double geodevidor = getgeoscale();
            geoscale = 1000 / geodevidor;
            double maximgeoline = -1500;
            #endregion ask some

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);
            #region horizline LOAD
            #region detect xl maximum -----> kartacounter

            string light = "green";
            int kartacounter = 1;
            for (; kartacounter < 5000; kartacounter++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(1, kartacounter, 1, ref cellcontent);//sheet row col
                if (((cellcontent == "") || (cellcontent == null)) && (kartacounter == 1))
                {
                    light = "red"; break; ;
                }
                if ((cellcontent == "") || (cellcontent == null))
                {
                    break; ;
                }
            }
            //kartacounter is maximum lines in exclell
            #endregion detect xl maximum -----> kartacounter
            #region drawline_no_scaling
            double maxyininput = 0;
            double maxy = 0;

            string firstpntx = "0";
            string firstpntxplus = "0";
            string firstpnty = "0";

            string nextx = "0";
            string nextxplus = "0";
            string nexty = "0";
            /*
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            firstpntx = " ";
                            xlserv2.UniRead(1, 1, 1, ref firstpntx);//sheet row col
                             firstpntxplus = " "; xlserv2.UniRead(1, 1, 2, ref firstpntxplus);//sheet row col
                            firstpnty = " ";
                            xlserv2.UniRead(1, 1, 3, ref firstpnty);//sheet row col

                            var acPoint = new Point2d((double.Parse(firstpntx) * 100 + double.Parse(firstpntxplus))*xscale-xbegin, double.Parse(firstpnty) * yscale - uhoriz);

                            if (double.Parse(firstpnty) * yscale - uhoriz > maxy) {maxy = double.Parse(firstpnty) * yscale - uhoriz; }
                            var acPoly = new Polyline(kartacounter-1);
                            acPoly.Normal = Vector3d.ZAxis;
                            acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                            for (int lineinxl = 2; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                            {
                                 nextx = " "; xlserv2.UniRead(1, lineinxl, 1, ref nextx);//sheet row col
                                 nextxplus = " "; xlserv2.UniRead(1, lineinxl, 2, ref nextxplus);//sheet row col
                                nexty = " "; xlserv2.UniRead(1, lineinxl, 3, ref nexty);//sheet row col
                                acPoly.AddVertexAt(lineinxl - 1, new Point2d((double.Parse(nextx) * 100 + double.Parse(nextxplus))*xscale-xbegin, double.Parse(nexty) * yscale - uhoriz), 0, -1, -1);
                                if (double.Parse(nexty) * yscale - uhoriz > maxy) { maxy = double.Parse(nexty) * yscale - uhoriz; }
                            }
                            acPoly.Closed = false;

                            acBlkTblRec.AppendEntity(acPoly);
                            acTrans.AddNewlyCreatedDBObject(acPoly, true);
                            acTrans.Commit();
                        }
                    */
            #endregion drawline_no_scaling
            #region horiz load
            //#region #endregion horiz load
            double[] lineforscalingpikets = new double[5000];
            double[] lineforscalingplusos = new double[5000];
            double[] lineforscalingy = new double[5000];
            double[] lineforscalingx = new double[5000];
            string[] ordinatapart1 = new string[5000];
            string[] ordinatapart2 = new string[5000];
            string[] ordinatabothparts = new string[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                string cellcontentpk = " "; xlserv2.UniRead(1, lineinxl, 1, ref cellcontentpk);
                lineforscalingpikets[lineinxl] = double.Parse(cellcontentpk);
                string cellcontentps = " "; xlserv2.UniRead(1, lineinxl, 2, ref cellcontentps);
                lineforscalingplusos[lineinxl] = double.Parse(cellcontentps);
                string cellcontenty = " "; xlserv2.UniRead(1, lineinxl, 3, ref cellcontenty);
                lineforscalingy[lineinxl] = double.Parse(cellcontenty);
                lineforscalingx[lineinxl] = lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl];

                string ordinatapart1content = " "; xlserv2.UniRead(1, lineinxl, 4, ref ordinatapart1content);
                ordinatapart1[lineinxl] = ordinatapart1content;
                string ordinatapart2content = " "; xlserv2.UniRead(1, lineinxl, 5, ref ordinatapart2content);
                ordinatapart2[lineinxl] = ordinatapart2content;
                ordinatabothparts[lineinxl] = ordinatapart1[lineinxl] + ordinatapart2[lineinxl];

                if (double.Parse(cellcontenty) > maximgeoline) { maximgeoline = double.Parse(cellcontenty); };
            }
            #endregion horiz load

            #region   ися9 как по базе с скв
            //paragraph.ToLower(culture).Contains(word.ToLower(culture)) with CultureInfo.InvariantCulture
            List<string> allowedskwnameslist = new List<string>(999);
            List<double> allowedskw_piket = new List<double>(999);
            List<double> allowedskw_plus = new List<double>(999);
            List<int> allowedskwnumlist = new List<int>(999);
            //if NEED proverka IF mode exeptional skwas not ALL swas as usual
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if (Regex.IsMatch(ordinatabothparts[lineinxl], "скв", RegexOptions.IgnoreCase))
                {
                    allowedskwnameslist.Add(ordinatabothparts[lineinxl]);
                    //allowedskwnameslist.Add("sposob 1 reg");
                    allowedskw_piket.Add(lineforscalingpikets[lineinxl]);
                    allowedskw_plus.Add(lineforscalingplusos[lineinxl]);
                }
            }

            TextWriter nextoutfile = new StreamWriter("c:\\AtlPProf\\rezultat_proverki.txt", true, Encoding.Default);
            nextoutfile.WriteLine(" _ _ _ ");

            int ap5counter = 0;
            foreach (string value in allowedskwnameslist)
            {
                nextoutfile.WriteLine(value);
                Regex regex = new Regex(@"\d+");
                Match match = regex.Match(value);
                if (match.Success)
                {
                    nextoutfile.WriteLine(match.Value + "only digit");
                    allowedskwnumlist.Add(int.Parse(match.Value));
                }
                nextoutfile.WriteLine(match.Value + " piket " + allowedskw_piket[ap5counter] + " plus " + allowedskw_plus[ap5counter]);
                ap5counter++;

                Match match2 = Regex.Match(value.ToLower(), @"скв([.]+)$", RegexOptions.IgnoreCase);
                if (match2.Success)
                {
                    // Finally, we get the Group value and display it.
                    string key = match2.Groups[1].Value;
                    nextoutfile.WriteLine(key);
                }
            };
            nextoutfile.WriteLine(" _ _ _ ");
            nextoutfile.Close();

            /*
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match("Dot 55 Perls");
            if (match.Success)
            {
                Console.WriteLine(match.Value);
            }
            */
            #endregion

            #endregion horizline
            xlserv2.Close();

            ComplexObjects.ExcelServ xlserv = new ComplexObjects.ExcelServ();
            #region open xl
            var ofd =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );
            var dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK)
                return;
            // Display the name of the file and the contained sheets
            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd.Filename
            );
            #endregion open xl
            xlserv.Open(ofd.Filename);
            #region kolonki
            int goodskwacount = 0;
            int xlstartline = 5;//todo configit starting line
            int colforstartbur = 15;
            int colforstopbur = 16;
            int colforabsustbur = 17;//Q absolute height of hole
            int colforpodos = 6;
            int colforopisanie = 8;
            int colforvozrast = 18;
            int colforshtrih = 19;
            int colforkonsist = 20;
            int colforige = 11;
            int colforobraznar = 9;
            int colforobrazmon = 10;
            int colforobrazwod = 14;
            int colforupw = 12;
            int colforuuw = 13;
            int debuglevel = 0;
            int colforpiketplus = 3;
            int colforpiket = 2;
            double dbleb = 0.0;
            //dbleb = getopissize();
            double xbegin = 0.0;
            //xbegin = getxbegin();
            //isa18 ap18 task dont ask too many questions
            string egesizetxt = "2.5";
            double egesize = 2.5;
            //double replace221= 0.0;
            egesizetxt = readconfig(@"egesize = ");
            if ((egesizetxt != "2") && (egesizetxt != "2.5") && (egesizetxt != "3") && (egesizetxt != "3.5") && (egesizetxt != "4") && (egesizetxt != "4.5") && (egesizetxt != "5") && (egesizetxt != "6"))
            {
                egesize = Getegesize();
            }
            else
            {
                egesize = double.Parse(egesizetxt);
            }
            double textsize = egesize;
            double deepestholepixels = 999;

            skwazya[] skwazi = new skwazya[1000];
            #region узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            for (int maxxllines = xlstartline; maxxllines < 5000; maxxllines++) //todo configit or autodetect
            {
                double picketallowed = 0;
                double plusallowed = 0;
                // ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка
                #region ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка

                string cellcontent = " "; xlserv.UniRead(1, maxxllines, 1, ref cellcontent);//sheet row col
                //04062016 ap6 check if skvanomer in allow list
                bool allowtofill = false;
                if ((cellcontent != "") && (cellcontent != null))
                {
                    allowtofill = true;

                    int numtosearch = 0;

                    Regex regex = new Regex(@"\d+");
                    Match match = regex.Match(cellcontent);
                    if (match.Success)
                    {
                        numtosearch = int.Parse(match.Value);
                    }
                    else
                    {
                        //Notification   SKWA bez numera
                    }

                    if (allowedskwnumlist.Contains(numtosearch))
                    {
                        allowtofill = true;
                        picketallowed = allowedskw_piket.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                        plusallowed = allowedskw_plus.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                    }
                    else
                    {
                        allowtofill = false;
                    }
                }
                #endregion

                if (allowtofill) //found NEW SKWA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                {
                    #region read first line
                    int maxgr = 0;
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    skwazi[goodskwacount] = new skwazya(goodskwacount);
                    skwazi[goodskwacount].name = cellcontent;
                    skwazi[goodskwacount].replace221 = 0.0;
                    string startbur = " "; xlserv.UniRead(1, maxxllines, colforstartbur, ref startbur); if ((startbur != "") && (startbur != null)) { skwazi[goodskwacount].start = startbur; }
                    string stopbur = " "; xlserv.UniRead(1, maxxllines, colforstopbur, ref stopbur); if ((stopbur != "") && (stopbur != null)) { skwazi[goodskwacount].stop = stopbur; }
                    string absustbur = " "; xlserv.UniRead(1, maxxllines, colforabsustbur, ref absustbur); if ((absustbur != "") && (absustbur != null))
                    {
                        skwazi[goodskwacount].absust = absustbur;
                        skwazi[goodskwacount].absust = Regex.Replace(skwazi[goodskwacount].absust, ",", ".");
                    }

                    /*
                    string piket = " "; xlserv.UniRead(1, maxxllines, colforpiket, ref piket); if ((piket != "") && (piket != null)) { skwazi[goodskwacount].piket = piket; } //Q column
                    string piketplus = " "; xlserv.UniRead(1, maxxllines, colforpiketplus, ref piketplus); if ((piketplus != "") && (piketplus != null))
                    {
                        skwazi[goodskwacount].piketplus = piketplus;
                        skwazi[goodskwacount].piketplus = Regex.Replace(skwazi[goodskwacount].piketplus, ",", ".");
                    }
                    */

                    skwazi[goodskwacount].skwax = (picketallowed * 100 + plusallowed) * xscale - xbegin;//todo start of left position at 0 even if starting piket is not 0/but 44meters
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    #endregion read first line

                    #region count Grunts IN SKWA
                    int gruntsinskwa = 0;
                    for (gruntsinskwa = 0; gruntsinskwa < 43; gruntsinskwa++)
                    {
                        string tempmaxlinesinskwa = "";
                        xlserv.UniRead(1, skwazi[goodskwacount].skwastartline + gruntsinskwa, colforopisanie, ref tempmaxlinesinskwa);
                        if ((tempmaxlinesinskwa == "") || (tempmaxlinesinskwa == null)) break;
                    }
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    #endregion count Grunts IN SKWA
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    skwazi[goodskwacount].layerscount = skwazi[goodskwacount].gruntsinskwa;
                    goodskwacount++;
                }

                picketallowed = 0;
                plusallowed = 0;
            }
            #endregion узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            #region дозаполнил скважины (а в планах = если есть только описания -и ничего кроме описаний, то нарисовать описания уже. а если описание и мощность, то еще лучше)
            for (int readmoreeveryskwa = 0; readmoreeveryskwa < goodskwacount; readmoreeveryskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[readmoreeveryskwa].gruntsinskwa; gruntinskwa++)
                {
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                    skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva = Regex.Replace(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva, ",", ".");
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforopisanie, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].opisanie);
                    //xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos + 1, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh);
                    if (gruntinskwa == 0) { skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva; }
                    else
                    {
                        // skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva - skwazi[readmoreeveryskwa].gruntiki[gruntinskwa-1].podoshva; }
                        //A__________                                               =  B______                                                  -   C____________
                        double B = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                        double C = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa - 1].podoshva);
                        double A = B - C;
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = A.ToString();
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobraznar, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obraznar);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobrazmon, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obrazmon);

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforige, ref ige);
                    if ((ige != "") && (ige != null)) skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].ige = ige;

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforvozrast, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].vozrast);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforshtrih, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforkonsist, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforupw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].upw);
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforuuw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].uuw);
                }
                if (debuglevel > 1) MgdAcApplication.ShowAlertDialog("из екселя прочитана скважина  : " + readmoreeveryskwa + "по счету от 0. например 0,1,2");
            }
            #endregion

            #region calc all skwa for totalmosh
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                double temptotalmosh = 0.0;
                int templaycoutmax = skwazi[writeskvcnt].layerscount;
                for (int s = 0; s < templaycoutmax; s++)
                {
                    double unparce = 0.0;
                    string tempsting = skwazi[writeskvcnt].gruntiki[s].tolshmosh;
                    double.TryParse(tempsting, out unparce);
                    temptotalmosh = temptotalmosh + unparce;
                }
                //double totalmosh = skwazi[goodskwacount].gettotalmosh();
                skwazi[writeskvcnt].totalmosh = temptotalmosh;
            }
            #endregion calc all skwa for totalmosh
            #region Check all values
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                if ((skwazi[writeskvcnt].absust == "") || (skwazi[writeskvcnt].absust == null)) skwazi[writeskvcnt].absust = "0";
                skwazi[writeskvcnt].absdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) - skwazi[writeskvcnt].totalmosh;
                skwazi[writeskvcnt].visualdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) * yscale - skwazi[writeskvcnt].totalmosh * geoscale;

                if (deepesthole > skwazi[writeskvcnt].absdeepestpnt) { deepesthole = skwazi[writeskvcnt].absdeepestpnt; };
                if (deepestholepixels > skwazi[writeskvcnt].visualdeepestpnt) { deepestholepixels = skwazi[writeskvcnt].visualdeepestpnt; };
            }
            deepestholepixels = deepestholepixels + zapas;
            #endregion Check all values

            if (autohor == "yes") { uhorizm = deepestholepixels / yscale; uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ }//se29 todo print uhorizm
            else { uhorizm = getuhorizm(); uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ };

            #region calc FLOOR and MID levels
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region calc floor
                skwazi[writeskvcnt].replace221 = (double.Parse(skwazi[writeskvcnt].absust) - uhorizm) * yscale;
                double tempnextceil = skwazi[writeskvcnt].replace221;
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    skwazi[writeskvcnt].gruntiki[readgr].konsiceiling = tempnextceil;
                    skwazi[writeskvcnt].gruntiki[readgr].konsifloor = tempnextceil - geoscale * double.Parse(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh);//10 - > geoscale
                    skwazi[writeskvcnt].gruntiki[readgr].circen = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2;
                    tempnextceil = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                }
                #endregion calc floor
            }
            #endregion draw shapki calc FLOOR and MID levels

            #region drawall
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region text ige and other
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }

                #endregion
                #region REZ shtrih - podoshva text - REZ
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    //left part of skwa
                    //ися10 без штриховок - другии версии командnohole(skwazi[writeskvcnt].skwax - 5.25 - varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);//left part of skwa

                    //text(skwazi[writeskvcnt].skwax + 17, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax + 4, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur));

                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);
                    //text(skwazi[writeskvcnt].skwax - 17, floorbydata, (tempabsuseverylayer  - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax - 12, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        //ися10 без штриховок - другии версии командhole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                    }
                    //ися10 без штриховок - другии версии командelse nohole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                }
                #endregion
                #region konsist
                //konsist
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    string temp = skwazi[writeskvcnt].gruntiki[readgr].konsist;
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((temp == "") || (temp == null))
                    {
                        varemptykonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, varpipe);
                        continue;
                    }
                    varnoholekonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].konsist, varpipe);
                }
                #endregion
                #region podval isa-1
                text((skwazi[writeskvcnt].skwax), (-39), skwazi[writeskvcnt].name);

                string absusttext = skwazi[writeskvcnt].absust;
                absusttext = absusttext.Replace(",", ".");
                double tempabsus = double.Parse(absusttext);
                text((skwazi[writeskvcnt].skwax), (-49), tempabsus.ToString(podvalotmetkaustaccur));

                text((skwazi[writeskvcnt].skwax), (-59), skwazi[writeskvcnt].totalmosh.ToString(podvaldeepaccur));

                //fe9 2016 task to do sort before kalkulating distances between/ because holes are at random in source file 
                /* 
                if (writeskvcnt > 0)
                {
                    text((skwazi[writeskvcnt - 1].skwax + (skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / 2), (-39), ((skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwazi[writeskvcnt].skwax / 2), (-39), (skwazi[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }
*/
                text((skwazi[writeskvcnt].skwax), (-79), skwazi[writeskvcnt].stop);

                palkaatfloor(4, skwazi[writeskvcnt].skwax);
                #endregion podval
            }
            #endregion drawall

            #region   isa-2 draw sorted dist at podval - we can use this sorted as main but it is bad for debugging - all skwaz mixed as the rezult and we have fk if input bad
            skwazya[] skwagood = new skwazya[goodskwacount];
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                skwagood[writeskvcnt] = skwazi[writeskvcnt];
            }

            skwazya[] skwasortedx = sortskwa(skwagood);

            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region podval

                if (writeskvcnt > 0)
                {
                    text((skwasortedx[writeskvcnt - 1].skwax + (skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / 2), (-69), ((skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwasortedx[writeskvcnt].skwax / 2), (-69), (skwasortedx[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }

                #endregion podval
            }

            #endregion draw sorted dist at podval


            #region XL Read samples of every HOLE of every LAYER
            for (int curskwa = 0; curskwa < goodskwacount; curskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string monstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazmon, ref monstring);
                    monstring = Regex.Replace(monstring, @"\t|\n|\r", ";");
                    string[] split = monstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "mon");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string narstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobraznar, ref narstring);
                    narstring = Regex.Replace(narstring, @"\t|\n|\r", ";");
                    string[] split = narstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "nar");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "nar");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string wodstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazwod, ref wodstring);
                    wodstring = Regex.Replace(wodstring, @"\t|\n|\r", ";");
                    string[] split = wodstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "wod");
                        }
                    }
                }
            }
            #endregion XL Read samples of every HOLE of every LAYER
            #region sort    samples //todo sort (take into consideration level)
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {
                skwazi[writeskvcnt].obriki = sortobriki(skwazi[writeskvcnt].obriki);
                // skwazi[writeskvcnt].obriki = pushl1tol2(skwazi[writeskvcnt].obriki);
                for (int curproba = 1; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    double diff = 0.0; diff = skwazi[writeskvcnt].obriki[curproba].bottom - skwazi[writeskvcnt].obriki[curproba - 1].bottom; if (diff < 0) { diff *= -1; }
                    if (diff < 1.9 / geoscale) skwazi[writeskvcnt].obriki[curproba].level = skwazi[writeskvcnt].obriki[curproba - 1].level + 1;//do se25 0.19
                }
            }
            #endregion
            #region drawall samples
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {                                                              // colonka by kolonka
                for (int curproba = 0; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "mon") probamon(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "nar") probanar(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "wod") probawod(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                }
            }
            #endregion drawall probs
            #region draw wodka
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                int ifanywaterhere = 0;

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if ((skwazi[writeskvcnt].gruntiki[readgr].uuw != "") && (skwazi[writeskvcnt].gruntiki[readgr].uuw != null))
                    {
                        double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw) * geoscale;
                        uuwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw).ToString("0.00"), skwazi[writeskvcnt].stop);
                        ifanywaterhere++;
                    };

                    if (skwazi[writeskvcnt].gruntiki[readgr].uuw != skwazi[writeskvcnt].gruntiki[readgr].upw)
                    {
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            upwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw).ToString("0.00"), skwazi[writeskvcnt].start);
                            ifanywaterhere++;
                        };
                    }
                }

                /*if (ifanywaterhere == 0)
                {
                    double viziblebottomy = skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor;
                    verttext((210 * writeskvcnt + 175), (222 + viziblebottomy) / 2 - 3, "Воды");
                    verttext((210 * writeskvcnt + 178), (222 + viziblebottomy) / 2 - 2, "нет");
                }*/
            }
            #endregion drawall
            #region task wodavpeske
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("PESOK") || skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("pesok"))
                    {
                        //search uuw ??upw?  in ALL GRUNTS or only this Grunt layer?
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double watertop = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                            double waterbot = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            //ne budu //if ((double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) > skwazi[writeskvcnt].gruntiki[readgr].konsiceiling)&&(double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) < skwazi[writeskvcnt].gruntiki[readgr].konsifloor)){  };
                            //noholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID");
                            //noholekonsist(skwazi[writeskvcnt].skwax-2.35, thislevel, waterbot, "SOLID");
                            varnoholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID", varpipe);
                        }
                    }
                }
            }
            #endregion task wodavpeske
            #region task palochki pod kolonkoy
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                palkapodkol(skwazi[writeskvcnt].skwax, skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor);
            }
            #endregion task palochki pod kolonkoy
            #endregion kolonki
            xlserv.Close();

            #region Horiz prodolshzhenie (no loading)
            #region horiz scale
            //#region #endregion 
            double[] lineyscaled = new double[5000];
            double[] linexscaled = new double[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                linexscaled[lineinxl] = lineforscalingx[lineinxl] * xscale;
                lineyscaled[lineinxl] = lineforscalingy[lineinxl] * yscale;
            }
            #endregion horiz scale
            #region horiz scale again add USLOV HORIZ -conditional horizon
            //#region #endregion scale again
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                //lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhoriz * yscale;
                lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhorizm * yscale;
            }
            #endregion  scale again add USLOV HORIZ -conditional horizon
            #region horiz draw scaled
            //#region #endregion draw scaled
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var pl = new Polyline();
                for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                {
                    pl.AddVertexAt(lineinxl - 1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, 0, 0);//original here --> pl.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);last here --->pl.AddVertexAt(3, new Point2d(0, 10), 0, 0, 0); pl.Closed = true;
                }

                pl.Closed = false;
                acBlkTblRec.AppendEntity(pl);
                acTrans.AddNewlyCreatedDBObject(pl, true);
                acTrans.Commit();
            }
            #endregion draw scaled

            #region     task ordinata //ися1 нужны вертикальные подписи и вертикальные палочки на отметки с комментариями даже без комментов
            int isnewpicket = 0;

            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if ((ordinatapart2[lineinxl] != null) && (ordinatapart2[lineinxl] != "") && (ordinatapart2[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart2[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
                if ((ordinatapart1[lineinxl] != null) && (ordinatapart1[lineinxl] != "") && (ordinatapart1[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart1[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

                //ися3 поправки ися1 теперь нужны все верт отметки даже без комментов
                //ися4 if modeconfigordinata >0 than

                if (lineinxl > 0)
                { //try isa6 
                    if (lineforscalingpikets[lineinxl] != lineforscalingpikets[lineinxl - 1]) isnewpicket = 1;
                    else isnewpicket = 0;
                }

                if (isnewpicket == 0) //
                {
                    #region ися3 верт linii даже без комментов
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.ColorIndex = 211;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);

                        var acPoly2 = new Polyline(1);
                        acPoly2.Normal = Vector3d.ZAxis;
                        acPoly2.ColorIndex = 211;
                        acPoly2.AddVertexAt(0, new Point2d(linexscaled[lineinxl], -10), 0, -1, -1);
                        acPoly2.AddVertexAt(1, new Point2d(linexscaled[lineinxl], -20), 0, -1, -1);
                        acPoly2.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly2);
                        acTrans.AddNewlyCreatedDBObject(acPoly2, true);


                        // ися2 новые3строки подвала отметки
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = lineforscalingy[lineinxl].ToString();
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, -9, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        acBlkTblRec.AppendEntity(actext);
                        acTrans.AddNewlyCreatedDBObject(actext, true);

                        if (lineinxl > 1)
                        {//try isa6 
                            double dist_between_otmetki = lineforscalingx[lineinxl] - lineforscalingx[lineinxl - 1];
                            double distscaled = linexscaled[lineinxl] - linexscaled[lineinxl - 1];

                            MText acMText = new MText();
                            acMText.Location = new Point3d(linexscaled[lineinxl] - distscaled / 2, -19, 0);
                            acMText.Height = 2;
                            acMText.TextHeight = 2;
                            acMText.Width = 20;
                            if (distscaled < 10) { acMText.Rotation = 1.57079; acMText.Location = new Point3d(linexscaled[lineinxl] - distscaled / 2 + 1, -19 + 4, 0); }
                            else acMText.Rotation = 0;
                            acMText.Contents = dist_between_otmetki.ToString("0.00");
                            acMText.Attachment = AttachmentPoint.BottomCenter;
                            acBlkTblRec.AppendEntity(acMText);
                            acTrans.AddNewlyCreatedDBObject(acMText, true);
                            //nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                        }

                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

                if (isnewpicket == 1) //
                {
                    // текст1 = 100 - последняя плюсовка предыдущего пикета 
                    // текст2 = плюсовка нового пикета
                    #region  PICKEt line
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        //differense in hight do picket and after  //diffhppr diffxun koef
                        double diffhppr = lineyscaled[lineinxl] - lineyscaled[lineinxl - 1];
                        // текст2 = плюсовка нового пикета delim na distance
                        double diffxun = lineforscalingx[lineinxl] - lineforscalingx[lineinxl - 1];
                        double koef = lineforscalingplusos[lineinxl] / diffxun;
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        #region linia na podvale
                        var acPoly2 = new Polyline(1);             //linia na podvale 
                        acPoly2.Normal = Vector3d.ZAxis;
                        acPoly2.ColorIndex = 50;
                        acPoly2.AddVertexAt(0, new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, -10), 0, -1, -1);
                        acPoly2.AddVertexAt(1, new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, -30), 0, -1, -1);
                        acPoly2.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly2);
                        acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                        #endregion linia na podvale
                        #region     текст1 = 100 - последняя плюсовка предыдущего пикета

                        // текст1 = 100 - последняя плюсовка предыдущего пикета 
                        double distafterprevpiket = 100 - lineforscalingplusos[lineinxl - 1];


                        MText acMText = new MText();
                        acMText.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 - distafterprevpiket / 2) * xscale, -19, 0);
                        acMText.Height = 2;
                        acMText.TextHeight = 2;
                        acMText.Width = 20;
                        if ((distafterprevpiket / 2 * xscale) < 10) { acMText.Rotation = 1.57079; acMText.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 - distafterprevpiket / 2 + 1) * xscale, -19 + 4, 0); }
                        else acMText.Rotation = 0;
                        acMText.Contents = distafterprevpiket.ToString("0.00");
                        acMText.Attachment = AttachmentPoint.BottomCenter;
                        acBlkTblRec.AppendEntity(acMText);
                        acTrans.AddNewlyCreatedDBObject(acMText, true);
                        #endregion
                        #region  TextString = "ПК "

                        // ися2 новые3строки подвала отметки
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        //actext.TextString = "ПК " + lineforscalingpikets[lineinxl].ToString()  ;

                        //diffhppr diffxun koef
                        //actext.TextString = "ПК " + lineforscalingpikets[lineinxl].ToString() + " otlad diffhppr:" + diffhppr + " otlad diffxun:" + diffxun + " otlad koef:" + koef;
                        actext.TextString = "ПК " + lineforscalingpikets[lineinxl].ToString();

                        actext.Position = new Point3d(lineforscalingpikets[lineinxl] * xscale * 100, -29, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 0;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        acBlkTblRec.AppendEntity(actext);
                        acTrans.AddNewlyCreatedDBObject(actext, true);
                        #endregion

                        #region linia piketa
                        var acPoint = new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, lineyscaled[lineinxl] - diffhppr * koef), 0, -1, -1);
                        acPoly.ColorIndex = 50;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        #endregion linia piketa
                        #region текст2 = плюсовка нового пикета

                        /*
                        DBText actext2 = new DBText();
                        actext2.SetDatabaseDefaults();
                        actext2.TextString = lineforscalingplusos[lineinxl].ToString("0.00");
                        actext2.Position = new Point3d((lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl]/2) * xscale , -19, 0);
                        actext2.Height = 2;
                        actext2.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext2.Rotation = 0;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        acBlkTblRec.AppendEntity(actext2);
                        acTrans.AddNewlyCreatedDBObject(actext2, true);
*/

                        MText acMText2 = new MText();
                        acMText2.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl] / 2) * xscale, -19, 0);
                        acMText2.Height = 2;
                        acMText2.TextHeight = 2;
                        acMText2.Width = 20;
                        if ((lineforscalingplusos[lineinxl] / 2 * xscale) < 10) { acMText2.Rotation = 1.57079; acMText2.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl] / 2 + 1) * xscale, -19 + 4, 0); }
                        else acMText2.Rotation = 0;
                        acMText2.Contents = lineforscalingplusos[lineinxl].ToString("0.00");
                        acMText2.Attachment = AttachmentPoint.BottomCenter;
                        acBlkTblRec.AppendEntity(acMText2);
                        acTrans.AddNewlyCreatedDBObject(acMText2, true);
                        #endregion
                        /*
                        #region 
                        #endregion
                        */
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                    #region линия сразу после пикета
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.ColorIndex = 211;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);

                        var acPoly2 = new Polyline(1);
                        acPoly2.Normal = Vector3d.ZAxis;
                        acPoly2.ColorIndex = 211;
                        acPoly2.AddVertexAt(0, new Point2d(linexscaled[lineinxl], -10), 0, -1, -1);
                        acPoly2.AddVertexAt(1, new Point2d(linexscaled[lineinxl], -20), 0, -1, -1);
                        acPoly2.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly2);
                        acTrans.AddNewlyCreatedDBObject(acPoly2, true);


                        /* нужно разбить на две части
                         * 
                                                // ися2 новые3строки подвала отметки
                                                //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                                                DBText actext = new DBText();
                                                actext.SetDatabaseDefaults();
                                                actext.TextString = lineforscalingy[lineinxl].ToString();
                                                actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, -70, 0);
                                                actext.Height = 2;
                                                actext.Thickness = 22;
                                                //actext.ColorIndex = 0;
                                                actext.Rotation = 1.57079;
                                                //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                                                acBlkTblRec.AppendEntity(actext);
                                                acTrans.AddNewlyCreatedDBObject(actext, true);
                                                */

                        /*
                                                DBText actext2 = new DBText();
                                                actext2.SetDatabaseDefaults();

                                                double dist_between_otmetki = lineforscalingx[lineinxl] - lineforscalingx[lineinxl - 1];
                                                actext2.TextString = dist_between_otmetki.ToString("0.00");

                                                double distscaled = linexscaled[lineinxl] - linexscaled[lineinxl - 1];
                                                actext2.Position = new Point3d(linexscaled[lineinxl] - distscaled / 2, -90, 0);
                                                actext2.Height = 2;
                                                actext2.Thickness = 22;
                                                //actext.ColorIndex = 0;
                                                if (distscaled < 30) actext2.Rotation = 1.57079;
                                                else actext2.Rotation = 0;
                                                //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                                                acBlkTblRec.AppendEntity(actext2);
                                                acTrans.AddNewlyCreatedDBObject(actext2, true);
                                      */

                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

            }
            #endregion  task ordinata


            #region boxunderground
            double firstx = linexscaled[1];
            double firsty = lineyscaled[1];
            double lastx = linexscaled[kartacounter - 1];
            double lasty = lineyscaled[kartacounter - 1];

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(firstx, firsty), 0, -1, -1);  //do ok27 acPoly.AddVertexAt(0, new Point2d(firstx, firsty)   , 0, -1, -1); 
                acPoly.AddVertexAt(1, new Point2d(0, 0), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, lasty), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion boxunderground
            int podvalcount = 1;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметки поверхности земли", -83, -9, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 2;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояния между отметками", -83, -19, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 3;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Пикетаж", -83, -29, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 4;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Номеp скважины", -83, -39, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 5;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметка устья, м", -83, -49, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 6;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Глубина, м", -83, -59, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 7;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояние, м", -83, -69, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 8;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Дата пpоходки", -83, -79, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval

            string scalextext = "Масштаб по оси X 1 : " + xdevidor;
            lefttext(scalextext, -65, 12, 2);
            string scaleytext = "Масштаб по оси Y 1 : " + ydevidor;
            lefttext(scaleytext, -65, 7.5, 2);
            string geoscaletext = "Масштаб геол. 1 : " + geodevidor;
            lefttext(geoscaletext, -65, 3, 2);

            #endregion
            lineikaok26(maximgeoline, uhorizm, yscale);
        }
        [CommandMethod("profile_a_sboku_shtrihovka")]
        public void profileasboku()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region ask some
            //dynamic acadObject = MgdAcApplication.AcadApplication;//zoom
            //acadObject.ZoomExtents();//zoom
            double yscale = 10;//calculated in this way 1000/100= 10
            double xscale = 2; //calculated in this way 1000/500= 2
            double uhorizm = 60;
            string autohor = "";

            autohor = readconfig(@"autohor = ");//isa18 ap18 task dont ask too many questions
            if ((autohor != "no") && (autohor != "yes"))
            {
                autohor = getautohor();
            } 
            if (autohor == "no") { uhorizm = getuhorizm(); };

            double zapas = 0;
            zapas = readzapas();
            string podvaldistbetweenaccur = readpodvaldistbetweenaccur();
            string lefttoholeaccur = readlefttoholeaccur();
            string righttoholeaccur = readrighttoholeaccur();
            string podvalotmetkaustaccur = readpodvalotmetkaustaccur();
            string podvaldeepaccur = readpodvaldeepaccur();
            double varpipe = readdiamskwa(); varpipe = varpipe / 2;
            double uhoriz = uhorizm * 10;
            double deepesthole = 999;

            double xdevidor = getxscale();
            xscale = 1000 / xdevidor;

            double ydevidor = getyscale();
            yscale = 1000 / ydevidor;

            double geoscale = 10;
            double geodevidor = getgeoscale();
            geoscale = 1000 / geodevidor;
            double maximgeoline = -1500;
            #endregion ask some

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );
            
            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);
            #region horizline LOAD
            #region detect xl maximum -----> kartacounter

            string light = "green";
            int kartacounter = 1;
            for (; kartacounter < 5000; kartacounter++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(1, kartacounter, 1, ref cellcontent);//sheet row col
                if (((cellcontent == "") || (cellcontent == null)) && (kartacounter == 1))
                {
                    light = "red"; break; ;
                }
                if ((cellcontent == "") || (cellcontent == null))
                {
                    break; ;
                }
            }
            //kartacounter is maximum lines in exclell
            #endregion detect xl maximum -----> kartacounter
            #region drawline_no_scaling
            double maxyininput = 0;
            double maxy = 0;

            string firstpntx = "0";
            string firstpntxplus = "0";
            string firstpnty = "0";

            string nextx = "0";
            string nextxplus = "0";
            string nexty = "0";
            /*
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            firstpntx = " ";
                            xlserv2.UniRead(1, 1, 1, ref firstpntx);//sheet row col
                             firstpntxplus = " "; xlserv2.UniRead(1, 1, 2, ref firstpntxplus);//sheet row col
                            firstpnty = " ";
                            xlserv2.UniRead(1, 1, 3, ref firstpnty);//sheet row col

                            var acPoint = new Point2d((double.Parse(firstpntx) * 100 + double.Parse(firstpntxplus))*xscale-xbegin, double.Parse(firstpnty) * yscale - uhoriz);

                            if (double.Parse(firstpnty) * yscale - uhoriz > maxy) {maxy = double.Parse(firstpnty) * yscale - uhoriz; }
                            var acPoly = new Polyline(kartacounter-1);
                            acPoly.Normal = Vector3d.ZAxis;
                            acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                            for (int lineinxl = 2; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                            {
                                 nextx = " "; xlserv2.UniRead(1, lineinxl, 1, ref nextx);//sheet row col
                                 nextxplus = " "; xlserv2.UniRead(1, lineinxl, 2, ref nextxplus);//sheet row col
                                nexty = " "; xlserv2.UniRead(1, lineinxl, 3, ref nexty);//sheet row col
                                acPoly.AddVertexAt(lineinxl - 1, new Point2d((double.Parse(nextx) * 100 + double.Parse(nextxplus))*xscale-xbegin, double.Parse(nexty) * yscale - uhoriz), 0, -1, -1);
                                if (double.Parse(nexty) * yscale - uhoriz > maxy) { maxy = double.Parse(nexty) * yscale - uhoriz; }
                            }
                            acPoly.Closed = false;

                            acBlkTblRec.AppendEntity(acPoly);
                            acTrans.AddNewlyCreatedDBObject(acPoly, true);
                            acTrans.Commit();
                        }
                    */
            #endregion drawline_no_scaling
            #region horiz load
            //#region #endregion horiz load
            double[] lineforscalingpikets = new double[5000];
            double[] lineforscalingplusos = new double[5000];
            double[] lineforscalingy = new double[5000];
            double[] lineforscalingx = new double[5000];
            string[] ordinatapart1 = new string[5000];
            string[] ordinatapart2 = new string[5000];
            string[] ordinatabothparts = new string[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                string cellcontentpk = " "; xlserv2.UniRead(1, lineinxl, 1, ref cellcontentpk);
                lineforscalingpikets[lineinxl] = double.Parse(cellcontentpk);
                string cellcontentps = " "; xlserv2.UniRead(1, lineinxl, 2, ref cellcontentps);
                lineforscalingplusos[lineinxl] = double.Parse(cellcontentps);
                string cellcontenty = " "; xlserv2.UniRead(1, lineinxl, 3, ref cellcontenty);
                lineforscalingy[lineinxl] = double.Parse(cellcontenty);
                lineforscalingx[lineinxl] = lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl];

                string ordinatapart1content = " "; xlserv2.UniRead(1, lineinxl, 4, ref ordinatapart1content);
                ordinatapart1[lineinxl] = ordinatapart1content;
                string ordinatapart2content = " "; xlserv2.UniRead(1, lineinxl, 5, ref ordinatapart2content);
                ordinatapart2[lineinxl] = ordinatapart2content;
                ordinatabothparts[lineinxl] = ordinatapart1[lineinxl] + " " + ordinatapart2[lineinxl];

                if (double.Parse(cellcontenty) > maximgeoline) { maximgeoline = double.Parse(cellcontenty); };
            }
            #endregion horiz load

            #region   ися9 как по базе с скв
            //paragraph.ToLower(culture).Contains(word.ToLower(culture)) with CultureInfo.InvariantCulture
            List<string> allowedskwnameslist = new List<string>(999);
            List<double> allowedskw_piket = new List<double>(999);
            List<double> allowedskw_plus = new List<double>(999);
            List<int> allowedskwnumlist = new List<int>(999);
            //if NEED proverka IF mode exeptional skwas not ALL swas as usual
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if (Regex.IsMatch(ordinatabothparts[lineinxl], "скв", RegexOptions.IgnoreCase))
                {
                    allowedskwnameslist.Add(ordinatabothparts[lineinxl]);
                    //allowedskwnameslist.Add("sposob 1 reg");
                    allowedskw_piket.Add(lineforscalingpikets[lineinxl]);
                    allowedskw_plus.Add(lineforscalingplusos[lineinxl]);
                }
            }

            TextWriter nextoutfile = new StreamWriter("c:\\AtlPProf\\rezultat_proverki.txt", true, Encoding.Default);
            nextoutfile.WriteLine(" _ _ _ ");
            
            int ap5counter = 0;
            foreach (string value in allowedskwnameslist) 
            {
                nextoutfile.WriteLine(value);
                Regex regex = new Regex(@"\d+");
                Match match = regex.Match(value);
                if (match.Success)
                {
                    nextoutfile.WriteLine(match.Value + "only digit");
                    allowedskwnumlist.Add(int.Parse(match.Value));
                }
                nextoutfile.WriteLine(match.Value + " piket " + allowedskw_piket[ap5counter] + " plus " +  allowedskw_plus[ap5counter]);
                    ap5counter++;

                Match match2 = Regex.Match(value.ToLower(), @"скв([.]+)$", RegexOptions.IgnoreCase);
                if (match2.Success)
                {
                    // Finally, we get the Group value and display it.
                    string key = match2.Groups[1].Value;
                    nextoutfile.WriteLine(key);
                }
            };
            nextoutfile.WriteLine(" _ _ _ ");
            nextoutfile.Close();
            
            /*
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match("Dot 55 Perls");
            if (match.Success)
            {
                Console.WriteLine(match.Value);
            }
            */
            #endregion

            #endregion horizline
            xlserv2.Close();

            ComplexObjects.ExcelServ xlserv = new ComplexObjects.ExcelServ();
            #region open xl
            var ofd =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "*",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );
            var dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK)
                return;
            // Display the name of the file and the contained sheets
            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd.Filename
            );
            #endregion open xl
            xlserv.Open(ofd.Filename);
            #region kolonki
            int goodskwacount = 0;
            int xlstartline = 5;//todo configit starting line
            int colforstartbur = 15;
            int colforstopbur = 16;
            int colforabsustbur = 17;//Q absolute height of hole
            int colforpodos = 6;
            int colforopisanie = 8;
            int colforvozrast = 18;
            int colforshtrih = 19;
            int colforkonsist = 20;
            int colforige = 11;
            int colforobraznar = 9;
            int colforobrazmon = 10;
            int colforobrazwod = 14;
            int colforupw = 12;
            int colforuuw = 13;
            int debuglevel = 0;
            int colforpiketplus = 3;
            int colforpiket = 2;
            double dbleb = 0.0;
            //dbleb = getopissize();
            double xbegin = 0.0;
            //xbegin = getxbegin();
            //isa18 ap18 task dont ask too many questions
            string egesizetxt = "2.5";
            double egesize = 2.5;
            //double replace221= 0.0;
            egesizetxt = readconfig(@"egesize = ");
            if ((egesizetxt != "2") && (egesizetxt != "2.5") && (egesizetxt != "3") && (egesizetxt != "3.5") && (egesizetxt != "4") && (egesizetxt != "4.5") && (egesizetxt != "5") && (egesizetxt != "6"))
            {
                egesize = Getegesize();
            }
            else
            {
                egesize = double.Parse(egesizetxt);
            }
            double textsize = egesize;
            double deepestholepixels = 999;

            skwazya[] skwazi = new skwazya[1000];
            #region узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            for (int maxxllines = xlstartline; maxxllines < 5000; maxxllines++) //todo configit or autodetect
            {
                double picketallowed = 0;
                double plusallowed = 0;
                // ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка
                #region ися9 как по базе с сквами 9 штук 751 и 953 для юзера подстройка
                
                string cellcontent = " "; xlserv.UniRead(1, maxxllines, 1, ref cellcontent);//sheet row col
                //04062016 ap6 check if skvanomer in allow list
                bool allowtofill = false;
                if ((cellcontent != "") && (cellcontent != null))
                {
                    allowtofill = true;

                    int numtosearch = 0;

                    Regex regex = new Regex(@"\d+");
                    Match match = regex.Match(cellcontent);
                    if (match.Success)
                    {
                        numtosearch = int.Parse(match.Value);
                    }
                    else
                    {
                        //Notification   SKWA bez numera
                    }

                    if (allowedskwnumlist.Contains(numtosearch))
                    {
                        allowtofill = true;
                        picketallowed = allowedskw_piket.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                        plusallowed = allowedskw_plus.ElementAt(allowedskwnumlist.IndexOf(numtosearch));
                    }
                    else
                    {
                        allowtofill = false;
                    }
                }
#endregion

                if (allowtofill) //found NEW SKWA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                {
                    #region read first line
                    int maxgr = 0;
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    skwazi[goodskwacount] = new skwazya(goodskwacount);
                    skwazi[goodskwacount].name = cellcontent;
                    skwazi[goodskwacount].replace221 = 0.0;
                    string startbur = " "; xlserv.UniRead(1, maxxllines, colforstartbur, ref startbur); if ((startbur != "") && (startbur != null)) { skwazi[goodskwacount].start = startbur; }
                    string stopbur = " "; xlserv.UniRead(1, maxxllines, colforstopbur, ref stopbur); if ((stopbur != "") && (stopbur != null)) { skwazi[goodskwacount].stop = stopbur; }
                    string absustbur = " "; xlserv.UniRead(1, maxxllines, colforabsustbur, ref absustbur); if ((absustbur != "") && (absustbur != null))
                    {
                        skwazi[goodskwacount].absust = absustbur;
                        skwazi[goodskwacount].absust = Regex.Replace(skwazi[goodskwacount].absust, ",", ".");
                    }

                    /*
                    string piket = " "; xlserv.UniRead(1, maxxllines, colforpiket, ref piket); if ((piket != "") && (piket != null)) { skwazi[goodskwacount].piket = piket; } //Q column
                    string piketplus = " "; xlserv.UniRead(1, maxxllines, colforpiketplus, ref piketplus); if ((piketplus != "") && (piketplus != null))
                    {
                        skwazi[goodskwacount].piketplus = piketplus;
                        skwazi[goodskwacount].piketplus = Regex.Replace(skwazi[goodskwacount].piketplus, ",", ".");
                    }
                    */

                    skwazi[goodskwacount].skwax = (picketallowed * 100 + plusallowed) * xscale - xbegin;//todo start of left position at 0 even if starting piket is not 0/but 44meters
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    #endregion read first line

                    #region count Grunts IN SKWA
                    int gruntsinskwa = 0;
                    for (gruntsinskwa = 0; gruntsinskwa < 43; gruntsinskwa++)
                    {
                        string tempmaxlinesinskwa = "";
                        xlserv.UniRead(1, skwazi[goodskwacount].skwastartline + gruntsinskwa, colforopisanie, ref tempmaxlinesinskwa);
                        if ((tempmaxlinesinskwa == "") || (tempmaxlinesinskwa == null)) break;
                    }
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    #endregion count Grunts IN SKWA
                    skwazi[goodskwacount].skwastartline = maxxllines;
                    skwazi[goodskwacount].gruntsinskwa = gruntsinskwa;
                    skwazi[goodskwacount].layerscount = skwazi[goodskwacount].gruntsinskwa;
                    goodskwacount++;
                }

                picketallowed = 0;
                plusallowed = 0;
            }
            #endregion узнал сколько скважин и сколько грунтов в скважине . судя по колонке с описанием
            #region дозаполнил скважины (а в планах = если есть только описания -и ничего кроме описаний, то нарисовать описания уже. а если описание и мощность, то еще лучше)
            for (int readmoreeveryskwa = 0; readmoreeveryskwa < goodskwacount; readmoreeveryskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[readmoreeveryskwa].gruntsinskwa; gruntinskwa++)
                {
                    string lastpodoshva = ""; string opisanie = ""; string tolshmosh = ""; string vozrast = ""; string shtrih = ""; string ige = ""; string obraznar = ""; string obrazmon = "";
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                    skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva = Regex.Replace(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva, ",", ".");
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforopisanie, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].opisanie);
                    //xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforpodos + 1, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh);
                    if (gruntinskwa == 0) { skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva; }
                    else
                    {
                        // skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva - skwazi[readmoreeveryskwa].gruntiki[gruntinskwa-1].podoshva; }
                        //A__________                                               =  B______                                                  -   C____________
                        double B = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].podoshva);
                        double C = double.Parse(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa - 1].podoshva);
                        double A = B - C;
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].tolshmosh = A.ToString();
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobraznar, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obraznar);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforobrazmon, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].obrazmon);

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforige, ref ige);
                    if ((ige != "") && (ige != null)) skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].ige = ige;

                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforvozrast, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].vozrast);
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforshtrih, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].shtrih = stringSql;
                    }
                    xlserv.UniRead(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforkonsist, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                    if (skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist.Contains("снег"))
                    {
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].hassnow = "yes";
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = cleansnow(skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist);
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    else
                    {
                        string stringSql = skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist;
                        stringSql = new String(stringSql.Where(x => !Char.IsWhiteSpace(x)).ToArray());
                        skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].konsist = stringSql;
                    }
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforupw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].upw);
                    xlserv.xlvodaread(1, skwazi[readmoreeveryskwa].skwastartline + gruntinskwa, colforuuw, ref skwazi[readmoreeveryskwa].gruntiki[gruntinskwa].uuw);
                }
                if (debuglevel > 1) MgdAcApplication.ShowAlertDialog("из екселя прочитана скважина  : " + readmoreeveryskwa + "по счету от 0. например 0,1,2");
            }
            #endregion

            #region calc all skwa for totalmosh
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                double temptotalmosh = 0.0;
                int templaycoutmax = skwazi[writeskvcnt].layerscount;
                for (int s = 0; s < templaycoutmax; s++)
                {
                    double unparce = 0.0;
                    string tempsting = skwazi[writeskvcnt].gruntiki[s].tolshmosh;
                    double.TryParse(tempsting, out unparce);
                    temptotalmosh = temptotalmosh + unparce;
                }
                //double totalmosh = skwazi[goodskwacount].gettotalmosh();
                skwazi[writeskvcnt].totalmosh = temptotalmosh;
            }
            #endregion calc all skwa for totalmosh
            #region Check all values
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                if ((skwazi[writeskvcnt].absust == "") || (skwazi[writeskvcnt].absust == null)) skwazi[writeskvcnt].absust = "0";
                skwazi[writeskvcnt].absdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) - skwazi[writeskvcnt].totalmosh;
                skwazi[writeskvcnt].visualdeepestpnt = double.Parse(skwazi[writeskvcnt].absust) * yscale - skwazi[writeskvcnt].totalmosh * geoscale;

                if (deepesthole > skwazi[writeskvcnt].absdeepestpnt) { deepesthole = skwazi[writeskvcnt].absdeepestpnt; };
                if (deepestholepixels > skwazi[writeskvcnt].visualdeepestpnt) { deepestholepixels = skwazi[writeskvcnt].visualdeepestpnt; };
            }
            deepestholepixels = deepestholepixels + zapas;
            #endregion Check all values

            if (autohor == "yes") { uhorizm = deepestholepixels / yscale; uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ }//se29 todo print uhorizm
            else { uhorizm = getuhorizm(); uhoriz = uhorizm * yscale; /*leftgeotxt(0, 0, uhorizm.ToString("0.0"));*/ };

            #region calc FLOOR and MID levels
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region calc floor
                skwazi[writeskvcnt].replace221 = (double.Parse(skwazi[writeskvcnt].absust) - uhorizm) * yscale;
                double tempnextceil = skwazi[writeskvcnt].replace221;
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    skwazi[writeskvcnt].gruntiki[readgr].konsiceiling = tempnextceil;
                    skwazi[writeskvcnt].gruntiki[readgr].konsifloor = tempnextceil - geoscale * double.Parse(skwazi[writeskvcnt].gruntiki[readgr].tolshmosh);//10 - > geoscale
                    skwazi[writeskvcnt].gruntiki[readgr].circen = (skwazi[writeskvcnt].gruntiki[readgr].konsiceiling + skwazi[writeskvcnt].gruntiki[readgr].konsifloor) / 2;
                    tempnextceil = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                }
                #endregion calc floor
            }
            #endregion draw shapki calc FLOOR and MID levels

            #region drawall
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region text ige and other
                /*for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);

                    undertext(skwazi[writeskvcnt].skwax, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur), varpipe);
                    leftgeotxt(skwazi[writeskvcnt].skwax - varpipe, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }*/

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        igeincircle(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].ige, textsize);
                    }
                }

                #endregion
                #region REZ shtrih - podoshva text - REZ
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                    //left part of skwa
                    nohole(skwazi[writeskvcnt].skwax - 5.25 - varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);//left part of skwa

                    //text(skwazi[writeskvcnt].skwax + 17, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax + 4, floorbydata, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva).ToString(righttoholeaccur));

                    string absusttexteverylayer = skwazi[writeskvcnt].absust;
                    absusttexteverylayer = absusttexteverylayer.Replace(",", ".");
                    double tempabsuseverylayer = double.Parse(absusttexteverylayer);
                    ///text(skwazi[writeskvcnt].skwax - 17, floorbydata, (tempabsuseverylayer  - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString("0.00"));
                    undertext(skwazi[writeskvcnt].skwax - 12, floorbydata, (tempabsuseverylayer - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].podoshva)).ToString(lefttoholeaccur));

                    if ((skwazi[writeskvcnt].gruntiki[readgr].ige != "") && (skwazi[writeskvcnt].gruntiki[readgr].ige != null))
                    {
                        hole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                    }
                    else nohole(skwazi[writeskvcnt].skwax + varpipe, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].shtrih, 93);
                }
                #endregion
                #region konsist
                //konsist
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    string temp = skwazi[writeskvcnt].gruntiki[readgr].konsist;
                    double clngbydata = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                    double floorbydata = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;

                    if ((temp == "") || (temp == null))
                    {
                        varemptykonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, varpipe);
                        continue;
                    }
                    varnoholekonsist(skwazi[writeskvcnt].skwax, floorbydata, clngbydata, skwazi[writeskvcnt].gruntiki[readgr].konsist, varpipe);
                }
                #endregion
                #region podval isa-1
                text((skwazi[writeskvcnt].skwax), (-39), skwazi[writeskvcnt].name);

                string absusttext = skwazi[writeskvcnt].absust;
                absusttext = absusttext.Replace(",", ".");
                double tempabsus = double.Parse(absusttext);
                text((skwazi[writeskvcnt].skwax), (-49), tempabsus.ToString(podvalotmetkaustaccur));

                text((skwazi[writeskvcnt].skwax), (-59), skwazi[writeskvcnt].totalmosh.ToString(podvaldeepaccur));

                //fe9 2016 task to do sort before kalkulating distances between/ because holes are at random in source file 
                /* 
                if (writeskvcnt > 0)
                {
                    text((skwazi[writeskvcnt - 1].skwax + (skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / 2), (-39), ((skwazi[writeskvcnt].skwax - skwazi[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwazi[writeskvcnt].skwax / 2), (-39), (skwazi[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }
*/
                text((skwazi[writeskvcnt].skwax), (-79), skwazi[writeskvcnt].stop);

                palkaatfloor(4, skwazi[writeskvcnt].skwax);
                #endregion podval
            }
            #endregion drawall

            #region   isa-2 draw sorted dist at podval - we can use this sorted as main but it is bad for debugging - all skwaz mixed as the rezult and we have fk if input bad
            skwazya[] skwagood = new skwazya[goodskwacount];
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                skwagood[writeskvcnt] = skwazi[writeskvcnt];
            }

            skwazya[] skwasortedx = sortskwa(skwagood);

            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                #region podval

                if (writeskvcnt > 0)
                {
                    text((skwasortedx[writeskvcnt - 1].skwax + (skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / 2), (-69), ((skwasortedx[writeskvcnt].skwax - skwasortedx[writeskvcnt - 1].skwax) / xscale).ToString(podvaldistbetweenaccur));
                }
                else
                {
                    text((skwasortedx[writeskvcnt].skwax / 2), (-69), (skwasortedx[writeskvcnt].skwax / xscale).ToString(podvaldistbetweenaccur));//realy it is null - skwax // so it means if we start not at null - for example at 5 ant skwa10 we should have 75 as center, not 5
                }

                #endregion podval
            }

            #endregion draw sorted dist at podval


            #region XL Read samples of every HOLE of every LAYER
            for (int curskwa = 0; curskwa < goodskwacount; curskwa++)
            {
                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string monstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazmon, ref monstring);
                    monstring = Regex.Replace(monstring, @"\t|\n|\r", ";");
                    string[] split = monstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "mon");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string narstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobraznar, ref narstring);
                    narstring = Regex.Replace(narstring, @"\t|\n|\r", ";");
                    string[] split = narstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "nar");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "nar");
                        }
                    }
                }

                for (int gruntinskwa = 0; gruntinskwa < skwazi[curskwa].layerscount; gruntinskwa++)
                {
                    string wodstring = "";
                    xlserv.xlskwaread(1, skwazi[curskwa].skwastartline + gruntinskwa, colforobrazwod, ref wodstring);
                    wodstring = Regex.Replace(wodstring, @"\t|\n|\r", ";");
                    string[] split = wodstring.Split(new Char[] { ';', ':' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")//without minuses it is good -=24au15=- skwazi[goodskwacount].addproba(double.Parse(s), "mon");
                        {
                            string[] splitmin = s.Split(new Char[] { '-', ':' });
                            skwazi[curskwa].addproba(splitmin[0], "wod");
                        }
                    }
                }
            }
            #endregion XL Read samples of every HOLE of every LAYER
            #region sort    samples //todo sort (take into consideration level)
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {
                skwazi[writeskvcnt].obriki = sortobriki(skwazi[writeskvcnt].obriki);
                // skwazi[writeskvcnt].obriki = pushl1tol2(skwazi[writeskvcnt].obriki);
                for (int curproba = 1; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    double diff = 0.0; diff = skwazi[writeskvcnt].obriki[curproba].bottom - skwazi[writeskvcnt].obriki[curproba - 1].bottom; if (diff < 0) { diff *= -1; }
                    if (diff < 1.9 / geoscale) skwazi[writeskvcnt].obriki[curproba].level = skwazi[writeskvcnt].obriki[curproba - 1].level + 1;//do se25 0.19
                }
            }
            #endregion
            #region drawall samples
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)    // colonka by kolonka
            {                                                              // colonka by kolonka
                for (int curproba = 0; curproba < skwazi[writeskvcnt].probacount + 1; curproba++)
                {
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "mon") probamon(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "nar") probanar(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                    if (skwazi[writeskvcnt].obriki[curproba].forma == "wod") probawod(skwazi[writeskvcnt].skwax + 5 + skwazi[writeskvcnt].obriki[curproba].level * 2.5, skwazi[writeskvcnt].replace221 - skwazi[writeskvcnt].obriki[curproba].bottom * geoscale);
                }
            }
            #endregion drawall probs
            #region draw wodka
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                int ifanywaterhere = 0;

                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if ((skwazi[writeskvcnt].gruntiki[readgr].uuw != "") && (skwazi[writeskvcnt].gruntiki[readgr].uuw != null))
                    {
                        double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw) * geoscale;
                        uuwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].uuw).ToString("0.00"), skwazi[writeskvcnt].stop);
                        ifanywaterhere++;
                    };

                    if (skwazi[writeskvcnt].gruntiki[readgr].uuw != skwazi[writeskvcnt].gruntiki[readgr].upw)
                    {
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            upwtext(skwazi[writeskvcnt].skwax, thislevel, double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw).ToString("0.00"), skwazi[writeskvcnt].start);
                            ifanywaterhere++;
                        };
                    }
                }

                /*if (ifanywaterhere == 0)
                {
                    double viziblebottomy = skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor;
                    verttext((210 * writeskvcnt + 175), (222 + viziblebottomy) / 2 - 3, "Воды");
                    verttext((210 * writeskvcnt + 178), (222 + viziblebottomy) / 2 - 2, "нет");
                }*/
            }
            #endregion drawall
            #region task wodavpeske
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                for (int readgr = 0; readgr < skwazi[writeskvcnt].layerscount; readgr++)
                {
                    if (skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("PESOK") || skwazi[writeskvcnt].gruntiki[readgr].shtrih.Contains("pesok"))
                    {
                        //search uuw ??upw?  in ALL GRUNTS or only this Grunt layer?
                        if ((skwazi[writeskvcnt].gruntiki[readgr].upw != "") && (skwazi[writeskvcnt].gruntiki[readgr].upw != null))
                        {
                            double watertop = skwazi[writeskvcnt].gruntiki[readgr].konsiceiling;
                            double waterbot = skwazi[writeskvcnt].gruntiki[readgr].konsifloor;
                            double thislevel = skwazi[writeskvcnt].replace221 - double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) * geoscale;
                            //ne budu //if ((double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) > skwazi[writeskvcnt].gruntiki[readgr].konsiceiling)&&(double.Parse(skwazi[writeskvcnt].gruntiki[readgr].upw) < skwazi[writeskvcnt].gruntiki[readgr].konsifloor)){  };
                            //noholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID");
                            //noholekonsist(skwazi[writeskvcnt].skwax-2.35, thislevel, waterbot, "SOLID");
                            varnoholekonsist(skwazi[writeskvcnt].skwax, thislevel, waterbot, "SOLID", varpipe);
                        }
                    }
                }
            }
            #endregion task wodavpeske
            #region task palochki pod kolonkoy
            for (int writeskvcnt = 0; writeskvcnt < goodskwacount; writeskvcnt++)
            {
                palkapodkol(skwazi[writeskvcnt].skwax, skwazi[writeskvcnt].gruntiki[skwazi[writeskvcnt].layerscount - 1].konsifloor);
            }
            #endregion task palochki pod kolonkoy
            #endregion kolonki
            xlserv.Close();

            #region Horiz prodolshzhenie (no loading)
            #region horiz scale
            //#region #endregion 
            double[] lineyscaled = new double[5000];
            double[] linexscaled = new double[5000];
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                linexscaled[lineinxl] = lineforscalingx[lineinxl] * xscale;
                lineyscaled[lineinxl] = lineforscalingy[lineinxl] * yscale;
            }
            #endregion horiz scale
            #region horiz scale again add USLOV HORIZ -conditional horizon
            //#region #endregion scale again
            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                //lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhoriz * yscale;
                lineyscaled[lineinxl] = lineyscaled[lineinxl] - uhorizm * yscale;
            }
            #endregion  scale again add USLOV HORIZ -conditional horizon
            #region horiz draw scaled
            //#region #endregion draw scaled
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var pl = new Polyline();
                for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                {
                    pl.AddVertexAt(lineinxl - 1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, 0, 0);//original here --> pl.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);last here --->pl.AddVertexAt(3, new Point2d(0, 10), 0, 0, 0); pl.Closed = true;
                }

                pl.Closed = false;
                acBlkTblRec.AppendEntity(pl);
                acTrans.AddNewlyCreatedDBObject(pl, true);
                acTrans.Commit();
            }
            #endregion draw scaled

            #region     task ordinata //ися1 нужны вертикальные подписи и вертикальные палочки на отметки с комментариями даже без комментов
            int isnewpicket = 0;

            for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
            {
                if ((ordinatapart2[lineinxl] != null) && (ordinatapart2[lineinxl] != "") && (ordinatapart2[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart2[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }
                if ((ordinatapart1[lineinxl] != null) && (ordinatapart1[lineinxl] != "") && (ordinatapart1[lineinxl] != " "))
                {
                    #region text transaction
                    Transaction tr =
                      doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        // We'll add the objects to the model space
                        BlockTable bt = (BlockTable)tr.GetObject(
                            doc.Database.BlockTableId,
                            OpenMode.ForRead
                          );
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite
                          );
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = "ПК" + lineforscalingpikets[lineinxl] + "+" + lineforscalingplusos[lineinxl] + " " + ordinatapart1[lineinxl];
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, 2.2, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        btr.AppendEntity(actext);
                        tr.AddNewlyCreatedDBObject(actext, true);
                        tr.Commit();
                    }
                    #endregion text transaction
                    #region     line transaction
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        //acPoly.ColorIndex = 0;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

                //ися3 поправки ися1 теперь нужны все верт отметки даже без комментов
                //ися4 if modeconfigordinata >0 than

                if (lineinxl > 0)
                { //try isa6 
                    if (lineforscalingpikets[lineinxl] != lineforscalingpikets[lineinxl - 1]) isnewpicket = 1;
                    else isnewpicket = 0;
                }

                if (isnewpicket == 0) //
                {
                    #region ися3 верт linii даже без комментов
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.ColorIndex = 211;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);

                        var acPoly2 = new Polyline(1);
                        acPoly2.Normal = Vector3d.ZAxis;
                        acPoly2.ColorIndex = 211;
                        acPoly2.AddVertexAt(0, new Point2d(linexscaled[lineinxl], -10), 0, -1, -1);
                        acPoly2.AddVertexAt(1, new Point2d(linexscaled[lineinxl], -20), 0, -1, -1);
                        acPoly2.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly2);
                        acTrans.AddNewlyCreatedDBObject(acPoly2, true);


                        // ися2 новые3строки подвала отметки
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        actext.TextString = lineforscalingy[lineinxl].ToString();
                        actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, -9, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 1.57079;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        acBlkTblRec.AppendEntity(actext);
                        acTrans.AddNewlyCreatedDBObject(actext, true);

                        if (lineinxl > 1)
                        {//try isa6 
                            double dist_between_otmetki = lineforscalingx[lineinxl] - lineforscalingx[lineinxl - 1];
                            double distscaled = linexscaled[lineinxl] - linexscaled[lineinxl - 1];

                            MText acMText = new MText();
                            acMText.Location = new Point3d(linexscaled[lineinxl] - distscaled / 2, -19, 0);
                            acMText.Height = 2;
                            acMText.TextHeight = 2;
                            acMText.Width = 20;
                            if (distscaled < 10) { acMText.Rotation = 1.57079; acMText.Location = new Point3d(linexscaled[lineinxl] - distscaled / 2 + 1, -19 + 4, 0); }
                            else acMText.Rotation = 0;
                            acMText.Contents = dist_between_otmetki.ToString("0.00");
                            acMText.Attachment = AttachmentPoint.BottomCenter;
                            acBlkTblRec.AppendEntity(acMText);
                            acTrans.AddNewlyCreatedDBObject(acMText, true);
                            //nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                        }

                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

                if (isnewpicket == 1) //
                {
                    // текст1 = 100 - последняя плюсовка предыдущего пикета 
                    // текст2 = плюсовка нового пикета
                    #region  PICKEt line
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        //differense in hight do picket and after  //diffhppr diffxun koef
                        double diffhppr = lineyscaled[lineinxl] - lineyscaled[lineinxl - 1];
                        // текст2 = плюсовка нового пикета delim na distance
                        double diffxun = lineforscalingx[lineinxl] - lineforscalingx[lineinxl - 1];
                        double koef = lineforscalingplusos[lineinxl] / diffxun;
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        #region linia na podvale
                        var acPoly2 = new Polyline(1);             //linia na podvale 
                        acPoly2.Normal = Vector3d.ZAxis;
                        acPoly2.ColorIndex = 50;
                        acPoly2.AddVertexAt(0, new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, -10), 0, -1, -1);
                        acPoly2.AddVertexAt(1, new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, -30), 0, -1, -1);
                        acPoly2.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly2);
                        acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                        #endregion linia na podvale
                        #region     текст1 = 100 - последняя плюсовка предыдущего пикета

                        // текст1 = 100 - последняя плюсовка предыдущего пикета 
                        double distafterprevpiket = 100 - lineforscalingplusos[lineinxl - 1];


                        MText acMText = new MText();
                        acMText.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 - distafterprevpiket / 2) * xscale, -19, 0);
                        acMText.Height = 2;
                        acMText.TextHeight = 2;
                        acMText.Width = 20;
                        if ((distafterprevpiket / 2 * xscale) < 10) { acMText.Rotation = 1.57079; acMText.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 - distafterprevpiket / 2 + 1) * xscale, -19 + 4, 0); }
                        else acMText.Rotation = 0;
                        acMText.Contents = distafterprevpiket.ToString("0.00");
                        acMText.Attachment = AttachmentPoint.BottomCenter;
                        acBlkTblRec.AppendEntity(acMText);
                        acTrans.AddNewlyCreatedDBObject(acMText, true);
                        #endregion
                        #region  TextString = "ПК "

                        // ися2 новые3строки подвала отметки
                        //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        DBText actext = new DBText();
                        actext.SetDatabaseDefaults();
                        //actext.TextString = "ПК " + lineforscalingpikets[lineinxl].ToString()  ;

                        //diffhppr diffxun koef
                        //actext.TextString = "ПК " + lineforscalingpikets[lineinxl].ToString() + " otlad diffhppr:" + diffhppr + " otlad diffxun:" + diffxun + " otlad koef:" + koef;
                        actext.TextString = "ПК " + lineforscalingpikets[lineinxl].ToString();

                        actext.Position = new Point3d(lineforscalingpikets[lineinxl] * xscale * 100, -29, 0);
                        actext.Height = 2;
                        actext.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext.Rotation = 0;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        acBlkTblRec.AppendEntity(actext);
                        acTrans.AddNewlyCreatedDBObject(actext, true);
                        #endregion

                        #region linia piketa
                        var acPoint = new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(lineforscalingpikets[lineinxl] * xscale * 100, lineyscaled[lineinxl] - diffhppr * koef), 0, -1, -1);
                        acPoly.ColorIndex = 50;
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);
                        #endregion linia piketa
                        #region текст2 = плюсовка нового пикета

                        /*
                        DBText actext2 = new DBText();
                        actext2.SetDatabaseDefaults();
                        actext2.TextString = lineforscalingplusos[lineinxl].ToString("0.00");
                        actext2.Position = new Point3d((lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl]/2) * xscale , -19, 0);
                        actext2.Height = 2;
                        actext2.Thickness = 22;
                        //actext.ColorIndex = 0;
                        actext2.Rotation = 0;
                        //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                        acBlkTblRec.AppendEntity(actext2);
                        acTrans.AddNewlyCreatedDBObject(actext2, true);
*/

                        MText acMText2 = new MText();
                        acMText2.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl] / 2) * xscale, -19, 0);
                        acMText2.Height = 2;
                        acMText2.TextHeight = 2;
                        acMText2.Width = 20;
                        if ((lineforscalingplusos[lineinxl] / 2 * xscale) < 10) { acMText2.Rotation = 1.57079; acMText2.Location = new Point3d((lineforscalingpikets[lineinxl] * 100 + lineforscalingplusos[lineinxl] / 2 + 1) * xscale, -19 + 4, 0); }
                        else acMText2.Rotation = 0;
                        acMText2.Contents = lineforscalingplusos[lineinxl].ToString("0.00");
                        acMText2.Attachment = AttachmentPoint.BottomCenter;
                        acBlkTblRec.AppendEntity(acMText2);
                        acTrans.AddNewlyCreatedDBObject(acMText2, true);
                        #endregion
                        /*
                        #region 
                        #endregion
                        */
                        acTrans.Commit();
                    }
                    #endregion  line transaction
                    #region линия сразу после пикета
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var acPoint = new Point2d(linexscaled[lineinxl], 0);
                        var acPoly = new Polyline(1);
                        acPoly.Normal = Vector3d.ZAxis;
                        acPoly.ColorIndex = 211;
                        acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                        acPoly.AddVertexAt(1, new Point2d(linexscaled[lineinxl], lineyscaled[lineinxl]), 0, -1, -1);
                        acPoly.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);

                        var acPoly2 = new Polyline(1);
                        acPoly2.Normal = Vector3d.ZAxis;
                        acPoly2.ColorIndex = 211;
                        acPoly2.AddVertexAt(0, new Point2d(linexscaled[lineinxl], -10), 0, -1, -1);
                        acPoly2.AddVertexAt(1, new Point2d(linexscaled[lineinxl], -20), 0, -1, -1);
                        acPoly2.Closed = false;
                        acBlkTblRec.AppendEntity(acPoly2);
                        acTrans.AddNewlyCreatedDBObject(acPoly2, true);


                        /* нужно разбить на две части
                         * 
                                                // ися2 новые3строки подвала отметки
                                                //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                                                DBText actext = new DBText();
                                                actext.SetDatabaseDefaults();
                                                actext.TextString = lineforscalingy[lineinxl].ToString();
                                                actext.Position = new Point3d(linexscaled[lineinxl] - 0.5, -70, 0);
                                                actext.Height = 2;
                                                actext.Thickness = 22;
                                                //actext.ColorIndex = 0;
                                                actext.Rotation = 1.57079;
                                                //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                                                acBlkTblRec.AppendEntity(actext);
                                                acTrans.AddNewlyCreatedDBObject(actext, true);
                                                */

                        /*
                                                DBText actext2 = new DBText();
                                                actext2.SetDatabaseDefaults();

                                                double dist_between_otmetki = lineforscalingx[lineinxl] - lineforscalingx[lineinxl - 1];
                                                actext2.TextString = dist_between_otmetki.ToString("0.00");

                                                double distscaled = linexscaled[lineinxl] - linexscaled[lineinxl - 1];
                                                actext2.Position = new Point3d(linexscaled[lineinxl] - distscaled / 2, -90, 0);
                                                actext2.Height = 2;
                                                actext2.Thickness = 22;
                                                //actext.ColorIndex = 0;
                                                if (distscaled < 30) actext2.Rotation = 1.57079;
                                                else actext2.Rotation = 0;
                                                //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                                                acBlkTblRec.AppendEntity(actext2);
                                                acTrans.AddNewlyCreatedDBObject(actext2, true);
                                      */

                        acTrans.Commit();
                    }
                    #endregion  line transaction
                }

            }
            #endregion  task ordinata


            #region boxunderground
            double firstx = linexscaled[1];
            double firsty = lineyscaled[1];
            double lastx = linexscaled[kartacounter - 1];
            double lastpprx = linexscaled[kartacounter - 1];
            double lastrlx = lineforscalingx[kartacounter - 1];
            double lasty = lineyscaled[kartacounter - 1];

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(3);
                acPoly.AddVertexAt(0, new Point2d(firstx, firsty), 0, -1, -1);  //do ok27 acPoly.AddVertexAt(0, new Point2d(firstx, firsty)   , 0, -1, -1); 
                acPoly.AddVertexAt(1, new Point2d(0, 0), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, lasty), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion boxunderground
            int podvalcount = 1;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметки поверхности земли", -83, -9, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 2;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояния между отметками", -83, -19, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 3;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Пикетаж", -83, -29, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 4;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Номеp скважины", -83, -39, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 5;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Отметка устья, м", -83, -49, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 6;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Глубина, м", -83, -59, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 7;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Расстояние, м", -83, -69, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval
            podvalcount = 8;
            #region box1podval
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(3);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(lastx, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(lastx, 0), 0, -1, -1);
                acPoly.Closed = false;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            lefttext("Дата пpоходки", -83, -79, 3);
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline();
                acPoly.AddVertexAt(0, new Point2d(-85, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(0, -podvalcount * 10), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(0, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.AddVertexAt(3, new Point2d(-85, -podvalcount * 10 + 10), 0, -1, -1);
                acPoly.Closed = true;

                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
            #endregion box1podval

            string scalextext = "Масштаб по оси X 1 : " + xdevidor;
            lefttext(scalextext, -65, 12, 2);
            string scaleytext = "Масштаб по оси Y 1 : " + ydevidor;
            lefttext(scaleytext, -65, 7.5, 2);
            string geoscaletext = "Масштаб геол. 1 : " + geodevidor;
            lefttext(geoscaletext, -65, 3, 2);

            #endregion
            lineikaok26(maximgeoline, uhorizm, yscale);


            #region isa14 turn left and right on the road
            string debuganswertwist = readconfig(@"povorot = ");//ВУ ВУ- a= внутренние углы на трассе
            //MgdAcApplication.ShowAlertDialog("kontent of включить повороты " + debuganswertwist);
            if (debuganswertwist == "yes")
            {
                #region     find all WU
                int wucounter = 0;
                povorot[] povorotiki = new povorot[100];
                //ordinatapart2[lineinxl]  ordinatabothparts[lineinxl]
                for (int lineinxl = 1; lineinxl < kartacounter; lineinxl++) //todo configit or autodetect
                {
                    if (ordinatabothparts[lineinxl].Contains(" a="))
                    {
                        wucounter++;
                        povorotiki[wucounter] = new povorot();
                        povorotiki[wucounter].cellseredini = lineinxl;
                        povorotiki[wucounter].ugol = ordinatabothparts[lineinxl];
                        povorotiki[wucounter].textformod = ordinatabothparts[lineinxl];
                        povorotiki[wucounter].textformod = povorotiki[wucounter].textformod.Replace("а= -", "У-").Replace("a= +", "У-").Replace("а= -", "У-").Replace("a= +", "У-");
                        povorotiki[wucounter].textformod = povorotiki[wucounter].textformod.Replace("а=-", "У-").Replace("a=+", "У-").Replace("а=-", "У-").Replace("a=+", "У-");
                        povorotiki[wucounter].textformod = povorotiki[wucounter].textformod.Replace("a=", "У").Replace("a =", "У");
                        povorotiki[wucounter].textformod = povorotiki[wucounter].textformod.Replace("(", "").Replace(")", "");
                    }
                }
                //MgdAcApplication.ShowAlertDialog("Naidena wu kolwo = "+ wucounter);
                #endregion  find all WU
                #region     find all "нк"  "кк"
                if (wucounter == 1)
                {
                    #region     first wu
                    for (int scanforbgnpovorota = 1; scanforbgnpovorota < povorotiki[1].cellseredini; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "нк")
                        {
                            povorotiki[1].cellbegin = scanforbgnpovorota;
                            povorotiki[1].beginx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[1].beginppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = 1" );
                            //MgdAcApplication.ShowAlertDialog("kurrent beginx  = " + povorotiki[1].beginx);
                        }
                    }

                    for (int scanforbgnpovorota = povorotiki[wucounter].cellseredini; scanforbgnpovorota < kartacounter; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "кк")
                        {

                            povorotiki[wucounter].cellend = scanforbgnpovorota;
                            povorotiki[wucounter].endx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[wucounter].endppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = " + wucounter);
                            //MgdAcApplication.ShowAlertDialog("kurrent кк x  = " + povorotiki[wucounter].endx);
                        }
                    }
                    #endregion  f wu
                }

                if (wucounter == 2)
                {
                    #region     first wu
                    for (int scanforbgnpovorota = 1; scanforbgnpovorota < povorotiki[1].cellseredini; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "нк")
                        {
                            povorotiki[1].cellbegin = scanforbgnpovorota;
                            povorotiki[1].beginx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[1].beginppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = 1");
                            // MgdAcApplication.ShowAlertDialog("kurrent beginx  = " + povorotiki[1].beginx);
                        }
                    }

                    for (int scanforbgnpovorota = povorotiki[1].cellseredini; scanforbgnpovorota < povorotiki[2].cellseredini; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "кк")
                        {
                            povorotiki[1].cellend = scanforbgnpovorota;
                            povorotiki[1].endx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[1].endppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = 1" );
                            //MgdAcApplication.ShowAlertDialog("kurrent кк x  = " + povorotiki[1].endx);
                        }
                    }
                    #endregion  f wu
                    #region     2st wu
                    for (int scanforbgnpovorota = povorotiki[1].cellseredini; scanforbgnpovorota < povorotiki[2].cellseredini; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "нк")
                        {
                            povorotiki[2].cellbegin = scanforbgnpovorota;
                            povorotiki[2].beginx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[2].beginppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = 2");
                            //MgdAcApplication.ShowAlertDialog("kurrent beginx  = " + povorotiki[2].beginx);
                        }
                    }

                    for (int scanforbgnpovorota = povorotiki[2].cellseredini; scanforbgnpovorota < kartacounter; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "кк")
                        {
                            povorotiki[2].cellend = scanforbgnpovorota;
                            povorotiki[2].endx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[2].endppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = 2");
                            //MgdAcApplication.ShowAlertDialog("kurrent кк x  = " + povorotiki[2].endx);
                        }
                    }
                    #endregion 2st wu
                }

                if (wucounter > 2)
                {
                    #region     first wu
                    for (int scanforbgnpovorota = 1; scanforbgnpovorota < povorotiki[1].cellseredini; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "нк")
                        {
                            povorotiki[1].cellbegin = scanforbgnpovorota;
                            povorotiki[1].beginx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[1].beginppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = 1");
                            //MgdAcApplication.ShowAlertDialog("kurrent beginx  = " + povorotiki[1].beginx);
                        }
                    }

                    for (int scanforbgnpovorota = povorotiki[1].cellseredini; scanforbgnpovorota < povorotiki[2].cellseredini; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "кк")
                        {
                            povorotiki[1].cellend = scanforbgnpovorota;
                            povorotiki[1].endx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[1].endppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = 1");
                            //MgdAcApplication.ShowAlertDialog("kurrent кк x  = " + povorotiki[1].endx);
                        }
                    }
                    #endregion  f wu

                    for (int curwu = 2; curwu < wucounter; curwu++)
                    {
                        //int nextwu = 3;
                        #region midst wu
                        for (int scanforbgnpovorota = povorotiki[curwu - 1].cellseredini; scanforbgnpovorota < povorotiki[curwu].cellseredini; scanforbgnpovorota++)
                        {
                            string nkprepare = ordinatabothparts[scanforbgnpovorota];
                            nkprepare = nkprepare.Trim();
                            nkprepare = nkprepare.ToLower();
                            if (nkprepare == "нк")
                            {
                                povorotiki[curwu].cellbegin = scanforbgnpovorota;
                                povorotiki[curwu].beginx = lineforscalingx[scanforbgnpovorota];
                                povorotiki[curwu].beginppr = linexscaled[scanforbgnpovorota];
                                //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = " + curwu);
                                //MgdAcApplication.ShowAlertDialog("kurrent beginx  = " + povorotiki[curwu].beginx);
                            }
                        }

                        for (int scanforbgnpovorota = povorotiki[curwu].cellseredini; scanforbgnpovorota < povorotiki[curwu + 1].cellseredini; scanforbgnpovorota++)
                        {
                            string nkprepare = ordinatabothparts[scanforbgnpovorota];
                            nkprepare = nkprepare.Trim();
                            nkprepare = nkprepare.ToLower();
                            if (nkprepare == "кк")
                            {
                                povorotiki[curwu].cellend = scanforbgnpovorota;
                                povorotiki[curwu].endx = lineforscalingx[scanforbgnpovorota];
                                povorotiki[curwu].endppr = linexscaled[scanforbgnpovorota];
                                //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = " + curwu);
                                //MgdAcApplication.ShowAlertDialog("kurrent кк x  = " + povorotiki[curwu].endx);
                            }
                        }
                        #endregion midst wu
                    }

                    #region     3st wu
                    for (int scanforbgnpovorota = povorotiki[wucounter - 1].cellseredini; scanforbgnpovorota < povorotiki[wucounter].cellseredini; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "нк")
                        {
                            povorotiki[wucounter].cellbegin = scanforbgnpovorota;
                            povorotiki[wucounter].beginx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[wucounter].beginppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = " + wucounter);
                            //MgdAcApplication.ShowAlertDialog("kurrent beginx  = " + povorotiki[wucounter].beginx);
                        }
                    }

                    for (int scanforbgnpovorota = povorotiki[wucounter].cellseredini; scanforbgnpovorota < kartacounter; scanforbgnpovorota++)
                    {
                        string nkprepare = ordinatabothparts[scanforbgnpovorota];
                        nkprepare = nkprepare.Trim();
                        nkprepare = nkprepare.ToLower();
                        if (nkprepare == "кк")
                        {
                            povorotiki[wucounter].cellend = scanforbgnpovorota;
                            povorotiki[wucounter].endx = lineforscalingx[scanforbgnpovorota];
                            povorotiki[wucounter].endppr = linexscaled[scanforbgnpovorota];
                            //MgdAcApplication.ShowAlertDialog("kurrent wucounter  = " + wucounter);
                            // MgdAcApplication.ShowAlertDialog("kurrent кк x  = " + povorotiki[wucounter].endx);
                        }
                    }
                    #endregion 3st wu
                    //for (int pereborwushek = 2; pereborwushek < wucounter; pereborwushek++) { }

                }

                #endregion  load WU

                #region     DraW Wu
                int yplana = -100;
                #region	find	left right
                for (int curwu = 1; curwu < wucounter; curwu++)
                {
                    // MgdAcApplication.ShowAlertDialog("kontent of WU " + wucounter + "is " + povorotiki[curwu].ugol);

                    if (povorotiki[curwu].ugol.Contains("a=-") || povorotiki[curwu].ugol.Contains("a= -")) { povorotiki[curwu].direction = "left"; }
                    else povorotiki[curwu].direction = "right";

                    if (povorotiki[curwu].beginx == 0) //aisa15 ap15 task sharp angles
                    {
                        povorotiki[curwu].beginx = lineforscalingx[povorotiki[curwu].cellseredini];
                        povorotiki[curwu].endx = lineforscalingx[povorotiki[curwu].cellseredini];
                        povorotiki[curwu].beginppr = linexscaled[povorotiki[curwu].cellseredini];
                        povorotiki[curwu].endppr = linexscaled[povorotiki[curwu].cellseredini];
                    };
                    if (povorotiki[curwu].endx == 0) //aisa15 ap15 task sharp angles
                    {
                        povorotiki[curwu].beginx = lineforscalingx[povorotiki[curwu].cellseredini];
                        povorotiki[curwu].endx = lineforscalingx[povorotiki[curwu].cellseredini];
                        povorotiki[curwu].beginppr = linexscaled[povorotiki[curwu].cellseredini];
                        povorotiki[curwu].endppr = linexscaled[povorotiki[curwu].cellseredini];
                    };
                }
                #endregion	left right
                #region		draw Between Wu
                //if we have wu at all 
                //part of line before first wu
                if (wucounter > 0)
                {
                    nigmaline(0, -100, povorotiki[1].beginppr, -100);
                    nigmatext(povorotiki[1].beginppr / 2, -97, (povorotiki[1].beginx).ToString("0.00"));
                };

                if (wucounter > 1)
                {
                    for (int curwu = 2; curwu <= wucounter; curwu++)
                    {
                        nigmaline(povorotiki[curwu - 1].endppr, -100, povorotiki[curwu].beginppr, -100);
                        nigmatext((povorotiki[curwu - 1].endppr + povorotiki[curwu].beginppr) / 2, -97, (povorotiki[curwu].beginx - povorotiki[curwu - 1].endx).ToString("0.00"));
                    }
                }
                //part of lineafter last wu
                if (wucounter > 0) nigmaline(povorotiki[wucounter].endppr, -100, lastx, -100);
                #endregion	draw Between Wu
                #region		draw if Plavno povorot
                if (wucounter > 0)
                {
                    for (int curwu = 1; curwu <= wucounter; curwu++)
                    {
                        if (povorotiki[curwu].beginppr!=povorotiki[curwu].endppr)
                        {
                            nigmaline(povorotiki[curwu].beginppr, yplana - 5, povorotiki[curwu].beginppr, yplana + 10);
                            nigmatext(povorotiki[curwu].beginppr - 3, yplana + 6, (povorotiki[curwu].beginx % 100).ToString("0.00"), "vertikal");
                            nigmaline(povorotiki[curwu].endppr, yplana - 5, povorotiki[curwu].endppr, -85);
                            nigmatext(povorotiki[curwu].endppr - 3, yplana + 6, (povorotiki[curwu].endx % 100).ToString("0.00"), "vertikal");
                            nigmatext((povorotiki[curwu].beginppr + povorotiki[curwu].endppr) / 2, yplana + 8, povorotiki[curwu].textformod);

                            if (povorotiki[curwu].direction == "left")
                            {
                                nigmaline(povorotiki[curwu].beginppr, yplana + 10, povorotiki[curwu].endppr, yplana + 10);
                            }
                            if (povorotiki[curwu].direction == "right")
                            {
                                nigmaline(povorotiki[curwu].beginppr, yplana - 5, povorotiki[curwu].endppr, yplana - 5);
                            }
                        }
                        else //isa15 ap15 task sharp angles
                        {
                            if (povorotiki[curwu].direction == "left")
                            {
                                nigmaline(povorotiki[curwu].beginppr, yplana + 20, povorotiki[curwu].beginppr, yplana + 6.2);
                                nigmaline(povorotiki[curwu].beginppr, yplana + 6.2, povorotiki[curwu].beginppr - 6.2, yplana);
                                nigmaline(povorotiki[curwu].beginppr, yplana + 6.2, povorotiki[curwu].beginppr + 6.2, yplana);
                                nigmatext(povorotiki[curwu].endppr - 3, yplana + 11, (povorotiki[curwu].endx % 100).ToString("0.00"), "vertikal");
                                nigmatext(povorotiki[curwu].endppr + 1, yplana + 11, (100 - povorotiki[curwu].endx % 100).ToString("0.00"), "vertikal");
                                nigmatext((povorotiki[curwu].beginppr + povorotiki[curwu].endppr) / 2, yplana - 1, /*"left" +*/ povorotiki[curwu].textformod);
                            }
                            if (povorotiki[curwu].direction == "right")
                            {
                                nigmatext((povorotiki[curwu].beginppr + povorotiki[curwu].endppr) / 2, yplana + 12, /*"right" +*/ povorotiki[curwu].textformod);
                                nigmatext(povorotiki[curwu].endppr - 3, yplana - 11, (povorotiki[curwu].endx % 100).ToString("0.00"), "vertikal");
                                nigmatext(povorotiki[curwu].endppr + 1, yplana - 11, (100 - povorotiki[curwu].endx % 100).ToString("0.00"), "vertikal");
                                nigmaline(povorotiki[curwu].beginppr, yplana - 20, povorotiki[curwu].beginppr, yplana - 6.2);
                                nigmaline(povorotiki[curwu].beginppr, yplana - 6.2, povorotiki[curwu].beginppr - 6.2, yplana);
                                nigmaline(povorotiki[curwu].beginppr, yplana - 6.2, povorotiki[curwu].beginppr + 6.2, yplana);
                            }
                        }
                    }
                }

                lefttext("Элементы плана", -63, yplana - 3, 3);//Элементы плана
                //nigmatext(-83, yplana +2, "Элементы плана");//Элементы плана
                #endregion	draw if Plavno povorot
                #endregion  DraW Wu
                #region     podvalkm
                int ypodvalakm = -100;//ap15 same to plan elements

                double reallifekm = lastrlx / 1000;
                int skokokm = Convert.ToInt32(reallifekm);

                for (int curkm = 1; curkm <= skokokm; curkm++)
                {
                    double xforcircle = curkm * 1000 * xscale;
                    nigmacirc(xforcircle, ypodvalakm - 5);
                    nigmatext(xforcircle, ypodvalakm - 16, curkm.ToString());
                    nigmaline(xforcircle, ypodvalakm-5,xforcircle, ypodvalakm+20);
                }

                //STATIC
                nigmacirchalf(0, ypodvalakm);
                nigmatext(1.5, ypodvalakm - 16, "0");

                lefttext("Километры", -63, ypodvalakm - 5, 3);//Элементы плана
                //nigmatext(-83, ypodvalakm - 13, "Километры");
                nigmaline(0, 0, 0, ypodvalakm - 20);//vertikalbar
                nigmaline(-85, 0, -85, ypodvalakm - 20);//vertikalbar
                //if it is most bottom floor ----->>>>>>>>>>> than draw this final border over Elements  of plan
                nigmaline(-85, ypodvalakm - 20, 0, ypodvalakm - 20);//finalbar
                nigmaline(0, ypodvalakm - 20, lastpprx, ypodvalakm - 20);//finalbar
                nigmaline(lastpprx, 0, lastpprx, ypodvalakm - 20);//finalbar
                #endregion  podvalkm
                #region     podvalpk
                int ypodvalapk = -83;

                double reallifepk = lastrlx / 100;
                int skokopk = Convert.ToInt32(reallifepk);

                for (int curpk = 1; curpk <= skokopk; curpk++)
                {
                    double xforpk = curpk * 100 * xscale;
                    nigmatext(xforpk, ypodvalapk, "ПК " + curpk.ToString());
                    nigmaline(xforpk, ypodvalapk + 0.5, xforpk, ypodvalapk + 3);
                }
                //static
                nigmatext(0, ypodvalapk, "ПК 0");

                lefttext("Пикет", -63, ypodvalapk - 3, 3);//Элементы плана
                //nigmatext(-83, ypodvalapk - 13, "Пикет");//Элементы плана
                //if it is last podval than do this
                //nigmaline(0, 0, 0, ypodvalapk - 40);//vertikalbar
                //nigmaline(-85, 0, -85, ypodvalapk - 40);//vertikalbar
                //nigmaline(-85, ypodvalakm - 20, 0, ypodvalakm - 20);//finalbar
                #endregion  podvalpk
            }
            #endregion isa14
        }


        [CommandMethod("geolog_статика_odin_grafik")]
        public void geolog_static_graphic_xy()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region     ask user
            int xlpage =      Convert.ToInt32(askpageint());
            int xlstartline = Convert.ToInt32(asklineint());
            int xlcol =       Convert.ToInt32(askcolint());
            #endregion  ask user

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "xls; xlsx",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);

            List<statmark> statmarks = new List<statmark>();
            List<string> landy = new List<string>();

            int probelscounter = 0;

            for (; xlstartline < 5500; xlstartline++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(xlpage, xlstartline, 1, ref cellcontent);//sheet row col

                if ((cellcontent == "") || (cellcontent == null))
                {
                    polifromlist(statmarks);
                    landy.Clear();
                    statmarks.Clear();

                    probelscounter++;
                    if(probelscounter>5) break;
                }
                else 
                {
                    probelscounter = 0;
                        landy.Add(cellcontent); string cellcontent2 = " "; xlserv2.UniRead(xlpage, xlstartline, xlcol, ref cellcontent2);// string cellcontent3 = " "; xlserv2.UniRead(xlpage, xlstartline, 3, ref cellcontent3);
                    statmarks.Add(new statmark() { marky = cellcontent , markx1 = cellcontent2 });
                }

            }

            xlserv2.Close();
        }
        [CommandMethod("geolog_статика_CU_i_CU")]
        public void geolog_static_graphic_CU_i_CU()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            #region     ask user
            int xlpage = Convert.ToInt32(askpageint());
            int xlstartline = Convert.ToInt32(asklineint());
            //int xlcol = Convert.ToInt32(askcolint());
            #endregion  ask user

            ComplexObjects.ExcelServ xlserv2 = new ComplexObjects.ExcelServ();
            #region open xl2
            var ofd2 =
              new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select Excel spreadsheet to link",
                null,
                "xls; xlsx",
                "ExcelFileToLink",
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.
                  DoNotTransferRemoteFiles
              );

            var dr2 = ofd2.ShowDialog();

            if (dr2 != System.Windows.Forms.DialogResult.OK)
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nFile selected was \"{0}\". Contains these sheets:",
              ofd2.Filename
            );
            #endregion open xl2
            xlserv2.Open(ofd2.Filename);

            List<statmark> statmarks = new List<statmark>();
            List<statmark> statmarks2 = new List<statmark>();
            List<string> landy = new List<string>();

            int probelscounter = 0;

            for (; xlstartline < 5500; xlstartline++) //todo configit or autodetect
            {
                string cellcontent = " "; xlserv2.UniRead(xlpage, xlstartline, 1, ref cellcontent);//sheet row col

                if ((cellcontent == "") || (cellcontent == null))//pri pervom dge razrive risuet poliliniyu iz teh tochek cho nakopilis ranshe
                {
                    polifromlist(statmarks);
                    polifromlistgreen(statmarks2);
                    statmarks.Clear();
                    statmarks2.Clear();

                    probelscounter++;
                    if (probelscounter > 5) break;
                }
                else
                {
                    probelscounter = 0;
                    string cellcontent1 = " "; xlserv2.UniRead(xlpage, xlstartline, 2, ref cellcontent1); 
                    string cellcontent2 = " "; xlserv2.UniRead(xlpage, xlstartline, 4, ref cellcontent2);
                    statmarks.Add(new statmark() { marky = cellcontent, markx1 = cellcontent1 });
                    statmarks2.Add(new statmark() { marky = cellcontent, markx1 = cellcontent2 });
                }

            }

            xlserv2.Close();
        }

        public class statmark
        {
            public string marky { get; set; }
            public string markx1 { get; set; }
            public override string ToString()
            {
                return "ID: " + marky + "   Name: " + markx1;
            }
        }
        public static void polifromlist(List<statmark> inlist)
        {
            TextWriter wodafile = new StreamWriter("d:\\inlist.txt", true, Encoding.Default);
            foreach (statmark statmark in inlist)
            {
                wodafile.WriteLine(statmark);
            }
            wodafile.Close();

            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(inlist.Capacity);
                acPoly.Normal = Vector3d.ZAxis;
                int something = 0;
                int inverty = 1;
                if (statconfigyinvert()=="no")
                { inverty = -1; }

                foreach (statmark statmark in inlist)
                {
                    if (statconfigxyswitch() == "no")
                        acPoly.AddVertexAt(something, new Point2d(double.Parse(statmark.markx1) * double.Parse(statconfigxskal()), double.Parse(statmark.marky )*inverty * double.Parse(statconfigyskal())), 0, -1, -1);
                    else
                        acPoly.AddVertexAt(something, new Point2d(double.Parse(statmark.marky ) *inverty* double.Parse(statconfigyskal()), double.Parse(statmark.markx1) * double.Parse(statconfigxskal())), 0, -1, -1);
                    something++;
                }
                acPoly.Closed = false;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static void polifromlistgreen(List<statmark> inlist)
        {
            TextWriter wodafile = new StreamWriter("d:\\inlist.txt", true, Encoding.Default);
            foreach (statmark statmark in inlist)
            {
                wodafile.WriteLine(statmark);
            }
            wodafile.Close();

            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(inlist.Capacity);
                acPoly.Normal = Vector3d.ZAxis;
                int something = 0;
                int inverty = -1;
                if (statconfigyinvert() == "no")
                { inverty = 1; }
                foreach (statmark statmark in inlist)
                {
                    if (statconfigxyswitch() == "no")
                        acPoly.AddVertexAt(something, new Point2d(double.Parse(statmark.markx1) * double.Parse(statconfigxskal()), double.Parse(statmark.marky) *inverty* double.Parse(statconfigyskal())), 0, -1, -1);
                    else
                        acPoly.AddVertexAt(something, new Point2d(double.Parse(statmark.marky) *inverty*  double.Parse(statconfigyskal()), double.Parse(statmark.markx1) *double.Parse(statconfigxskal())), 0, -1, -1);
                    something++;
                }
                acPoly.Closed = false;
                acPoly.ColorIndex = 93;
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.Commit();
            }
        }
        public static string statconfigxskal()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\statconfig.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\statconfig.txt"))
                {
                    if (line.Contains("xskal = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("xskal = ", "");
                return tempzap;
            }

            else
            {
                return "1.00";
            }
        }
        public static string statconfigyskal()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\statconfig.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\statconfig.txt"))
                {
                    if (line.Contains("yskal = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("yskal = ", "");
                return tempzap;
            }

            else
            {
                return "1.00";
            }
        }
        public static string statconfigxyswitch()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\statconfig.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\statconfig.txt"))
                {
                    if (line.Contains("xyswitch = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("xyswitch = ", "");
                return tempzap;
            }

            else
            {
                return "no";
            }
        }
        public static string statconfigyinvert()
        {
            string tempzap = "";
            if (File.Exists(@"C:\AtlPProf\statconfig.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\statconfig.txt"))
                {
                    if (line.Contains("yinvert = "))
                    {
                        tempzap = line;
                    }
                }
                tempzap = tempzap.Replace("yinvert = ", "");
                return tempzap;
            }

            else
            {
                return "no";
            }
        }
        public static double askpageint()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

                PromptStringOptions pStrOpts = new PromptStringOptions("\nкакой лист ексель файла?: ");
                pStrOpts.AllowSpaces = true;
                PromptResult pStrRes = ed.GetString(pStrOpts);
                //MgdAcApplication.ShowAlertDialog("The X entered was: " + pStrRes.StringResult);
                return double.Parse(pStrRes.StringResult);
        }
        public static double asklineint()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptStringOptions pStrOpts = new PromptStringOptions("\nкакая первая строка для графика?: ");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = ed.GetString(pStrOpts);
            //MgdAcApplication.ShowAlertDialog("The X entered was: " + pStrRes.StringResult);
            return double.Parse(pStrRes.StringResult);
        }
        public static double askcolint()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptStringOptions pStrOpts = new PromptStringOptions("\nкакой столбец для ЦУ?: ");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = ed.GetString(pStrOpts);
            //MgdAcApplication.ShowAlertDialog("The X entered was: " + pStrRes.StringResult);
            return double.Parse(pStrRes.StringResult);
        }
        /*
        [CommandMethod("GetStringFromUser")]
        public static void GetStringFromUser()
        {
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;

            PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter your name: ");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);

            MgdAcApplication.ShowAlertDialog("The name entered was: " +
                                        pStrRes.StringResult);
        }
        [CommandMethod("Getdebuglevel")]
        public static int Getdebuglevel()
        {
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;

            PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter debug level: ");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);

            MgdAcApplication.ShowAlertDialog("The name entered was: " +
                                        pStrRes.StringResult);
            //int rezlevel = int.Parse(pStrRes.StringResult);
            return int.Parse(pStrRes.StringResult);
        }
        */
        public static double getopissize()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nOpisanie size?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("2");
            oPko.Keywords.Add("2.5");
            oPko.Keywords.Add("other");
            oPko.Keywords.Default = "2.5";
            PromptResult oPkr = edtor.GetKeywords(oPko);

            if (oPkr.Status == PromptStatus.OK && oPkr.StringResult == "other")
            {
                PromptStringOptions pStrOpts = new PromptStringOptions("\nOpisanie size?: ");
                pStrOpts.AllowSpaces = true;
                PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);
                MgdAcApplication.ShowAlertDialog("The name entered was: " + pStrRes.StringResult);
                return double.Parse(pStrRes.StringResult);
            }
            else
            {
                return double.Parse(oPkr.StringResult);
            }
            
            return double.Parse("0.0");
        }
        public static double getxbegin()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nBegining X ?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("0");
            oPko.Keywords.Add("other");
            oPko.Keywords.Default = "0";
            PromptResult oPkr = edtor.GetKeywords(oPko);

            if (oPkr.Status == PromptStatus.OK && oPkr.StringResult == "other")
            {
                PromptStringOptions pStrOpts = new PromptStringOptions("\nBegining X?: ");
                pStrOpts.AllowSpaces = true;
                PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);
                MgdAcApplication.ShowAlertDialog("The X entered was: " + pStrRes.StringResult);
                return double.Parse(pStrRes.StringResult);
            }
            else
            {
                return double.Parse(oPkr.StringResult);
            }

            return double.Parse("0.0");
        }
        public static double getuhorizm()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nУсловный горизонт ?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("0");
            oPko.Keywords.Add("no");
            oPko.Keywords.Default = "0";
            PromptResult oPkr = edtor.GetKeywords(oPko);

            if (oPkr.Status == PromptStatus.OK && oPkr.StringResult != "0")
            {
                PromptStringOptions pStrOpts = new PromptStringOptions("\nгоризонт(в метрах)?: ");
                pStrOpts.AllowSpaces = true;
                PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);
                //MgdAcApplication.ShowAlertDialog("The X entered was: " + pStrRes.StringResult);
                return double.Parse(pStrRes.StringResult);
                //return double.Parse(oPkr.StringResult);
            }
            else
            {
                return double.Parse(oPkr.StringResult);
            }
            return double.Parse(oPkr.StringResult);
            //return double.Parse("0.0");
        }
        public static double getxscale()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;
            PromptKeywordOptions oPko =
            new PromptKeywordOptions("\nXscale ?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("20");
            oPko.Keywords.Add("50");
            oPko.Keywords.Add("100");
            oPko.Keywords.Add("200");
            oPko.Keywords.Add("500");
            oPko.Keywords.Add("1000");
            oPko.Keywords.Add("2000");
            oPko.Keywords.Add("5000");
            oPko.Keywords.Add("no");
            oPko.Keywords.Default = "500";
            PromptResult oPkr = edtor.GetKeywords(oPko);
            return double.Parse(oPkr.StringResult);
        }
        public static double getyscale()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;
            PromptKeywordOptions oPko =
            new PromptKeywordOptions("\nYscale ?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("20");
            oPko.Keywords.Add("50");
            oPko.Keywords.Add("100");
            oPko.Keywords.Add("200");
            oPko.Keywords.Add("500");
            oPko.Keywords.Add("1000");
            oPko.Keywords.Add("2000");
            oPko.Keywords.Add("5000");
            oPko.Keywords.Default = "100";
            PromptResult oPkr = edtor.GetKeywords(oPko);
            return double.Parse(oPkr.StringResult);
        }
        public static double getgeoscale()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;
            PromptKeywordOptions oPko =
            new PromptKeywordOptions("\nGEO scale ?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("20");
            oPko.Keywords.Add("50");
            oPko.Keywords.Add("100");
            oPko.Keywords.Add("200");
            oPko.Keywords.Add("500");
            oPko.Keywords.Add("1000");
            oPko.Keywords.Add("2000");
            oPko.Keywords.Add("5000");
            oPko.Keywords.Default = "100";
            PromptResult oPkr = edtor.GetKeywords(oPko);
            return double.Parse(oPkr.StringResult);
        }
        public static string getautohor()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nAUTO Условный горизонт autoматически?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("yes");
            oPko.Keywords.Add("no");
            oPko.Keywords.Default = "yes";
            PromptResult oPkr = edtor.GetKeywords(oPko);
            return oPkr.StringResult;
        }
        public static string getxlpage()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nУкажите страницу екселя");
            oPko.AllowNone = false;
            oPko.Keywords.Add("1");
            oPko.Keywords.Add("2");
            oPko.Keywords.Add("3");
            oPko.Keywords.Add("4");
            oPko.Keywords.Add("5");
            oPko.Keywords.Add("6");
            oPko.Keywords.Add("7");
            oPko.Keywords.Add("8");
            oPko.Keywords.Add("9");
            oPko.Keywords.Default = "1";
            PromptResult oPkr = edtor.GetKeywords(oPko);
            return oPkr.StringResult;
        }
        public static string getxlstartline()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nУкажите с какой строки екселя начасть строить график");
            oPko.AllowNone = false;
            oPko.Keywords.Add("1");
            oPko.Keywords.Add("2");
            oPko.Keywords.Add("3");
            oPko.Keywords.Add("4");
            oPko.Keywords.Add("5");
            oPko.Keywords.Add("6");
            oPko.Keywords.Add("7");
            oPko.Keywords.Add("8");
            oPko.Keywords.Add("9");
            oPko.Keywords.Add("10");
            oPko.Keywords.Add("11");
            oPko.Keywords.Add("12");
            oPko.Keywords.Add("13");
            oPko.Keywords.Add("14");
            oPko.Keywords.Add("15");
            oPko.Keywords.Add("16");
            oPko.Keywords.Add("17");
            oPko.Keywords.Add("18");
            oPko.Keywords.Add("19");
            oPko.Keywords.Add("20");
            oPko.Keywords.Default = "1";
            PromptResult oPkr = edtor.GetKeywords(oPko);
            return oPkr.StringResult;
        }
        public static string getxlcol()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nУкажите столбец екселя");
            oPko.AllowNone = false;
            oPko.Keywords.Add("1");
            oPko.Keywords.Add("2");
            oPko.Keywords.Add("3");
            oPko.Keywords.Add("4");
            oPko.Keywords.Add("5");
            oPko.Keywords.Add("6");
            oPko.Keywords.Add("7");
            oPko.Keywords.Add("8");
            oPko.Keywords.Add("9");
            oPko.Keywords.Default = "1";
            PromptResult oPkr = edtor.GetKeywords(oPko);
            return oPkr.StringResult;
        }
        public static double Getegesize()
        {
            Editor edtor;
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            edtor = acDoc.Editor;

            PromptKeywordOptions oPko =
                        new PromptKeywordOptions("\nEge size?");
            oPko.AllowNone = false;
            oPko.Keywords.Add("2");
            oPko.Keywords.Add("2.5");
            oPko.Keywords.Add("other");
            oPko.Keywords.Default = "2.5";
            PromptResult oPkr = edtor.GetKeywords(oPko);

            if (oPkr.Status == PromptStatus.OK && oPkr.StringResult == "other")
            {
                PromptStringOptions pStrOpts = new PromptStringOptions("\nOpisanie size?: ");
                pStrOpts.AllowSpaces = true;
                PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);
                MgdAcApplication.ShowAlertDialog("The name entered was: " + pStrRes.StringResult);
                return double.Parse(pStrRes.StringResult);
            }
            else
            {
                return double.Parse(oPkr.StringResult);
            }

            return double.Parse("0.0");
        }
        public static int Gettextforige()
        {
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;

            PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter text size for ige: ");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);

            MgdAcApplication.ShowAlertDialog("The text size for ige: " +
                                        pStrRes.StringResult);
            //int rezlevel = int.Parse(pStrRes.StringResult);
            return int.Parse(pStrRes.StringResult);
        }
        public int nakladkaproba(double deepprobi, skwazya thisskwa)
       {
           int nakladkarez = 0;
              for (int readgr = 1; readgr < thisskwa.obrcount +1; readgr++)
                    {
                        double diff = thisskwa.obriki[readgr].bottom - 100;
                        if (diff<0) diff = diff * -1;
                        if (diff < 10) nakladkarez = nakladkarez + 1;
                    }
               return nakladkarez;
        }
        public skwazya sortproba(skwazya thisskwa)
        {
            return thisskwa;
        }
        public void mtext(double x, double y, string content)
        {
            Document doc =
              MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

                Transaction tr =
                  doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    // We'll add the objects to the model space

                    BlockTable bt =
                      (BlockTable)tr.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );

                    BlockTableRecord btr =
                      (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );

                    // Add our boundary objects to the drawing and
                    // collect their ObjectIds for later use
                    MText actext = new MText();
                    actext.SetDatabaseDefaults();
                    actext.Contents = content;
                    actext.Location = new Point3d(x,y,0);
                    actext.Width = 22;

                    btr.AppendEntity(actext);
                    tr.AddNewlyCreatedDBObject(actext, true);

                    tr.Commit();
                }
            }
        public void text(double x, double y, string content)
        {
            if ((content!="")&&(content!=null))
            {
            Document doc =
              MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Transaction tr =
              doc.TransactionManager.StartTransaction();
            using (tr)
            {
                // We'll add the objects to the model space

                BlockTable bt =
                  (BlockTable)tr.GetObject(
                    doc.Database.BlockTableId,
                    OpenMode.ForRead
                  );

                BlockTableRecord btr =
                  (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite
                  );

                //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                DBText actext = new DBText();
                actext.SetDatabaseDefaults();
                actext.TextString = content;
                actext.Position = new Point3d(x, y, 0);
                actext.Height = 2;
                actext.Thickness = 22;
                //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 

                btr.AppendEntity(actext);
                tr.AddNewlyCreatedDBObject(actext, true);

                tr.Commit();
            }
            }
        }
        public void bluetext(double x, double y, string content)
        {
            if ((content != "") && (content != null))
            {
                Document doc =
                MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                Transaction tr =
                  doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    // We'll add the objects to the model space

                    BlockTable bt =
                      (BlockTable)tr.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );

                    BlockTableRecord btr =
                      (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );

                    //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                    DBText actext = new DBText();
                    actext.SetDatabaseDefaults();
                    actext.TextString = content;
                    actext.Position = new Point3d(x-5, y, 0);
                    actext.Height = 2;
                    actext.Thickness = 22;
                    actext.ColorIndex = 5;
                    //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 

                    btr.AppendEntity(actext);
                    tr.AddNewlyCreatedDBObject(actext, true);

                    tr.Commit();
                }
            }
        }
        public void uuwtext(double x, double y, string content, string date)
        {
            if ((content != "") && (content != null))
            {
                Document doc =
                MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                Transaction tr =
                  doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    // We'll add the objects to the model space
                    BlockTable bt = (BlockTable)tr.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );
                    //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    DBText actext = new DBText();
                    actext.SetDatabaseDefaults();
                    actext.TextString = content;
                    actext.Position = new Point3d(x - 17, y+1, 0);
                    actext.Height = 2;
                    actext.Thickness = 22;
                    actext.ColorIndex = 5;
                    //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                    btr.AppendEntity(actext);
                    tr.AddNewlyCreatedDBObject(actext, true);
                    tr.Commit();
                }

                Transaction trdate =  doc.TransactionManager.StartTransaction();
                using (trdate)
                {
                    // We'll add the objects to the model space
                    BlockTable bt = (BlockTable)trdate.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );
                    BlockTableRecord btr = (BlockTableRecord)trdate.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );
                    //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    DBText actext = new DBText();
                    actext.SetDatabaseDefaults();
                    actext.TextString = date;
                    actext.Position = new Point3d(x - 19, y-3, 0);
                    actext.Height = 2;
                    actext.Thickness = 22;
                    actext.ColorIndex = 5;
                    //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                    btr.AppendEntity(actext);
                    trdate.AddNewlyCreatedDBObject(actext, true);
                    trdate.Commit();
                }

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    var acPoint = new Point2d(x - 19, y + 1.6+1);
                    var acPoly = new Polyline(3);
                    acPoly.Normal = Vector3d.ZAxis;
                    acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                    acPoly.AddVertexAt(1, new Point2d(x + 0.925 - 19, y+1), 0, -1, -1);
                    acPoly.AddVertexAt(2, new Point2d(x + 1.85 - 19, y + 1.6+1), 0, -1, -1);
                    acPoly.Closed = true;
                    // Add the rectangle to the block table record
                    acBlkTblRec.AppendEntity(acPoly);
                    acTrans.AddNewlyCreatedDBObject(acPoly, true);
                    // Create the hatch
                    var acHatch = new Hatch();
                    acHatch.PatternScale = 10;
                    acBlkTblRec.AppendEntity(acHatch);
                    acTrans.AddNewlyCreatedDBObject(acHatch, true);
                    acHatch.SetDatabaseDefaults();
                    acHatch.ColorIndex = 5;
                    acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    acHatch.Associative = true;
                    // Add the outer boundary
                    acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });
                    // Add the inner boundary
                    // Validate the hatch
                    acHatch.EvaluateHatch(true);
                    acTrans.Commit();
                }

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    var acPoint = new Point2d(x-20 ,  y );
                    var acPoly = new Polyline(1);
                    acPoly.Normal = Vector3d.ZAxis;
                    acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                    acPoly.AddVertexAt(1, new Point2d(x , y), 0, -1, -1);
                    acPoly.ColorIndex = 5;
                    acPoly.Closed = false;
                    acBlkTblRec.AppendEntity(acPoly);
                    acTrans.AddNewlyCreatedDBObject(acPoly, true);
                    acTrans.Commit();
                }
            }
        }
        public void upwtext(double x, double y, string content, string date)
        {
            if ((content != "") && (content != null))
            {
                Document doc =
                MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                Transaction tr =
                  doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    // We'll add the objects to the model space
                    BlockTable bt = (BlockTable)tr.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );
                    //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    DBText actext = new DBText();
                    actext.SetDatabaseDefaults();
                    actext.TextString = content;
                    actext.Position = new Point3d(x - 17, y+1, 0);
                    actext.Height = 2;
                    actext.Thickness = 22;
                    actext.ColorIndex = 5;
                    //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                    btr.AppendEntity(actext);
                    tr.AddNewlyCreatedDBObject(actext, true);
                    tr.Commit();
                }

                Transaction trdate = doc.TransactionManager.StartTransaction();
                using (trdate)
                {
                    // We'll add the objects to the model space
                    BlockTable bt = (BlockTable)trdate.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );
                    BlockTableRecord btr = (BlockTableRecord)trdate.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );
                    //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    DBText actext = new DBText();
                    actext.SetDatabaseDefaults();
                    actext.TextString = date;
                    actext.Position = new Point3d(x - 19, y-3, 0);
                    actext.Height = 2;
                    actext.Thickness = 22;
                    actext.ColorIndex = 5;
                    //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 
                    btr.AppendEntity(actext);
                    trdate.AddNewlyCreatedDBObject(actext, true);
                    trdate.Commit();
                }

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    var acPoint = new Point2d(x - 19, y + 1.6+1);
                    var acPoly = new Polyline(3);
                    acPoly.Normal = Vector3d.ZAxis;
                    acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                    acPoly.AddVertexAt(1, new Point2d(x + 0.925 - 19, y+1), 0, -1, -1);
                    acPoly.AddVertexAt(2, new Point2d(x + 1.85 - 19, y + 1.6+1), 0, -1, -1);
                    acPoly.ColorIndex = 5;
                    acPoly.Closed = true;
                    // Add the rectangle to the block table record
                    acBlkTblRec.AppendEntity(acPoly);
                    acTrans.AddNewlyCreatedDBObject(acPoly, true);
                    // Create the hatch
                    acTrans.Commit();
                }

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    var acPoint = new Point2d(x - 20, y);
                    var acPoly = new Polyline(1);
                    acPoly.Normal = Vector3d.ZAxis;
                    acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                    acPoly.AddVertexAt(1, new Point2d(x , y), 0, -1, -1);
                    acPoly.ColorIndex = 5;
                    acPoly.Closed = false;
                    acBlkTblRec.AppendEntity(acPoly);
                    acTrans.AddNewlyCreatedDBObject(acPoly, true);
                    acTrans.Commit();
                }
            }
        }
        public void undertext(double x, double y, string content)
        {//ися13 у подошвы текст с подставочкой и подставочка не касается скважины а надо чтоб касалось
            if ((content != "") && (content != null))
            {
                Document doc =
                  MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                Transaction tr =
                  doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    // We'll add the objects to the model space

                    BlockTable bt =
                      (BlockTable)tr.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );

                    BlockTableRecord btr =
                      (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );

                    //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                    DBText actext = new DBText();
                    actext.SetDatabaseDefaults();
                    actext.TextString = content;
                    actext.Position = new Point3d(x, y+1, 0);
                    actext.Height = 2;
                    actext.Thickness = 22;
                    //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 

                    btr.AppendEntity(actext);
                    tr.AddNewlyCreatedDBObject(actext, true);

                    tr.Commit();
                }
                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    var acPoly2 = new Polyline(1);
                    acPoly2.Normal = Vector3d.ZAxis;
                    acPoly2.AddVertexAt(0, new Point2d(x - 4, y), 0, -1, -1);//ися13 у подошвы текст с подставочкой и подставочка не касается скважины а надо чтоб касалось
                    acPoly2.AddVertexAt(1, new Point2d(x + 11, y), 0, -1, -1);
                    acPoly2.Closed = false;
                    //acPoly2.ColorIndex = 2;
                    acBlkTblRec.AppendEntity(acPoly2);
                    acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                    acTrans.Commit();
                }
            }
        }
        public void leftgeotxt(double x, double y, string content)
        {
            if ((content != "") && (content != null))
            {
                Document doc =
                  MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;     acBlkTbl = acTrans.GetObject(db.BlockTableId,  OpenMode.ForRead) as BlockTable;
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;   acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],     OpenMode.ForWrite) as BlockTableRecord;

                double nexttextceileng = 0.0;
                // Create a multiline text object
                using (MText acMText = new MText())
                {
                    acMText.Location = new Point3d(x-3, y+1, 0);
                    acMText.Width = 20;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.BottomRight;
                    acMText.TextHeight = 2;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                }
                acTrans.Commit();
                }
            }
        }
        public void verttext(double x, double y, string content)
        {
            if ((content != "") && (content != null))
            {
                Document doc =
                  MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                Transaction tr =
                  doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    // We'll add the objects to the model space

                    BlockTable bt =
                      (BlockTable)tr.GetObject(
                        doc.Database.BlockTableId,
                        OpenMode.ForRead
                      );

                    BlockTableRecord btr =
                      (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                      );

                    //TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                    DBText actext = new DBText();
                    actext.SetDatabaseDefaults();
                    actext.TextString = content;
                    actext.Position = new Point3d(x, y, 0);
                    actext.Height = 2;
                    actext.Thickness = 22;
                    actext.Rotation = 1.57;
                    //actext.VerticalMode = TextVerticalMode.TextVerticalMid;it gives me text in ther coordinate start=beginning
                    //actext.TextStyleId = newTextStyleTable.GetField("_CGEI"); 

                    btr.AppendEntity(actext);
                    tr.AddNewlyCreatedDBObject(actext, true);

                    tr.Commit();
                }
            }
        }
        public double MmText(string content, double mmx, double ceil)//AttachmentPoint.TopCenter;
        {
            // Get the current document and database
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                double nexttextceileng = 0.0;
                // Create a multiline text object
                using (MText acMText = new MText())
                {
                    acMText.Location = new Point3d(mmx, ceil, 0);
                    acMText.Width = 84;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    //acMText.TextHeight = 2.5;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                }
                acTrans.Commit();

                return nexttextceileng;
            }
        }
        public double MTextbymenu(string content, double mmx, double ceil, double size)//AttachmentPoint.TopCenter;
        {
            // Get the current document and database
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                double nexttextceileng = 0.0;
                // Create a multiline text object
                using (MText acMText = new MText())
                {
                    acMText.Location = new Point3d(mmx, ceil, 0);
                    acMText.Width = 84;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    acMText.TextHeight = size;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                }
                acTrans.Commit();

                return nexttextceileng;
            }
        }
        public double lefttext(string content, double mmx, double ceil, double size)//AttachmentPoint.TopCenter;
        {
            // Get the current document and database
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                double nexttextceileng = 0.0;
                // Create a multiline text object
                using (MText acMText = new MText())
                {
                    acMText.Location = new Point3d(mmx, ceil, 0);
                    acMText.Width = 84;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.BottomLeft;
                    acMText.TextHeight = size;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                }
                acTrans.Commit();

                return nexttextceileng;
            }
        }
        public double Textafterdot(string content, double mmx, double ceil)//AttachmentPoint.TopCenter;
        {
            content = content.Replace(",", ".");
            //if (!content.Contains(".")) { content = content + ".00"; }
            double tempcontent = double.Parse(content);

            // Get the current document and database
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                double nexttextceileng = 0.0;
                // Create a multiline text object
                using (MText acMText = new MText())
                {
                    acMText.Location = new Point3d(mmx, ceil, 0);
                    acMText.Width = 84;
                    acMText.Contents = tempcontent.ToString("0.00");
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    acMText.TextHeight = 2.5;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                }
                acTrans.Commit();

                return nexttextceileng;
            }
        }
        public double igemtext(string content, double mmx, double ceil,double textheight)//AttachmentPoint.TopCenter;
        {
            // Get the current document and database
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                double nexttextceileng = 0.0;
                // Create a multiline text object
                using (MText acMText = new MText())
                {
                    acMText.Location = new Point3d(mmx, ceil, 0);
                    acMText.Width = 84;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    //acMText.Height = textheight;
                    acMText.TextHeight = textheight;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                }
                acTrans.Commit();

                return nexttextceileng;
            }
        }
        public double snegovik(string content, double xxxx, double yyyy)//AttachmentPoint.TopCenter;
        {
            // Get the current document and database
            //yyyy -= 2.5;

            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly2 = new Polyline(1);
                acPoly2.Normal = Vector3d.ZAxis;
                acPoly2.AddVertexAt(0, new Point2d(xxxx, yyyy-1), 0, -1, -1);
                acPoly2.AddVertexAt(1, new Point2d(xxxx, yyyy+1), 0, -1, -1);
                acPoly2.Closed = false;
                acBlkTblRec.AppendEntity(acPoly2);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                acTrans.Commit();
            }
            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly2 = new Polyline(1);
                acPoly2.Normal = Vector3d.ZAxis;
                acPoly2.AddVertexAt(0, new Point2d(xxxx-1, yyyy), 0, -1, -1);
                acPoly2.AddVertexAt(1, new Point2d(xxxx+1, yyyy), 0, -1, -1);
                acPoly2.Closed = false;
                acBlkTblRec.AppendEntity(acPoly2);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                acTrans.Commit();
            }
            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly2 = new Polyline(1);
                acPoly2.Normal = Vector3d.ZAxis;
                acPoly2.AddVertexAt(0, new Point2d(xxxx - 0.7, yyyy-0.7), 0, -1, -1);
                acPoly2.AddVertexAt(1, new Point2d(xxxx + 0.7, yyyy+0.7), 0, -1, -1);
                acPoly2.Closed = false;
                acBlkTblRec.AppendEntity(acPoly2);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                acTrans.Commit();
            }

            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly2 = new Polyline(1);
                acPoly2.Normal = Vector3d.ZAxis;
                acPoly2.AddVertexAt(0, new Point2d(xxxx - 0.7, yyyy + 0.7), 0, -1, -1);
                acPoly2.AddVertexAt(1, new Point2d(xxxx + 0.7, yyyy - 0.7), 0, -1, -1);
                acPoly2.Closed = false;
                acBlkTblRec.AppendEntity(acPoly2);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                acTrans.Commit();
            }


            return 3.43;
        }
        public double palka(double xxxx, double yclng, double yflor)//разновидность штриховки ли еще как то можно пользовать
        {
            // Get the current document and database
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly2 = new Polyline(1);
                acPoly2.Normal = Vector3d.ZAxis;
                acPoly2.AddVertexAt(0, new Point2d(xxxx, yclng), 0, -1, -1);
                acPoly2.AddVertexAt(1, new Point2d(xxxx, yflor), 0, -1, -1);
                acPoly2.Closed = false;
                acBlkTblRec.AppendEntity(acPoly2);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                acTrans.Commit();
            }
            return 8.88;
        }
        public double MmTextvozrast(string content, double mmx, double ceil)
        {
            // Get the current document and database
            Document acDoc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                double nexttextceileng = 0.0;
                // Create a multiline text object
                using (MText acMText = new MText())
                {
                    acMText.Location = new Point3d(mmx, ceil, 0);
                    acMText.Width = 11;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    acMText.TextHeight = 2.5;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    nexttextceileng = acMText.GeometricExtents.MinPoint.Y;
                }
                acTrans.Commit();

                return nexttextceileng;
            }
        }
        private void grafon(skwazya sk1, skwazya sk2)//отрисовка линий между одинаковыми иге между двух соседних скважин - глюк если не по полядку заполнены в файле
        {
            skwazya skh;
            skwazya skl;
            if (double.Parse(sk1.absust) > double.Parse(sk2.absust)) { skh = sk1; skl = sk2; }
            else { skh = sk2; skl = sk1; }
            //if (skh.gruntiki[0].ige != skl.gruntiki[0].ige) { }

            for (int eachgr = 0; eachgr < skh.layerscount; eachgr++)
            {
                getpair(skh.gruntiki[eachgr].ige, skh.gruntiki[eachgr].konsiceiling, skh.gruntiki[eachgr].konsifloor, skh.skwax, skl);
            }
        }
        private void getpair(string p, double top, double bot, double x, skwazya sk2)//отрисовка линий между одинаковыми иге между двух соседних скважин - глюк если не по полядку заполнены в файле
        {
            int tryfind = 0;
            string findrez = "";
            for (tryfind = 0; tryfind < sk2.layerscount; tryfind++)
            {
                if (sk2.gruntiki[tryfind].ige == p) { findrez = "ok";

                Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;



                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    
                    // Create the rectangle
                    var acPoint2 = new Point2d(x , top);
                    var acPoly2 = new Polyline(3);
                    acPoly2.Normal = Vector3d.ZAxis;
                    acPoly2.AddVertexAt(0, acPoint2, 0, -1, -1);
                    acPoly2.AddVertexAt(1, new Point2d(x , bot), 0, -1, -1);
                    acPoly2.AddVertexAt(2, new Point2d(sk2.skwax,sk2.gruntiki[tryfind].konsifloor), 0, -1, -1);
                    acPoly2.AddVertexAt(3, new Point2d(sk2.skwax, sk2.gruntiki[tryfind].konsiceiling), 0, -1, -1);
                    acPoly2.Closed = true;
                    // Add the rectangle to the block table record
                    acBlkTblRec.AppendEntity(acPoly2);
                    acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                    // Create the hatch
                    var acHatch = new Hatch(); string pattiern = sk2.gruntiki[tryfind].shtrih; //string is temp
                    if ((pattiern == "") || (pattiern == null)) { acHatch.Visible = false; pattiern = "SOLID"; }
                    acHatch.PatternScale = 10;
                    acBlkTblRec.AppendEntity(acHatch);
                    acTrans.AddNewlyCreatedDBObject(acHatch, true);
                    acHatch.SetDatabaseDefaults();
                    acHatch.SetHatchPattern(HatchPatternType.PreDefined, pattiern);
                    acHatch.Associative = true;

                    // Add the outer boundary
                    acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly2.ObjectId });

                    // Add the inner boundary

                    // Validate the hatch
                    acHatch.EvaluateHatch(true);

                    acTrans.Commit();
                }
                };
            }

        }
        private void drawscaledhole(skwazya sk, double geoscale)// считает на какой высоте центры кружков иге у каждоо грунта в скважине и пишет в скв.
        {
        #region sc calc floor
            double tempnextceil = 0;
                for (int readgr = 0; readgr < sk.layerscount; readgr++)
                {
                    sk.gruntiki[readgr].scceiling = tempnextceil;
                    sk.gruntiki[readgr].scfloor = tempnextceil - geoscale * double.Parse(sk.gruntiki[readgr].tolshmosh);
                    sk.gruntiki[readgr].sccircen = (sk.gruntiki[readgr].scceiling + sk.gruntiki[readgr].scfloor) / 2;
                    tempnextceil = sk.gruntiki[readgr].scfloor;
                }
        #endregion sc calc floor
        }
        [CommandMethod("geolog_info")]
        public void geologinfo()
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //dynamic acadObject = MgdAcApplication.AcadApplication;//zoom
            //acadObject.ZoomExtents();//zoom

            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            // Prompt for the start point
            pPtOpts.Message = "\n апрель 2016 профиль с штриховкой и новыми подвалами 8 линий в подвале";
            pPtRes = doc.Editor.GetPoint(pPtOpts);
            Point3d ptStart = pPtRes.Value;

            //http://through-the-interface.typepad.com/through_the_interface/2010/12/jigging-an-autocad-polyline-with-arc-segments-using-net.html
            //http://www.acadnetwork.com/index.php?topic=250.0
            //not very helped
            //http://stackoverflow.com/questions/19922816/how-to-cut-a-hole-in-a-hatch-autocad/19930730
            //http://through-the-interface.typepad.com/through_the_interface/2010/12/jigging-an-autocad-polyline-with-arc-segments-using-net.html
            //http://through-the-interface.typepad.com/through_the_interface/2010/06/tracing-a-boundary-defined-by-autocad-geometry-using-net.html
            //http://adndevblog.typepad.com/autocad/2013/07/create-hatch-objects-using-trace-boundaries-using-net.html
        }
        public void nigmaline(double x1ppr,double y1ppr, double x2ppr,double y2ppr)
        {
            Document doc =
  MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    var acPoly2 = new Polyline(1);
                    acPoly2.Normal = Vector3d.ZAxis;
                    acPoly2.AddVertexAt(0, new Point2d(x1ppr, y1ppr), 0, -1, -1);//ися13 у подошвы текст с подставочкой и подставочка не касается скважины а надо чтоб касалось
                    acPoly2.AddVertexAt(1, new Point2d(x2ppr, y2ppr), 0, -1, -1);
                    acPoly2.Closed = false;
                    //acPoly2.ColorIndex = 2;
                    acBlkTblRec.AppendEntity(acPoly2);
                    acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                    acTrans.Commit();
                }
  
        }
        public void nigmatext(double x, double y, string content)
        {
            if ((content != "") && (content != null))
            {
                Document doc =
                MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    /*
                    // Add the inner boundary
                    Circle acCirc = new Circle();
                    acCirc.Center = new Point3d(x + 3, (y1 + y2) / 2, 0);
                    acCirc.Radius = 2.5;
                    //acCirc.ColorIndex = 5;
                    acBlkTblRec.AppendEntity(acCirc);
                    acTrans.AddNewlyCreatedDBObject(acCirc, true);
                    */
                    // Create a multiline text object
                    MText acMText = new MText();
                    acMText.Location = new Point3d(x, y, 0);
                    acMText.Width = 35;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    acMText.TextHeight = 2;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);

                    //nexttextceileng = acMText.GeometricExtents.MinPoint.Y;

                    acTrans.Commit();

                }

            }
        }
        public void nigmatext(double x, double y, string content, string vertikal)
        {
            if ((content != "") && (content != null))
            {
                Document doc =
                MgdAcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    /*
                    // Add the inner boundary
                    Circle acCirc = new Circle();
                    acCirc.Center = new Point3d(x + 3, (y1 + y2) / 2, 0);
                    acCirc.Radius = 2.5;
                    //acCirc.ColorIndex = 5;
                    acBlkTblRec.AppendEntity(acCirc);
                    acTrans.AddNewlyCreatedDBObject(acCirc, true);
                    */
                    // Create a multiline text object
                    MText acMText = new MText();
                    acMText.Location = new Point3d(x, y, 0);
                    acMText.Width = 84;
                    acMText.Contents = content;
                    acMText.Attachment = AttachmentPoint.TopCenter;
                    acMText.TextHeight = 2;
                    acMText.Rotation = 1.57;
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);

                    //nexttextceileng = acMText.GeometricExtents.MinPoint.Y;

                    acTrans.Commit();

                }

            }
        }
        public void nigmacirc(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(acPoint.X , acPoint.Y-0.01), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(acPoint.X , acPoint.Y-10), 1.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, acPoint, 0, -1, -1);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                // Create the rectangle hole

                // Add the hole rectangle to the block table record
                //acBlkTblRec.AppendEntity(acHole);
                //acTrans.AddNewlyCreatedDBObject(acHole, true);

                // Create the hatch
                var acHatch = new Hatch();
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.PatternScale = 10;
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary
                //acHatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { acHole.ObjectId });

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(acPoint.X, acPoint.Y - 0.01), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(acPoint.X, acPoint.Y - 10), -1.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, acPoint, 0, -1, -1);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                acTrans.Commit();
            }
            /*
            //vertikal bar above km circle
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var acPoly = new Polyline(2);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, new Point2d(x, y+36), 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(x, y), -1.0, -1.0, -1.0);
                acPoly.Closed = false;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                acTrans.Commit();
            }
            */
        }
        public void nigmacirchalf(double x, double y)
        {
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(x, y);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(acPoint.X, acPoint.Y - 0.01), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(acPoint.X, acPoint.Y - 10), 1.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, acPoint, 0, -1, -1);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                // Create the rectangle hole

                // Add the hole rectangle to the block table record
                //acBlkTblRec.AppendEntity(acHole);
                //acTrans.AddNewlyCreatedDBObject(acHole, true);

                // Create the hatch
                var acHatch = new Hatch();
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.PatternScale = 10;
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary
                //acHatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { acHole.ObjectId });

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        public obrazets[] sortobriki(obrazets[] thisobriki)
        {
            thisobriki[99] = new obrazets();
            thisobriki[99].bottom = 230;
            thisobriki[99].top = 232;
            thisobriki[99].forma = "invis";
            thisobriki[99].level = 0;

            obrazets[] tempobraz;
            //var tempobraz = thisskwa.OrderBy(item => item.bottom);
            tempobraz = thisobriki.OrderBy(item => item.bottom).ToArray();
            tempobraz = thisobriki.OrderByDescending(item => item.bottom).ToArray();
            return tempobraz;
        }
        public skwazya[] sortskwa(skwazya[] inskwa)
        {
            skwazya[] skwalistsortedbyx;
            //var tempobraz = thisskwa.OrderBy(item => item.bottom);
            skwalistsortedbyx = inskwa.OrderBy(item => item.skwax).ToArray();
            //tempobraz = thisobriki.OrderByDescending( -||- )
            return skwalistsortedbyx;
        }

        [CommandMethod("TestHatchHoleself", CommandFlags.UsePickSet)]
        public void CreateHatchself()
        {
            //http://stackoverflow.com/questions/19922816/how-to-cut-a-hole-in-a-hatch-autocad/19930730
            //http://through-the-interface.typepad.com/through_the_interface/2010/12/jigging-an-autocad-polyline-with-arc-segments-using-net.html
            //http://through-the-interface.typepad.com/through_the_interface/2010/06/tracing-a-boundary-defined-by-autocad-geometry-using-net.html
            //http://adndevblog.typepad.com/autocad/2013/07/create-hatch-objects-using-trace-boundaries-using-net.html


            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = (BlockTable)acTrans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create the rectangle
                var acPoint = new Point2d(0, 0);
                var acPoly = new Polyline(4);
                acPoly.Normal = Vector3d.ZAxis;
                acPoly.AddVertexAt(0, acPoint, 0, -1, -1);
                acPoly.AddVertexAt(1, new Point2d(acPoint.X + 20, acPoint.Y), 0, -1, -1);
                acPoly.AddVertexAt(2, new Point2d(acPoint.X + 20, acPoint.Y + 20), 1.0, -1.0, -1.0);
                acPoly.AddVertexAt(3, new Point2d(acPoint.X, acPoint.Y + 20), 0.0, -1.0, -1.0);
                acPoly.Closed = true;

                // Add the rectangle to the block table record
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                // Create the rectangle hole

                // Add the hole rectangle to the block table record
                //acBlkTblRec.AppendEntity(acHole);
                //acTrans.AddNewlyCreatedDBObject(acHole, true);

                // Create the hatch
                var acHatch = new Hatch();
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                acHatch.SetDatabaseDefaults();
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatch.PatternScale = 10;
                acHatch.Associative = true;

                // Add the outer boundary
                acHatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection { acPoly.ObjectId });

                // Add the inner boundary
                //acHatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { acHole.ObjectId });

                // Validate the hatch
                acHatch.EvaluateHatch(true);

                acTrans.Commit();
            }
        }
        [CommandMethod("CreateCompositeRegions")]
        public static void CreateCompositeRegions(){
            Document doc = MgdAcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using(Transaction acTrans = db.TransactionManager.StartTransaction()){
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(db.BlockTableId,OpenMode.ForRead) as BlockTable;
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],OpenMode.ForWrite) as BlockTableRecord;
                // Create two in memory circles
                Circle acCirc1 = new Circle();acCirc1.SetDatabaseDefaults();
                acCirc1.Center = new Point3d(4, 4, 0);
                acCirc1.Radius = 4;Circle acCirc2 = new Circle();acCirc2.SetDatabaseDefaults();
                acCirc2.Center = new Point3d(4, 1, 0);acCirc2.Radius = 3;
                // Adds the circle to an object array
                DBObjectCollection acDBObjColl = new DBObjectCollection();acDBObjColl.Add(acCirc1);acDBObjColl.Add(acCirc2);
                // Calculate the regions based on each closed loop
                DBObjectCollection myRegionColl = new DBObjectCollection();myRegionColl = Region.CreateFromCurves(acDBObjColl);
                Region acRegion1 = myRegionColl[0] as Region;Region acRegion2 = myRegionColl[1] as Region;
                // Subtract region 1 from region 2
                if (acRegion1.Area > acRegion2.Area){
                    // Subtract the smaller region from the larger one
                    acRegion1.BooleanOperation(BooleanOperationType.BoolSubtract,acRegion2);acRegion2.Dispose();
                    // Add the final region to the databaseacBlkTblRec.AppendEntity(acRegion1);
                    acTrans.AddNewlyCreatedDBObject(acRegion1, true);}
                else{// Subtract the smaller region from the larger one
                    acRegion2.BooleanOperation(BooleanOperationType.BoolSubtract,acRegion1);acRegion1.Dispose();
                    // Add the final region to the database
                    acBlkTblRec.AppendEntity(acRegion2);acTrans.AddNewlyCreatedDBObject(acRegion2, true);}
                // Dispose of the in memory objects not appended to the databaseacCirc1.Dispose();acCirc2.Dispose();
 
            acTrans.Commit();}}
        public static string readconfig(string whattosearch)
        {
            string turnonbend = "unknown";

            if (File.Exists(@"C:\AtlPProf\config.txt"))
            {
                foreach (string line in File.ReadLines(@"C:\AtlPProf\config.txt"))
                {
                    if (line.Contains(whattosearch))
                    {
                        turnonbend = line;
                    }
                }
                turnonbend = turnonbend.Replace(whattosearch, "");
                return turnonbend;
            }

            else
            {
                return "unknown";
            }
        }

    }
    public class skwazya {
        public double dblabs;
        public double replace221 = 0.0;
        public double skwax = 0.0;
        public string piket = "0";
        public string piketplus = "0";
        public int probacount = 0;
        public int obrcount = 0;
        public int gruntsinskwa = 0;
        public int number;
        public double totaldeep = 0.0;
        public double totalmosh = 0.0;
        public double absdeepestpnt = 999.0;
        public double visualdeepestpnt = -999.0;
        public string absust;
        public string name;
        public string uuw;
        public string upw;
        public string start;
        public string stop;
        public string ege;
        public int layerscount=0;
        public int skwastartline = 0;//in xl for example skwa20 starts on 520 and skwa 21 on 530
        public string lastpodoshva;
        public skwazya(int Number) { number = Number;
            for (int layerscount = 0; layerscount < 40; layerscount++)
            {
                this.gruntiki[layerscount] = new gruntlayer();
                this.gruntiki[layerscount].podoshva = " ";
            }

            for (int mnogoobr = 0; mnogoobr < 99; mnogoobr++)
            {
                this.obriki[mnogoobr] = new obrazets();
                this.obriki[mnogoobr].forma = " ";
                this.obriki[mnogoobr].level = 0;
            }
        }
        //method give last layer deepth //last layer always diffetent count of layers
        public gruntlayer[] gruntiki = new gruntlayer[47];//todo configit max amount of grunts in one skwa
        public obrazets[] obriki = new obrazets[100];
        public int maxigelevels;
        public skwazya() 
        {
            for (int layerscount = 0; layerscount < 10; layerscount++)
            {
                this.gruntiki[layerscount] = new gruntlayer();
                this.gruntiki[layerscount].podoshva = " ";
            }
        }
        public void initgrunt()
        {
            for (int layerscount = 0; layerscount < 10; layerscount++)
            {
                this.gruntiki[layerscount] = new gruntlayer();
                this.gruntiki[layerscount].podoshva = " ";
            }
        }
        public string getlastpodoshva() 
        {
            return gruntiki[layerscount-1].podoshva;
        }
        public string getltpodoshvaat(int podat)
        {
            return gruntiki[podat].podoshva;
        }
        public string getlopisanieat(int podat)
        {
            return gruntiki[podat].opisanie;
        }
        public void addproba(double probay, string type)
        {
            this.probacount++; obriki[probacount] = new obrazets();
            obriki[probacount].bottom = probay;
            //obriki[probacount].forma = "any";
            obriki[probacount].forma = type;
        }
        public void addproba(string probay, string type)
        {
            //if rusconfig coma dot delimiter                string replacement = Regex.Replace(s, @"\t|\n|\r", "");
            if ((probay != null) && (probay != ""))
            {
                probay = probay.Replace(",", ".");
                this.probacount++; obriki[probacount] = new obrazets();
                obriki[probacount].bottom = double.Parse(probay);
                obriki[probacount].forma = type;
            }
        }

        //public void setopisfloor(double incoming){ this.opisfloor = }
    }
    public class gruntlayer
    {
        public double scceiling;
        public double scfloor;
        public double sccircen;
        public double monoral;
        public double naru;
        public string podoshva;
        public double podoshvadigit;
        public string moshnost;
        public double moshnostdigit = 0.0;
        public double opisfloor;
        public double opisceiling;
        public double leftfloor;
        public double leftceiling;
        public double konsiceiling;
        public double circen;
        public double konsifloor;
        public string opisanie;
        public string narush;
        public string hassnow;
        public string haspalka;
        //public string monoral;
        public string tolshmosh;
        public string ige;
        public string upv;
        public string uuv;
        public string obrvod;
        public string vozrast;
        public string shtrih;
        public string konsist;
        public string obraznar;
        public string obrazmon;
        public string uuw;
        public string upw;
        public int thisigecounter;
        public gruntlayer() { }
        public void setopisfloor(double incoming) { opisfloor = incoming; }
    }
    public class obrazets
    {
        public string forma;
        public double bottom;
        public double top;
        public int level;
    }
    public class povorot
    {
        public int cellseredini =0;
        public int cellbegin =0;
        public double beginx = 0;
        public double beginppr = 0;
        public int cellend = 0;
        public double endx = 0;
        public double endppr = 0;
        public string ugol = "empty";
        public string direction = "left";
        public povorot() { }
        public string textformod = "empty";
    }
}
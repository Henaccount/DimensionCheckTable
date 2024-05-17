using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

public class DimensionTool
{
    [CommandMethod("DimensionCheckTable")]
    public void AddDimensionsWithBalloons()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        // Prompt user to select dimensions
        PromptSelectionOptions options = new PromptSelectionOptions();
        options.MessageForAdding = "Select dimensions:";
        PromptSelectionResult result = ed.GetSelection(options);

        if (result.Status != PromptStatus.OK)
        {
            ed.WriteMessage("\nNo dimensions selected.");
            return;
        }

        SelectionSet selectionSet = result.Value;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            // Create a new layer for the dimensions
            LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            string newLayerName = "DimensionsWithBalloons";
            ObjectId layerId;

            if (!layerTable.Has(newLayerName))
            {
                LayerTableRecord newLayer = new LayerTableRecord();
                newLayer.Name = newLayerName;
                layerTable.UpgradeOpen();
                layerId = layerTable.Add(newLayer);
                tr.AddNewlyCreatedDBObject(newLayer, true);
            }
            else
            {
                layerId = layerTable[newLayerName];
            }

            int balloonNumber = 1;

            // Process each selected dimension
            foreach (SelectedObject selObj in selectionSet)
            {
                if (selObj != null)
                {
                    Entity ent = (Entity)tr.GetObject(selObj.ObjectId, OpenMode.ForWrite);

                    if (ent is Dimension dimension)
                    {
                        // Copy dimension to the new layer
                        Dimension copiedDimension = (Dimension)dimension.Clone();
                        copiedDimension.Layer = newLayerName;
                        BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        modelSpace.AppendEntity(copiedDimension);
                        tr.AddNewlyCreatedDBObject(copiedDimension, true);

                        // Add balloon
                        MText balloon = new MText();
                        balloon.Contents = balloonNumber.ToString();
                        balloon.TextHeight = 2.5;
                        balloon.Location = copiedDimension.TextPosition + new Vector3d(5, 5, 0); // Adjust position as needed
                        balloon.Layer = newLayerName;
                        modelSpace.AppendEntity(balloon);
                        tr.AddNewlyCreatedDBObject(balloon, true);

                        // Add row to the table
                        ObjectId tableId = GetTableId(tr, db);
                        if (!tableId.IsNull)
                        {
                            Table table = (Table)tr.GetObject(tableId, OpenMode.ForWrite);
                            table.InsertRows(table.Rows.Count, table.Rows[0].Height, 1);

                            int rowIndex = table.Rows.Count - 1;
                            table.Cells[rowIndex, 0].TextString = balloonNumber.ToString();
                            table.Cells[rowIndex, 1].TextString = Math.Round(dimension.Measurement, 1).ToString();
                            table.Cells[rowIndex, 2].TextString = Math.Round(dimension.Measurement + GetTolerance(dimension.Measurement, 'm'), 1).ToString();
                            table.Cells[rowIndex, 3].TextString = Math.Round(dimension.Measurement - GetTolerance(dimension.Measurement, 'm'), 1).ToString();

                        }

                        balloonNumber++;
                    }
                }
            }

            tr.Commit();
        }
    }

    private ObjectId GetTableId(Transaction tr, Database db)
    {
        // This method assumes that there is a table already present in the drawing
        // Modify as necessary to correctly locate and return the table ObjectId
        BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        foreach (ObjectId objId in btr)
        {
            Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
            if (ent is Table table && ent.Layer.Equals("DimensionsWithBalloons"))
            {
                return table.ObjectId;
            }
        }

        return ObjectId.Null;
    }

    public static double GetTolerance(double nominalLength, char toleranceClass)
    {
        if (toleranceClass == 'f')
        {
            if (nominalLength >= 0.5 && nominalLength <= 3) return 0.05;
            if (nominalLength > 3 && nominalLength <= 6) return 0.05;
            if (nominalLength > 6 && nominalLength <= 30) return 0.1;
            if (nominalLength > 30 && nominalLength <= 120) return 0.15;
            if (nominalLength > 120 && nominalLength <= 400) return 0.2;
            if (nominalLength > 400 && nominalLength <= 1000) return 0.3;
            if (nominalLength > 1000 && nominalLength <= 2000) return 0.5;
        }
        else if (toleranceClass == 'm')
        {
            if (nominalLength >= 0.5 && nominalLength <= 3) return 0.1;
            if (nominalLength > 3 && nominalLength <= 6) return 0.1;
            if (nominalLength > 6 && nominalLength <= 30) return 0.2;
            if (nominalLength > 30 && nominalLength <= 120) return 0.3;
            if (nominalLength > 120 && nominalLength <= 400) return 0.5;
            if (nominalLength > 400 && nominalLength <= 1000) return 0.8;
            if (nominalLength > 1000 && nominalLength <= 2000) return 1.2;
            if (nominalLength > 2000 && nominalLength <= 4000) return 2.0;
        }
        else if (toleranceClass == 'c')
        {
            if (nominalLength >= 0.5 && nominalLength <= 3) return 0.2;
            if (nominalLength > 3 && nominalLength <= 6) return 0.3;
            if (nominalLength > 6 && nominalLength <= 30) return 0.5;
            if (nominalLength > 30 && nominalLength <= 120) return 0.8;
            if (nominalLength > 120 && nominalLength <= 400) return 1.2;
            if (nominalLength > 400 && nominalLength <= 1000) return 2.0;
            if (nominalLength > 1000 && nominalLength <= 2000) return 3.0;
            if (nominalLength > 2000 && nominalLength <= 4000) return 4.0;
        }
        else if (toleranceClass == 'v')
        {
            if (nominalLength > 3 && nominalLength <= 6) return 0.5;
            if (nominalLength > 6 && nominalLength <= 30) return 1.0;
            if (nominalLength > 30 && nominalLength <= 120) return 1.5;
            if (nominalLength > 120 && nominalLength <= 400) return 2.5;
            if (nominalLength > 400 && nominalLength <= 1000) return 4.0;
            if (nominalLength > 1000 && nominalLength <= 2000) return 5.0;
            if (nominalLength > 2000 && nominalLength <= 4000) return 8.0;
        }
        return double.NaN; // Return NaN if the tolerance class or nominal length is invalid
    }
}

/****************************** Module Header ******************************\
Module Name:  MyMLeaderJig.cs
Project:      ConduitQtyTable
Copyright (c) 2012 Kyle T. Morin.

<Description of the file>

All rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Aec.Building.ApplicationServices;
using Autodesk.Aec.Building.DatabaseServices;
using Autodesk.Aec.Building.Elec.DatabaseServices;
using Autodesk.AutoCAD.Colors;
using System.Runtime.InteropServices;

namespace ConduitQtyTable
{
    public class MyCmds
    {
        public static string _mlLayer = "DE-ANNO-TEXT";    //Set default layer to place final MLeader object on.

        class MLeaderJig : EntityJig
        {

            Point3dCollection m_pts;
            Point3d m_tempPoint;
            string m_contents;
            int m_leaderIndex;
            int m_leaderLineIndex;

            public MLeaderJig(string contents)
                : base(new MLeader())
            {
                // Store the string passed in

                m_contents = contents;

                // Create a point collection to store our vertices

                m_pts = new Point3dCollection();

                // Create mleader and set defaults

                MLeader ml = Entity as MLeader;
                ml.SetDatabaseDefaults();
                ml.EnableLanding = true;
                ml.Annotative = AnnotativeStates.True;
                ml.ExtendLeaderToText = true;
                

                // Set up the MText contents

                ml.ContentType = Autodesk.AutoCAD.DatabaseServices.ContentType.MTextContent;
                MText mt = new MText();
                mt.SetDatabaseDefaults();
                mt.Contents = m_contents;
                mt.Annotative = AnnotativeStates.True;
                mt.TextHeight = 9.0;
                ml.MText = mt;
                ml.TextAlignmentType = TextAlignmentType.LeftAlignment;
                ml.TextAttachmentType = TextAttachmentType.AttachmentMiddle;

                //Set the layer
                ml.Layer = _mlLayer;

                // Set the frame and landing properties
                /*
                ml.EnableDogleg = true;
                ml.EnableFrameText = true;
                ml.EnableLanding = true;
                */
                
                // TODO: Test, Reduce the standard landing gap
                //ml.LandingGap = 0.05;

                // Add a leader, but not a leader line (for now)

                m_leaderIndex = ml.AddLeader();
                m_leaderLineIndex = -1;
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                JigPromptPointOptions opts = new JigPromptPointOptions();

                // Not all options accept null response
                opts.UserInputControls =
                  (UserInputControls.Accept3dCoordinates |
                  UserInputControls.NoNegativeResponseAccepted
                  );

                // Get the first point
                if (m_pts.Count == 0)
                {
                    opts.UserInputControls |= UserInputControls.NullResponseAccepted;
                    opts.Message = "\nStart point of multileader: ";
                    opts.UseBasePoint = false;
                }
                // And the second
                else if (m_pts.Count == 1)
                {
                    opts.BasePoint = m_pts[m_pts.Count - 1];
                    opts.UseBasePoint = true;
                    opts.Message = "\nSpecify multileader vertex: ";
                }
                // And subsequent points
                else if (m_pts.Count > 1)
                {
                    opts.UserInputControls |= UserInputControls.NullResponseAccepted;
                    opts.BasePoint = m_pts[m_pts.Count - 1];
                    opts.UseBasePoint = true;
                    opts.SetMessageAndKeywords(
                      "\nSpecify multileader vertex or [End]: ",
                      "End"
                    );
                }
                else // Should never happen
                    return SamplerStatus.Cancel;

                PromptPointResult res = prompts.AcquirePoint(opts);

                if (res.Status == PromptStatus.Keyword)
                {
                    if (res.StringResult == "End")
                    {
                        return SamplerStatus.Cancel;
                    }
                }

                if (m_tempPoint == res.Value)
                {
                    return SamplerStatus.NoChange;
                }
                else if (res.Status == PromptStatus.OK)
                {
                    m_tempPoint = res.Value;
                    return SamplerStatus.OK;
                }
                return SamplerStatus.Cancel;
            }

            protected override bool Update()
            {
                try
                {
                    if (m_pts.Count > 0)
                    {
                        // Set the last vertex to the new value

                        MLeader ml = Entity as MLeader;
                        ml.SetLastVertex(m_leaderLineIndex, m_tempPoint);

                        // Adjust the text location

                        Vector3d dogvec = ml.GetDogleg(m_leaderIndex);
                        double doglen = ml.DoglegLength;
                        double landgap = ml.LandingGap;
                        ml.TextLocation = m_tempPoint + ((doglen + landgap) * dogvec);
                    }
                }
                catch (System.Exception ex)
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    doc.Editor.WriteMessage("\nException: " + ex.Message);
                    return false;
                }
                return true;
            }

            public void AddVertex()
            {
                MLeader ml = Entity as MLeader;

                // For the first point...

                if (m_pts.Count == 0)
                {
                    // Add a leader line

                    m_leaderLineIndex = ml.AddLeaderLine(m_leaderIndex);

                    // And a start vertex

                    ml.AddFirstVertex(m_leaderLineIndex, m_tempPoint);

                    // Add a second vertex that will be set
                    // within the jig

                    ml.AddLastVertex(m_leaderLineIndex, new Point3d(0, 0, 0));
                }
                else
                {
                    // For subsequent points,
                    // just add a vertex

                    ml.AddLastVertex(m_leaderLineIndex, m_tempPoint);
                }

                // Reset the attachment point, otherwise
                // it seems to get forgotten

                ml.TextAttachmentType = TextAttachmentType.AttachmentMiddle;

                // Add the latest point to our history

                m_pts.Add(m_tempPoint);
            }

            public void RemoveLastVertex()
            {
                // We don't need to actually remove
                // the vertex, just reset it

                MLeader ml = Entity as MLeader;
                if (m_pts.Count >= 1)
                {
                    Vector3d dogvec = ml.GetDogleg(m_leaderIndex);
                    double doglen = ml.DoglegLength;
                    double landgap = ml.LandingGap;
                    ml.TextLocation = m_pts[m_pts.Count - 1] + ((doglen + landgap) * dogvec);
                }
            }

            public Entity GetEntity()
            {
                return Entity;
            }
        }

        [CommandMethod("SETMLEADERLAYER")]
        public void SetMLeaderLayer()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptStringOptions pso = new PromptStringOptions("\nEnter new default layer for ConduitQuantityTable Mleader: ");

            PromptResult pr = ed.GetString(pso);

            if (pr.Status == PromptStatus.OK)
            {
                string pResult = pr.StringResult;

                if (pResult == "" || pResult == null)
                {
                    ed.WriteMessage("\nMust not be blank");
                }
                else
                {
                    if (LayerExists(pResult) == false)
                    {
                        pResult = pResult.ToUpper();
                        CreateLayer(pResult);   //create New layer
                        _mlLayer = pResult;     //set the layer to default.
                        ed.WriteMessage("\nNew default layer set to: {0}", pResult);
                    }
                    else if (LayerExists(pResult) == true)
                    {
                        _mlLayer = pResult;
                        ed.WriteMessage("\nNew default layer set to: {0}", pResult);
                    }
                    else
                    {
                        ed.WriteMessage("\nSomething went terribly wrong.");
                    }
                }

            }
        }

        private bool LayerExists(string pResult)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                //Open layertable for read
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                //Check layer table against input to see if match
                try
                {
                    if (lt.Has(pResult))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                    tr.Commit();
                }
                catch (System.Exception)
                {
                    tr.Abort();
                    ed.WriteMessage("ERROR:!!!");
                    return false;
                }
            }
        }

        private void CreateLayer(string pResult)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            //Open layertable for write
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);

                try
                {
                    LayerTableRecord ltr = new LayerTableRecord();
                    //set props
                    ltr.Name = pResult;
                    ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 0);
                    //add to table
                    ObjectId ltId = lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);

                    tr.Commit();
                }
                catch (System.Exception)
                {
                    ed.WriteMessage("ERROR Creating layer...");
                    tr.Abort();
                }
            }
        }

        [CommandMethod("CONDUITQTYTABLE")]
        public void MyConduitTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            //Get the text oustide of the jig
            string conduits = SelectConduits();

            //Set layer for ellipse to equal _mlLayer
            Transaction trLt = db.TransactionManager.StartTransaction();
            using (trLt)
            {
                try
                {
                    LayerTable lt = (LayerTable)trLt.GetObject(db.LayerTableId, OpenMode.ForRead);

                    if (lt.Has(_mlLayer) == true)
                    {
                        db.Clayer = lt[_mlLayer];
                        trLt.Commit();
                    }
                    else
                    {
                        try 
	                    {	        
                            CreateLayer(_mlLayer);
                            db.Clayer = lt[_mlLayer];
                            trLt.Commit();
	                    }
	                    catch (System.Exception ex)
	                    {
                            ed.WriteMessage("Error: Can't create layer");
                            trLt.Abort();
	                    }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("error: Can't make layer current.");
                    trLt.Abort();
                }
            }   //End transaction to set layer and/or create it if doesn't exist to prevent crash.

            if (conduits == "" || conduits == null || conduits == "FAIL")
            {
                ed.WriteMessage("\nMust select at least [1] Conduit");
            }
            else
            {
                //Call ellipse command first
                bool ec = EllipseCall.EllipseJigWait();

                if (ec == true)
                {
                    //Create MLeaderJig
                    MLeaderJig jig = new MLeaderJig(conduits);

                    //Loop to set vertices
                    bool bSuccess = true, bComplete = false;
                    while (bSuccess && !bComplete)
                    {
                        PromptResult dragres = ed.Drag(jig);
                        bSuccess = (dragres.Status == PromptStatus.OK);

                        if (bSuccess)
                            jig.AddVertex();

                        bComplete = (dragres.Status == PromptStatus.None);
                        if (bComplete)
                            jig.RemoveLastVertex();
                    }

                    if (bComplete)
                    {
                        //Append entity
                        Transaction tr = db.TransactionManager.StartTransaction();
                        using (tr)
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);
                            btr.AppendEntity(jig.GetEntity());
                            tr.AddNewlyCreatedDBObject(jig.GetEntity(), true);
                            tr.Commit();
                        }
                    }
                }
            }
        }

        public string SelectConduits()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //TODO: add selection filter for typeof(Conduit)

            try
            {
                PromptSelectionOptions selOpts = new PromptSelectionOptions();

                selOpts.MessageForAdding = "\nSelect Conduits: ";

                PromptSelectionResult psr;
                psr = ed.GetSelection(selOpts);

                // Grab object Ids and go through one by one
                Transaction tr = doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    try
                    {
                        List<string> partSizes = new List<string>();
                        int total = 0;
                        ObjectId[] objIds = psr.Value.GetObjectIds();

                        //Loop through the selected objectIds
                        foreach (ObjectId objId in objIds)
                        {

                            Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead); //use entity as object type to find
                            Member mem = tr.GetObject(objId, OpenMode.ForRead) as Member;   //set as member (needed to process AEC)

                            // Checks to make sure that ObjectId can be pulled as a Member, if not, then discards object from selection,
                            // Then checks to make sure Member is ofType 'Conduit' or else discards.
                            if (mem != null && mem.ToString() == "Autodesk.Aec.Building.Elec.DatabaseServices.Conduit")
                            {
                                partSizes.Add(GetPartSize(mem));
                                total++;
                            }
                            else
                            {
                                continue;
                            }
                        }

                        //Count each occurance of each distinct conduit and return in nice formatted text.

                        /* Have no clue how this works, but it works! */
                        var query = from x in partSizes
                                    group x by x into g
                                    let count = g.Count()
                                    orderby count descending
                                    select new { Name = g.Key, Count = count };

                        string res = "";

                        foreach (var result in query)
                        {
                            res += "\n(" + result.Count + ")  " + result.Name;
                        }
                        /**/

                        //Ugly way to remove starting new-line, but hey, it works.
                        res = res.TrimStart('\n');

                        //Return result as nice string block.
                        tr.Commit();
                        return res;
                    }
                    catch (NullReferenceException)
                    {
                        tr.Abort();
                        return "FAIL";
                    }
                    catch (SystemException ex)
                    {

                        //Sends system exception, but only for Debugging
                        //ed.WriteMessage("secondary catch:: Error: {0}", ex.ToString());
                        tr.Abort();
                        return "FAIL";
                    }
                }
            }
            catch (SystemException ex)
            {

                //ed.WriteMessage("Error: {0}", ex.ToString());
                //ed.WriteMessage("main catch:: Error: {0}", ex.ToString());
                return "FAIL";
            }
        }

        private string GetPartSize(Member mem)
        {
            DataRecord dr = PartManager.GetPartData(mem); //Pulls the part from the object as a datarecord
            DataField df = dr.FindByContextAndIndex(Context.CatalogPartSizeName, 0);    //Searches through MEP catalog for part name.
            string partSizeName = df.ValueString;

            //Remove trailing end leaving only number.
            //partSizeName = partSizeName.Remove(7);

            return partSizeName;
        }
    }
}
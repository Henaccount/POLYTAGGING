//
// (C) Copyright 2009 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Collections.Specialized;
using System.IO;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;



namespace pssCommands
{


    public class PropCopy
    {


        static List<KeyValuePair<string, string>> mb_listProp;



        public static void StartPexTagUpdate() // This method can have any name
        {
            // Set the environment to the current active database and editor when the command execute
            Helper.Initialize();

            Helper.oEditor.WriteMessage("\nUpdate tag values");

            //string dwgtype = PnPProjectUtils.GetActiveDocumentType();
            string dwgtype = "PnId";


            if (dwgtype != "PnId")
            {  //"PnId", "Piping", "Ortho", "Iso"

                Helper.oEditor.WriteMessage("\nDrawing must be a PnId drawing"

                              + " in the current project.\n");

                return;

            }

            PromptSelectionOptions pso = new PromptSelectionOptions();

            pso.MessageForAdding = "\nSelect " + dwgtype

                                 + " objects to update <All>: ";

            PromptSelectionResult selResult = Helper.oEditor.GetSelection(pso);


            if (selResult.Status == PromptStatus.Cancel)

                return;

            if (selResult.Status != PromptStatus.OK)

                selResult = Helper.oEditor.SelectAll();

            SelectionSet ss = selResult.Value;

            ObjectId[] objIds = ss.GetObjectIds();

            PexTagUpdate(objIds);

            Helper.Terminate();
        }





        public static void PexTagUpdate(ObjectId[] objIds) // This method can have any name
        {
            try
            {

                int nTags = 0, nTagsModified = 0;



                foreach (ObjectId objId in objIds)
                {

                    using (Transaction t = Helper.oDatabase.TransactionManager.StartTransaction())
                    {

                        StringCollection strNames = new StringCollection();

                        StringCollection strNewValues = new StringCollection();

                        strNames.Add("Tag");


                        strNewValues.Add("");  // A blank value causes a recalc.

                        // A tag that is currently blank will remain blank.



                        DBObject obj = t.GetObject(objId, OpenMode.ForRead);

                        if (!Helper.ActiveDataLinksManager.HasLinks(objId))
                        {
                            //oEditor.WriteMessage(" Not a P&ID object");
                            continue; // 
                        }

                        nTags++;

                        StringCollection strOldValues =

                            Helper.ActiveDataLinksManager.GetProperties(objId, strNames, true);



                        //----- HERE IT IS. UPDATE OF THE TAG ------------------------

                        Helper.ActiveDataLinksManager.SetProperties(objId, strNames, strNewValues);

                        //------------------------------------------------------------



                        t.Commit();



                        // Get new values

                        StringCollection strRecalcValues =

                            Helper.ActiveDataLinksManager.GetProperties(objId, strNames, true);

                        if (strRecalcValues[0] == "")

                            continue;

                        nTagsModified++;

                        // Report changes

                        Helper.oEditor.WriteMessage("\n   tag=" + strRecalcValues[0]);

                        if (strOldValues[0] != strRecalcValues[0])

                            Helper.oEditor.WriteMessage("  oldtag=" + strOldValues[0]);




                    }



                }

            }
            catch (System.Exception e)
            {

                /*if (t != null)
                {

                    t.Abort();

                    t.Dispose();

                }*/

                //Helper.oEditor.WriteMessage("\nCould not update tag: "+ e.Message.ToString());

            }

            //Helper.oEditor.WriteMessage("\nDone (" + nTagsModified + " of " + nTags + " modified).");

        }



        public static void mb_PropCopy()
        {


            try
            {
                Helper.Initialize();

                TypedValue[] filterlist = new TypedValue[4];

                filterlist[0] = new TypedValue(-4, "<AND");
                filterlist[1] = new TypedValue(0, "*POLYLINE");
                filterlist[2] = new TypedValue(48, "2.11");
                filterlist[3] = new TypedValue(-4, "AND>");

                SelectionFilter filter = new SelectionFilter(filterlist);

                PromptSelectionResult polysel = Helper.oEditor.SelectAll(filter);
                SelectionSet polyselSet = polysel.Value;

                if (polyselSet != null)
                {

                    



                    ObjectId[] polyIds = polyselSet.GetObjectIds();

                    //for all polylines
                    for (int j = 0; j < polyIds.Length; j++)
                    {
                        Point3dCollection polypointscoll = new Point3dCollection();
                        String aktValue = "";
                        Transaction oTransaction = null;
                        using (oTransaction = Helper.oDatabase.TransactionManager.StartTransaction())
                        {
                            Entity ent = oTransaction.GetObject(polyIds[j], OpenMode.ForRead) as Entity;
                            HyperLinkCollection linkCollection = ent.Hyperlinks;
                            aktValue = linkCollection[0].DisplayString;
                            Polyline thepoly = (Polyline)ent;
                            for (int z = 0; z < thepoly.NumberOfVertices; z++)
                            {
                                polypointscoll.Add(thepoly.GetPoint3dAt(z));
                            }
                        }
                        //oTransaction = null;
                        //if hyperlink display string not empty
                        //get polyline Vertices, select window
                        if (!aktValue.Equals(""))
                        {
                            TypedValue[] acpfilterlist = new TypedValue[1];

                            acpfilterlist[0] = new TypedValue(0, "ACPP*");

                            SelectionFilter acpfilter = new SelectionFilter(acpfilterlist);
                            PromptSelectionResult selResult = Helper.oEditor.SelectWindowPolygon(polypointscoll, acpfilter);
                            SelectionSet selSet = selResult.Value;

                            if (selSet != null)
                            {
                                
                                ObjectId[] objIds = selSet.GetObjectIds();
                                if (objIds != null)
                                {


                                    for (int i = 0; i < objIds.Length; i++)
                                    {
                                        ObjectId mb_objId = objIds[i];
                                        brush_configured(aktValue, mb_objId);
                                    }

                                }
                            }
                        }
                    }
                    //end if
                    //end for all polylines

                }
            }
            catch (System.Exception e)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                Helper.oEditor.WriteMessage(trace.ToString());
                Helper.oEditor.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                Helper.oEditor.WriteMessage("message: " + e.Message);
            }
            finally
            {
                Helper.Terminate();
            }



        }


        public static void brush_configured(String aktValue, ObjectId mb_objId)
        {


            try
            {
                Transaction oTransaction = null;
                using (oTransaction = Helper.oDatabase.TransactionManager.StartTransaction())
                {

                    try
                    {
                        String objtype = "Polyline";

                        if (Helper.ActiveDataLinksManager.HasLinks(mb_objId))
                        {
                            mb_listProp = Helper.ActiveDataLinksManager.GetAllProperties(mb_objId, true);



                            //config einlesen



                            string therows = String.Format("{0}_COPYTO_ALL", objtype);
                            string currentDir = Environment.CurrentDirectory;

                            string fullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                            string theDirectory = Path.GetDirectoryName(fullPath);

                            XDocument xdoc = XDocument.Load(theDirectory + "\\config.xml");

                            IEnumerable<XElement> theele = xdoc.Root.Elements("Rule").Where(e => e.Attribute("for").Value.Equals(therows));

                            List<string> xlist = theele.Select(element => element.Value).ToList();

                            bool changed = false;

                            if (xlist.Count != 0)
                            {

                                for (int i = 0; i < mb_listProp.Count; i++)
                                {
                                    String aktKey = "Hyperlink";


                                    if (!aktKey.StartsWith("PnP"))
                                    {
                                        try
                                        {
                                            for (int j = 0; j < mb_listProp.Count; j++)
                                            {
                                                String mb_aktKey = mb_listProp[j].Key;
                                                String mb_aktValue = mb_listProp[j].Value;

                                                if (!mb_aktKey.Equals("Tag") && xlist.Contains(aktKey + "_COPYTO_" + mb_aktKey))
                                                {
                                                    KeyValuePair<string, string> oldVal = new KeyValuePair<string, string>(mb_aktKey, mb_aktValue);
                                                    KeyValuePair<string, string> newVal = new KeyValuePair<string, string>(mb_aktKey, aktValue);

                                                    bool remAtt = mb_listProp.Remove(oldVal);
                                                    mb_listProp.Add(newVal);
                                                    changed = true;
                                                    break;
                                                }
                                            }

                                        }
                                        catch (System.Exception e)
                                        {
                                            Helper.oEditor.WriteMessage(e.Message);
                                        }

                                    }

                                }

                            }
                            else
                            {
                                //Helper.oEditor.WriteMessage("Kombination der Elemente nicht konfiguriert, Element wird übersprungen.");

                                return;
                            }

                            if (changed)
                            {
                                System.Collections.Specialized.StringCollection strNames = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection strValues = new System.Collections.Specialized.StringCollection();
                                for (int g = 0; g < mb_listProp.Count; g++)
                                {
                                    String name = mb_listProp[g].Key;
                                    String value = mb_listProp[g].Value;
                                    strNames.Add(name);
                                    strValues.Add(value);
                                }

                                Helper.ActiveDataLinksManager.SetProperties(mb_objId, strNames, strValues);
                                oTransaction.Commit();

                                ObjectId[] theset = new ObjectId[1];
                                theset[0] = mb_objId;
                                PexTagUpdate(theset);

                            }





                            //Helper.oEditor.WriteMessage("Alle konfigurierten Attribute, nicht aber PNP-Attr., Tag wurden kopiert.");
                        }
                    }
                    catch (System.Exception e)
                    {
                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                        Helper.oEditor.WriteMessage(trace.ToString());
                        Helper.oEditor.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                        Helper.oEditor.WriteMessage("message: " + e.Message);
                        /*if (e.GetType().ToString() == "Autodesk.ProcessPower.DataLinks.DLException")
                        {
                            if (oTransaction != null)
                            {
                                oTransaction.Abort();
                                oTransaction.Dispose();
                            }
                        }*/
                    }




                }

            }
            catch (System.Exception e)
            {
                /*if (oTransaction != null)
                {
                    oTransaction.Abort();
                    oTransaction.Dispose();
                }*/
                Helper.oEditor.WriteMessage(e.Message.ToString());
            }

        }




    }

}




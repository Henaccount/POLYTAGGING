//
// (C) Copyright 2013 by Autodesk, Inc.
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

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;


using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

/*
 * 0.2 put all in one file
 * 0.3 info now available through "pssCommands" command, because extensionapplication did prevent dll from autoloading in acad2016doc.lsp
 */

namespace pssCommands
{
    

    public class Helper
    {
        public static PnIdProject PnIdProject { get; set; }
        public static Document ActiveDocument { get; set; }
        public static DataLinksManager ActiveDataLinksManager { get; set; }
        public static Database oDatabase { get; set; }
        public static Editor oEditor { get; set; }


        public static void pssCommands()
        {
            Helper.Initialize();
            Helper.oEditor.WriteMessage("\npssbayv0.3-commands:\n");
            Helper.oEditor.WriteMessage("\ncommand RotateOnInsert \"Endkappe\"\n");
            Helper.oEditor.WriteMessage("\nqy_QuickRotate\n");
            Helper.oEditor.WriteMessage("\nqq_PexTagUpdate\n");
            Helper.oEditor.WriteMessage("\nqy_PropCopy\n");
            Helper.oEditor.WriteMessage("\nqy_MultiFlowReverse\n");
            Helper.Terminate();
        }

        public static ObjectId ConvertRowIdToObjectId(int iRowId)
        {
            try
            {
                // Fetch ObjectId of part.
                //
                var res = Helper.ActiveDataLinksManager.FindAcPpObjectIds(iRowId);
                var pnidId = res.First.Value;
                ObjectId pnidObjectId = Helper.ActiveDataLinksManager.MakeAcDbObjectId(pnidId);
                return pnidObjectId;
            }
            catch { return new ObjectId(); }
        }

        public static bool Initialize()
        {
            try
            {
                if (PlantApplication.CurrentProject == null)
                    return false;

                Helper.PnIdProject = (PnIdProject)PlantApplication.CurrentProject.ProjectParts["PnId"];
                Helper.ActiveDataLinksManager = Helper.PnIdProject.DataLinksManager;
                Helper.ActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Helper.oDatabase = Helper.ActiveDocument.Database;
                Helper.oEditor = Helper.ActiveDocument.Editor;
                return true;
            }
            catch { return false; }
        }

        public static void Terminate()
        {
            try
            {
            Helper.PnIdProject = null;
            Helper.ActiveDataLinksManager = null;
            Helper.ActiveDocument = null;
            Helper.oDatabase = null;
            Helper.oEditor = null;
            }
            catch { }
        }



    }
}

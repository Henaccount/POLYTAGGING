
using Autodesk.AutoCAD.Runtime;





[assembly: CommandClass(typeof(pssCommands.Commands))]



namespace pssCommands
{
    

    //in toolpalette command eintragen: command RotateOnInsert "Endkappe"

        public class Commands
        {
            #region Commands





            [CommandMethod(/*NOXLATE*/"qq_PexTagUpdate", CommandFlags.Redraw | CommandFlags.UsePickSet)]
            public static void StartPexTagUpdate()
            {
                PropCopy.StartPexTagUpdate();
            }

            [CommandMethod(/*NOXLATE*/"Polytagging")]
            public static void mb_PropCopy()
            {
                PropCopy.mb_PropCopy();
            }



            #endregion

           

        }
    }


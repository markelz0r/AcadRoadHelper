
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;

namespace GroundToBigTable
{
   public class CommonMethods
   {
      public static void FillPiket(Autodesk.AutoCAD.EditorInput.Editor editor, Table table, Transaction tr, int row)
      {
         var piketObj = editor.GetEntity("Выделите Пикет\n");

         if (piketObj.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
            throw new Exception();
            
         var piketStr = tr.GetObject(piketObj.ObjectId, OpenMode.ForRead) as DBText;

         if (table != null)

         {
            table.UpgradeOpen();
            table.Cells[row, 0].TextString = piketStr.TextString;
         }
      }
   }
}
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

[assembly: CommandClass(typeof(GroundToBigTable.GroundTablePlanning))]

namespace GroundToBigTable
{
   [SuppressMessage("ReSharper", "StringLiteralTypo")]
   public class GroundTablePlanning
   {
      [CommandMethod("Ground_Table_Planning")]
      public void Ground_Table_Planning()
      {
         var d = Application.DocumentManager.MdiActiveDocument;
         var editor = d.Editor;

         var result = editor.GetEntity("Выделите общую таблицу\n");

         if (result.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
         {
            editor.WriteMessage("Ошибка!, завершение программы. Дай Бог вам здоровья\n");
            return;
         }

         var tr = d.Database.TransactionManager.StartTransaction();

         using (tr)
         {
            var table = tr.GetObject(result.ObjectId, OpenMode.ForWrite, false) as Table;

               for (var row = 1; row < int.MaxValue; row++)
               {
                  try
                  {
                     CommonMethods.FillPiket(editor, table, tr, row);
                     ProcessDimensions(d, result, editor, tr, row);
                  }
                  catch (Exception)
                  {
                     tr.Commit();
                     editor.WriteMessage("Ошибка!, завершение программы. Дай Бог вам здоровья\n");
                     return;
                  }
               }
               tr.Commit();
         }
      }

      private static void ProcessDimensions(Document d, PromptEntityResult result, Editor editor,
         Transaction tr, int row)
      {
         var numberOfColumns = GetNumberOfColumns(d, result);
         for (var column = 1; column < numberOfColumns; column++)
         {
            var table = tr.GetObject(result.ObjectId, OpenMode.ForWrite, false) as Table;
            var columnName = table.Cells[0, column].TextString;
            if (columnName != string.Empty)
            {
               var dimensionResult = editor.GetEntity($"Выделите размер {columnName}:\n");
               if (result.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
               {
                  editor.WriteMessage("Ошибка!, завершение программы. Дай Бог вам здоровья\n");
                  break;
               }
               
               var mods = Control.ModifierKeys;
               var shift = (mods & Keys.Shift) > 0;
               var control = (mods & Keys.Control) > 0;
               
               if (shift || control)
                  continue;

               var dimension = tr.GetObject(dimensionResult.ObjectId, OpenMode.ForRead, false) as Dimension;
               if (dimension != null)
               {
                  var measurement = Math.Round(dimension.Measurement, 2);
                  table.Cells[row, column].TextString = measurement.ToString(CultureInfo.InvariantCulture);
               }
               else
               {
                  editor.WriteMessage("Неверный выбор, попробуйте ещё раз...\n");
                  column -= 1;
               }
            }
         }
      }

      private static int GetNumberOfColumns(Document d, PromptEntityResult result)
      {
         int numberOfRows;
         var tr = d.Database.TransactionManager.StartTransaction();

         using (tr)
         {
            var table = tr.GetObject(result.ObjectId, OpenMode.ForWrite, false) as Table;
            numberOfRows = table.Columns.Count;
         }

         return numberOfRows;
      }
   }
}
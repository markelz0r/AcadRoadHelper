using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(GroundToBigTable.GroundTableArea))]
namespace GroundToBigTable
{
    public class GroundTableArea
    {
        [CommandMethod("Ground_table_area")]
        public void Ground_table_area()
        {
            var d = Application.DocumentManager.MdiActiveDocument;
            var editor = d.Editor;

            var result = editor.GetEntity("Выделите таблицу\n");

            if (result.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
            {
                editor.WriteMessage("Ошибка!, завершение программы. Дай Бог вам здоровья\n");
                return;
            }

            var numberOfRows = GetNumberOfRows(d, result);

            for (var i = 1; i < numberOfRows - 4; i++)
            {
                var tr = d.Database.TransactionManager.StartTransaction();
                using (tr)
                {
                    var table = tr.GetObject(result.ObjectId, OpenMode.ForWrite, false) as Table;
                    var gruntText = table.Cells[i, 0].TextString;
                    if (gruntText != string.Empty)
                    {
                        var contourResult = editor.GetEntity($"Выделите контур {gruntText}:\n");if (result.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            editor.WriteMessage("Ошибка!, завершение программы. Дай Бог вам здоровья\n");
                            break;
                        }
                        var polyLine = tr.GetObject(contourResult.ObjectId, OpenMode.ForRead, false) as Polyline;
                        if (polyLine != null)
                        {
                            var objectId = polyLine.Id
                                .ToString()
                                .Replace("(", "")
                                .Replace(")", "");

                            var formula =
                                $"%<\\AcObjProp.16.2 Object(%<\\_ObjId {objectId}>%).Area \\f \"%lu2%pr2%ps[, м2]%ct8[0.04]\">%";
                            table.Cells[i, 1].TextString = formula;
                        }
                        else
                        {
                            editor.WriteMessage("Неверный выбор, попробуйте ещё раз...\n");
                            i -= 1;
                        }
                    }


                    tr.Commit();
                }
            }

        }

        private static int GetNumberOfRows(Document d, PromptEntityResult result)
        {
            int numberOfRows;
            var tr = d.Database.TransactionManager.StartTransaction();

            using (tr)
            {
                var table = tr.GetObject(result.ObjectId, OpenMode.ForWrite, false) as Table;
                numberOfRows = table.Rows.Count;
            }

            return numberOfRows;
        }
    }
}
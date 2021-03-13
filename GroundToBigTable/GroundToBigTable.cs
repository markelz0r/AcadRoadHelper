using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(GroundToBigTable.GroundToBigTable))]
namespace GroundToBigTable
{
    public partial class GroundToBigTable
    {
        [CommandMethod("Ground_to_bigTable")]
        public void Ground_to_bigTable()
        {
            var d = Application.DocumentManager.MdiActiveDocument;
            var editor = d.Editor;

            var result = editor.GetEntity("Выделите общую таблицу\n");

            if (result.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
            {
                editor.WriteMessage("Ошибка!, завершение программы. Дай Бог вам здоровья\n");
                return;
            }

            Table table;
            var tr = d.Database.TransactionManager.StartTransaction();

            using (tr)
            {
                table = tr.GetObject(result.ObjectId, OpenMode.ForWrite, false) as Table;

                    var columns = table.Columns;

                    UpdateTableWithValues(editor, table, tr, columns);
                    tr.Commit();
            }
        }

        private static void UpdateTableWithValues(Autodesk.AutoCAD.EditorInput.Editor editor, Table table, Transaction tr, ColumnsCollection columns)
        {
            for (int i = 1; i < int.MaxValue; i++)
            {
                try
                {
                    CommonMethods.FillPiket(editor, table, tr, i);
                    ProcessSmallTable(editor, table, tr, columns, i);
                }
                catch (Exception)
                {
                    editor.WriteMessage("Завершение программы. Дай Бог вам здоровья\n");
                    break;
                }
            }
        }

        private static void ProcessSmallTable(Autodesk.AutoCAD.EditorInput.Editor editor, Table table, Transaction tr, ColumnsCollection columns, int i)
        {
            var smallTable = editor.GetEntity("Выделите малую таблицу\n");
            if (smallTable.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                throw new Exception();

            var smallTableObj = tr.GetObject(smallTable.ObjectId, OpenMode.ForWrite, false) as Table;

            var numOfRow = smallTableObj.Rows.Count;
            for (int j = 1; j < numOfRow - 4; j++)
            {
                if (smallTableObj.Cells[j, 0].TextString != string.Empty)
                {
                    var grunt = smallTableObj.Cells[j, 0].TextString;
                    var area = smallTableObj.Cells[j, 1].TextString;

                    for (int k = 1; k < columns.Count; k++)
                    {
                        var title = table.Cells[0, k].TextString;
                        var formattedArea = area.Replace(".", ",").Replace("м2", "");
                        if (title != string.Empty)
                        {
                            if (title == grunt)
                            {
                                table.Cells[i, k].TextString = formattedArea;
                                break;
                            }
                        }
                        else
                        {
                            table.Cells[0, k].TextString = grunt;
                            table.Cells[i, k].TextString = formattedArea;
                            break;
                        }
                    }
                }
            }
        }
    }
    
}

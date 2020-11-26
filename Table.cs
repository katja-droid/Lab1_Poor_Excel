using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace PoorExcel
{
    public class Table
    {
        public const int defaultCol = 10;
        public const int defaultRow = 10;
        public int colCount;
        public int rowCount;
        public static List<List<Cell>> grid = new List<List<Cell>>();
        public Dictionary<string, string> dictionary = new Dictionary<string, string>();

        public Table()
        {

            setTable(defaultCol, defaultRow);

        }
        public void setTable(int col, int row)
        {
            Clear();
            colCount = col;
            rowCount = row;
            for (int i = 0; i < rowCount; i++)
            {
                List<Cell> newRow = new List<Cell>();
                for (int j = 0; j < colCount; j++)
                {
                    newRow.Add(new Cell(i, j));
                    dictionary.Add(newRow.Last().getName(), "");
                }
                grid.Add(newRow);
            }
        }
        public void Clear()
        {
            foreach (List<Cell> list in grid)
            {
                list.Clear();
            }
            grid.Clear();
            dictionary.Clear();
            rowCount = 0;
            colCount = 0;
        }
        public void ChangeCellWithAllPointers(int row, int col, string expression, System.Windows.Forms.DataGridView dataGridView1)
        {

            grid[row][col].DeletePointersAndReferences();
            grid[row][col].expression = expression;
            grid[row][col].new_referencesFromThis.Clear();

            string new_expression = ConvertReferences(row, col, expression);
            if (new_expression != "")
            {
                new_expression = new_expression.Remove(0, 1);
            }
            if (!grid[row][col].CheckLoop(grid[row][col].new_referencesFromThis))
            {
                System.Windows.Forms.MessageBox.Show("Виник цикл! Змініть вираз.");
                grid[row][col].expression = "";
                grid[row][col].value = "0";
                dataGridView1[col, row].Value = "0";
                return;
            }
            grid[row][col].AddPointersAndReferences();
            string val = Calculate(new_expression);
            if (val == "Помилка")
            {
                System.Windows.Forms.MessageBox.Show("Помилка в клітинці" + FullName(row, col));
                grid[row][col].expression = "";
                grid[row][col].value = "0";
                dataGridView1[col, row].Value = "0";
                return;
            }
            grid[row][col].value = val;
            dictionary[FullName(row, col)] = val;
            foreach (Cell cell in grid[row][col].pointersToThis)
                RefreshCellAndPointers(cell, dataGridView1);
        }
        private string FullName(int row, int col)
        {
            Cell cell = new Cell(row, col);
            return cell.getName();
        }
        private bool RefreshCellAndPointers(Cell cell, System.Windows.Forms.DataGridView dataGridView1)
        {
            cell.new_referencesFromThis.Clear();
            string new_expression = ConvertReferences(cell.row, cell.column, cell.expression);
            new_expression = new_expression.Remove(0, 1);
            string Value = Calculate(new_expression);
            if (Value == "Помилка")
            {
                System.Windows.Forms.MessageBox.Show("Помилка в Клітинці" + cell.getName());
                cell.expression = "";
                cell.value = "0";
                dataGridView1[cell.column, cell.row].Value = "0";
                return false;
            }
            grid[cell.row][cell.column].value = Value;
            dictionary[FullName(cell.row, cell.column)] = Value;
            dataGridView1[cell.column, cell.row].Value = Value;

            foreach (Cell point in cell.pointersToThis)
            {
                if (!RefreshCellAndPointers(point, dataGridView1))
                    return false;
            }

            return true;


        }
        public void RefreshReferences()
        {
            foreach (List<Cell> list in grid)
            {
                foreach (Cell cell in list)
                {
                    if (cell.referencesFromThis != null)
                        cell.referencesFromThis.Clear();
                    if (cell.new_referencesFromThis != null)
                        cell.new_referencesFromThis.Clear();
                    if (cell.expression == "")
                        continue;
                    string new_expression = cell.expression;
                    if (cell.expression[0] == '=')
                    {
                        new_expression = ConvertReferences(cell.row, cell.column, cell.expression);
                        cell.referencesFromThis.AddRange(cell.new_referencesFromThis);
                    }
                }
            }
        }
        public string ConvertReferences(int row, int col, string expr)
        {
            string cellPattern = @"[A-Z]+[0-9]+";
            Regex regex = new Regex(cellPattern, RegexOptions.IgnoreCase);
            Index nums;
            foreach (Match match in regex.Matches(expr))
            {
                if (dictionary.ContainsKey(match.Value))
                {
                    nums = NumberConverter.From26System(match.Value);
                    grid[row][col].new_referencesFromThis.Add(grid[nums.row][nums.column]);
                }
            }
            MatchEvaluator evaluator = new MatchEvaluator(referenceToValue);
            string new_expression = regex.Replace(expr, evaluator);
            return new_expression;
        }
        public string referenceToValue(Match m)
        {
            if (dictionary.ContainsKey(m.Value))
                if (dictionary[m.Value] == "")
                    return "0";
                else
                    return dictionary[m.Value];
            return m.Value;
        }
        public string Calculate(string expression)
        {
            string res = null;
            try
            {
                res = Convert.ToString(Calculator.Evaluate(expression));
                if (res == "∞" || res == "-∞")
                {
                    res = "Ділення на нуль";
                }
                return res;
            }
            catch
            {
                return "Помилка";
            }
        }
        public void AddRow(System.Windows.Forms.DataGridView dataGridView1)
        {
            List<Cell> newRow = new List<Cell>();
            for (int j = 0; j < colCount; j++)
            {
                newRow.Add(new PoorExcel.Cell(rowCount, j));
                dictionary.Add(newRow.Last().getName(), "");
            }
            grid.Add(newRow);
            RefreshReferences();
            foreach (List<Cell> list in grid)
            {
                foreach (Cell cell in list)
                {
                    if (cell.referencesFromThis != null)
                    {
                        foreach (Cell cell_ref in cell.referencesFromThis)
                        {
                            if (cell_ref.row == rowCount)
                            {
                                if (!cell_ref.pointersToThis.Contains(cell))
                                    cell_ref.pointersToThis.Add(cell);
                            }
                        }
                    }
                }
            }
            for (int j = 0; j < colCount; j++)
            {
                ChangeCellWithAllPointers(rowCount, j, "", dataGridView1);
            }
            rowCount++;
        }

        public void AddColumn(System.Windows.Forms.DataGridView dataGridView)
        {
            List<Cell> newCol = new List<Cell>();

            for (int j = 0; j < rowCount; j++)
            {
                string name = FullName(j, colCount);
                grid[j].Add(new Cell(j, colCount));
                dictionary.Add(name, "");
            }
            RefreshReferences();
            foreach (List<Cell> list in grid)
            {
                foreach (Cell cell in list)
                {
                    if (cell.referencesFromThis != null)
                    {
                        foreach (Cell cell_ref in cell.referencesFromThis)
                        {
                            if (cell_ref.column == colCount)
                            {
                                if (!cell_ref.pointersToThis.Contains(cell))
                                    cell_ref.pointersToThis.Add(cell);
                            }
                        }
                    }
                }
            }
            for (int j = 0; j < rowCount; j++)
            {
                ChangeCellWithAllPointers(j, colCount, "", dataGridView);
            }
            colCount++;
        }

        public bool DeleteRow(System.Windows.Forms.DataGridView dataGridView1)
        {
            List<Cell> lastRow = new List<Cell>();
            List<String> notEmptyCells = new List<string>();
            if (rowCount == 0)
                return false;
            int currCount = rowCount - 1;
            for (int i = 0; i < colCount; i++)
            {
                string name = FullName(currCount, i);
                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(name);
                if (grid[currCount][i].pointersToThis.Count() != 0)
                    lastRow.AddRange(grid[currCount][i].pointersToThis);
            }
            if (lastRow.Count() != 0 || notEmptyCells.Count() != 0)
            {
                string errorMessage = "";
                if (notEmptyCells.Count() != 0)
                {
                    errorMessage = "Не всі клітинки пусті:";
                    errorMessage += string.Join(";", notEmptyCells.ToArray());
                    errorMessage += ' ';
                }
                if (lastRow.Count != 0)
                {
                    errorMessage += "Є клітинки,які посилаються на клітинки з даного рядку :";
                    foreach (Cell cell in lastRow)
                    {
                        errorMessage += string.Join(";", cell.getName());
                        errorMessage += " ";
                    }
                }
                errorMessage += "Ви впевнені,що хочете видалити даний рядок?";
                System.Windows.Forms.DialogResult res = System.Windows.Forms.MessageBox.Show(errorMessage, "Попередження", MessageBoxButtons.YesNo);
                if (res == System.Windows.Forms.DialogResult.No)
                    return false;
            }
            for (int i = 0; i < colCount; i++)
            {
                string name = FullName(currCount, i);
                dictionary.Remove(name);
            }
            foreach (Cell cell in lastRow)
                RefreshCellAndPointers(cell, dataGridView1);
            grid.RemoveAt(currCount);
            rowCount--;
            return true;
        }
        public bool DeleteColumn(System.Windows.Forms.DataGridView dataGridView1)
        {

            List<Cell> lastCol = new List<Cell>();
            List<string> notEmptyCells = new List<string>();
            if (colCount == 0)
                return false;
            int currCount = colCount - 1;
            for (int i = 0; i < rowCount; i++)
            {
                string name = FullName(i, currCount);
                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(name);

                if (grid[i][currCount].pointersToThis.Count != 0)
                    lastCol.AddRange(grid[i][currCount].pointersToThis);
            }
            if (lastCol.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";
                if (notEmptyCells.Count() != 0)
                {
                    errorMessage = "Не всі клітинки пусті:";
                    errorMessage += string.Join("; ", notEmptyCells.ToArray());
                }
                if (lastCol.Count() != 0)
                {
                    errorMessage += "Є клітинки,що посилаються на клітинки з даної колонки:";
                    foreach (Cell cell in lastCol)
                        errorMessage += string.Join("; ", cell.getName());
                }
                errorMessage += "Ви впевнені,що хочете видалити даний стовпчик?";
                System.Windows.Forms.DialogResult res = System.Windows.Forms.MessageBox.Show(errorMessage, "Попередження", System.Windows.Forms.MessageBoxButtons.YesNo);
                if (res == System.Windows.Forms.DialogResult.No)
                    return false;
            }
            for (int i = 0; i < rowCount; i++)
            {
                string name = FullName(i, currCount);
                dictionary.Remove(name);
            }

            foreach (Cell cell in lastCol)
                RefreshCellAndPointers(cell, dataGridView1);

            for (int i = 0; i < rowCount; i++)
            {
                string name = FullName(currCount, i);
                grid[i].RemoveAt(currCount);
            }
            colCount--;
            return true;
        }
        public void Open(int r, int c, System.IO.StreamReader sr, System.Windows.Forms.DataGridView dataGridView)
        {

            for (int i = 0; i < r; i++)
            {
                for (int j = 0; j < c; j++)
                {
                    string index = sr.ReadLine();
                    string expression = sr.ReadLine();
                    string value = sr.ReadLine();
                    if (expression != "")
                        dictionary[index] = value;
                    else
                        dictionary[index] = "";
                    int refCount = Convert.ToInt32(sr.ReadLine());
                    List<Cell> newRef = new List<Cell>();
                    string refer;
                    for (int k = 0; k < refCount; k++)
                    {
                        refer = sr.ReadLine();
                        if (NumberConverter.From26System(refer).row < rowCount && NumberConverter.From26System(refer).column < colCount)
                            newRef.Add(grid[NumberConverter.From26System(refer).row][NumberConverter.From26System(refer).column]);
                    }
                    int pointCount = Convert.ToInt32(sr.ReadLine());
                    List<Cell> newPoint = new List<Cell>();
                    string point;
                    for (int k = 0; k < pointCount; k++)
                    {
                        point = sr.ReadLine();
                        newPoint.Add(grid[NumberConverter.From26System(point).row][NumberConverter.From26System(point).column]);
                    }
                    grid[i][j].setCell(expression, value, newRef, newPoint);

                    int currCol = grid[i][j].column;
                    int currRow = grid[i][j].row;
                    dataGridView[currCol, currRow].Value = dictionary[index];
                }
            }
        }
        public void Save(System.IO.StreamWriter sw)
        {
            sw.WriteLine(rowCount);
            sw.WriteLine(colCount);
            foreach (List<Cell> list in grid)
            {
                foreach (Cell cell in list)
                {
                    sw.WriteLine(cell.getName());
                    sw.WriteLine(cell.expression);
                    sw.WriteLine(cell.value);
                    if (cell.referencesFromThis == null)
                        sw.WriteLine("0");
                    else
                    {
                        sw.WriteLine(cell.referencesFromThis.Count);
                        foreach (Cell refCell in cell.referencesFromThis)
                            sw.WriteLine(refCell.getName());
                    }
                    if (cell.pointersToThis == null)
                        sw.WriteLine("0");
                    else
                    {
                        sw.WriteLine(cell.pointersToThis.Count);
                        foreach (Cell pointCell in cell.pointersToThis)
                            sw.WriteLine(pointCell.getName());
                    }
                }
            }
        }

    }
}

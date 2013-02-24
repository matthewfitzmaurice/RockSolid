using System;
using System.Collections.Generic;
using System.Linq;
using Office = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;

namespace RockSolidOffice
{
    public class InvestmentSchedule
    {
        public static void Update(Wd.Document doc)
        {
            Wd.Table table = GetTable(doc);
            AddHeadingsToTable(doc, table);
            DeleteUnwantedRows(doc, table);
            ReorderRows(table);
            UpdateFormulas(table);
        }

        static Wd.Table GetTable(Wd.Document doc)
        {
            const string FeeSchedule = "Fee Schedule";
            foreach (Wd.Table table in doc.Tables)
                if (table.Descr == FeeSchedule)
                    return table;
            throw new InvalidOperationException(string.Format("Unable to find table '{0}' in document '{1}'", FeeSchedule, doc.Name));
        }

        static void AddHeadingsToTable(Wd.Document doc, Wd.Table table)
        {
            Wd.Range range = doc.Range();
            Wd.Style style = doc.Styles[Wd.WdBuiltinStyle.wdStyleHeading2];
            range = FindRange(range, style);

            while (range.Find.Found)
            {
                string text = GetTextFromHeading(range);
                string number = GetNumberFromHeading(range);
                if (IsTextAlreadyInTable(table, text))
                {
                    if (IsNumberDifferent(table, text, number))
                        UpdateNumber(table, text, number);
                }
                else
                    AddRow(table, text, number);

                range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
                range.End = doc.Range().End;
                range = FindRange(range, style);
            }
        }

        static Wd.Range FindRange(Wd.Range range, Wd.Style style, string value = "")
        {
            Wd.Range search = range.Duplicate;
            search.Find.ClearFormatting();
            search.Find.set_Style(style);
            search.Find.Text = value;
            search.Find.Format = true;
            search.Find.Wrap = Wd.WdFindWrap.wdFindStop;
            search.Find.Format = true;
            search.Find.MatchCase = false;
            search.Find.MatchWholeWord = false;
            search.Find.MatchWildcards = false;
            search.Find.MatchSoundsLike = false;
            search.Find.MatchAllWordForms = false;
            search.Find.Execute();
            return search;
        }

        static string GetTextFromHeading(Wd.Range range)
        {
            return range.Text.Substring(0, range.Text.Length - 1);
        }

        static string GetTextFromRow(Wd.Row row)
        {
            string value = row.Cells[1].Range.Text;
            if (!value.Contains('\t'))
                throw new InvalidOperationException("Unable to update Investment Schedule table.\nOne of the rows is missing a Tab character.");

            value = value.Substring(value.IndexOf('\t') + 1);
            value = value.Substring(0, value.Length - 2);
            return value;
        }

        static string GetNumberFromHeading(Wd.Range range)
        {
            return range.ListFormat.ListString;
        }

        static int GetNumberFromRow(Wd.Row row)
        {
            string value = row.Cells[1].Range.Text;
            if (!value.Contains('\t'))
                throw new InvalidOperationException("Unable to update Investment Schedule table.\nOne of the rows is missing a Tab character.");
            value = value.Substring(0, value.IndexOf('\t'));
            if (!value.Contains('.'))
                throw new InvalidOperationException("Unable to update Investment Schedule table.\nOne of the rows is missing a '.' character.");
            value = value.Substring(value.IndexOf('.') + 1);
            return int.Parse(value);
        }

        static bool IsTextAlreadyInTable(Wd.Table table, string text)
        {
            foreach (Wd.Row row in table.Rows)
                if (row.Cells[1].Range.Text.Contains(text))
                    return true;
            return false;
        }

        static bool IsNumberDifferent(Wd.Table table, string text, string number)
        {
            foreach (Wd.Row row in table.Rows)
                if (row.Cells[1].Range.Text.Contains(text))
                    return !row.Cells[1].Range.Text.Contains(number);
            throw new InvalidOperationException(string.Format("Unable to find text '{0}'.", text));
        }

        static void UpdateNumber(Wd.Table table, string text, string number)
        {
            foreach (Wd.Row row in table.Rows)
                if (row.Cells[1].Range.Text.Contains(text))
                {
                    if (!row.Cells[1].Range.Text.Contains('\t'))
                        throw new InvalidOperationException("Unable to update Investment Schedule table.\nOne of the rows is missing a Tab character.");
                    Wd.Range range = row.Cells[1].Range;
                    range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
                    range.MoveEndUntil('\t');
                    range.Text = number;
                    return;
                }
        }

        static void AddRow(Wd.Table table, string text, string number)
        {
            Wd.Row row = table.Rows.Add(table.Rows[2]);
            row.Cells[1].Range.Text = String.Format("{0}\t{1}", number, text);
            row.Cells[2].Range.Text = "1";
            Wd.Range range = row.Cells[3].Range;
            range.Collapse(Wd.WdCollapseDirection.wdCollapseStart);
            range.Fields.Add(range, Wd.WdFieldType.wdFieldEmpty, @" =B2*0 \# ""$#,##0.00;($#,##0.00)""");
        }

        static void DeleteUnwantedRows(Wd.Document doc, Wd.Table table)
        {
            for (int i = table.Rows.Count - 1; i > 1; i--) // Loop backwards because we're deleting. Not including first and last rows
            {
                Wd.Row row = table.Rows[i];
                if (IsEmpty(row) || !Heading2Exists(doc, GetTextFromRow(row)))
                    row.Delete();
            }
        }

        static bool IsEmpty(Wd.Row row)
        {
            foreach (Wd.Cell cell in row.Cells)
            {
                string value = cell.Range.Text;
                value = value.Substring(0, value.Length - 2);
                if (!string.IsNullOrEmpty(value))
                    return false;
            }
            return true;
        }

        static bool Heading2Exists(Wd.Document doc, string value)
        {
            return FindRange(doc.Range(), doc.Styles[Wd.WdBuiltinStyle.wdStyleHeading2], value).Find.Found;
        }

        static void ReorderRows(Wd.Table table)
        {
            for (int i = 1; i < table.Rows.Count - 1; i++) // Repeat lots of times until the table has been sorted
                for (int j = table.Rows.Count - 1; j > 2; j--) // Not including row one, two and the last row.
                {
                    Wd.Row row = table.Rows[j];
                    if (GetNumberFromRow(row) < GetNumberFromRow(row.Previous))
                        SwapRows(row, row.Previous);
                }
        }

        static void SwapRows(Wd.Row row1, Wd.Row row2)
        {
            row1.Range.Copy();
            row2.Range.Paste();
            row1.Delete();
        }

        static void UpdateFormulas(Wd.Table table)
        {
            foreach (Wd.Field field in table.Range.Fields)
            {
                if (!field.Code.Text.Contains("SUM"))
                {
                    string value = field.Code.Text;
                    value = value.Substring(0, value.IndexOf("=B") + 2) + field.Result.Rows[1].Index + value.Substring(value.IndexOf('*'));
                    field.Code.Text = value;
                }
                field.Update();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using CPI_Beta_v1.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace CPI_Beta_v1
{
    public class ExcelBuilder
    {
        readonly string[] _months = { "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO" };
        private int _lastInterventionPosition;
        readonly string[] _cellPosition = { "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" };

        /// <summary>
        /// Generates an Excel Table Plan for all interventions for a specific year.
        /// </summary>
        /// <param name="interventionsList">List with all interventions.</param>
        /// <param name="interventionYear">Plan year.</param>
        public void GenerateExcel(List<Intervention> interventionsList, int interventionYear)
        {
            try
            {
                //Start Excel and get Application object.
                var oXl = new Excel.Application { Visible = true };

                //Get a new workbook.
                Excel._Workbook oWb = oXl.Workbooks.Add(Missing.Value);

                var oSheet = (Excel._Worksheet)oWb.ActiveSheet;

                _lastInterventionPosition = interventionsList.Count + 3;

                var interventionsIdDescription = new string[interventionsList.Count, 2];
                var index = 0;
                //Builds an bidimensional array with the identifier and description of each intervention
                foreach (var intervention in interventionsList)
                {
                    interventionsIdDescription[index, 0] = intervention.Identifier;
                    interventionsIdDescription[index, 1] = intervention.Description;
                    index++;

                }

                //Headers 2
                oSheet.Cells[2, 1] = "INTERVENÇÃO";
                oSheet.Cells[2, 3] = "PERIODICIDADE";

                //Headers 3
                oSheet.Cells[3, 1] = "Nº";
                oSheet.Cells[3, 2] = "Descrição";
                oSheet.Cells[3, 3] = "[Dias]";


                //STYLES********************************
                var oRng = oSheet.Range["A2", "C2"];
                oRng.EntireColumn.AutoFit();

                Header2Style(oRng);

                oRng = oSheet.Range["A3", "C3"];
                Header3Style(oRng);

                //Fill A4:B... with an array of values (Identifier and Description).
                oSheet.Range["A4", "B" + _lastInterventionPosition].Value2 = interventionsIdDescription;

                oRng = oSheet.Range["A4", "C" + _lastInterventionPosition];
                CommonStyle(oRng);


                //Specific modifications.
                oRng = oSheet.Range["B1"];
                oRng.EntireColumn.AutoFit();

                oRng = oSheet.Range["B4", "B" + _lastInterventionPosition];
                oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                oRng = oSheet.Range["A2", "B2"];
                oRng.Merge();

                //STYLES********************************

                //Configurations to get all weeks of this year
                var jan1 = new DateTime(interventionYear, 1, 1);
                var startOfFirstWeek = jan1.AddDays(1 - (int)(jan1.DayOfWeek));

                //will store the last column of the table
                string finalCell;

                //Reset variable to handle cells position ahead
                index = 0;

                //For each months of the year get the correspondent weeks
                for (var j = 1; j <= 12; j++)
                {
                    //Finds the correspondent weeks for a month
                    var weeksOfMonth =
                   Enumerable
                       .Range(0, 54)
                       .Select(y => new
                       {
                           weekStart = startOfFirstWeek.AddDays(y * 7)
                       })
                       .TakeWhile(x => x.weekStart.Year <= jan1.Year)
                       .Select(x => new
                       {
                           x.weekStart,
                           weekFinish = x.weekStart.AddDays(6)
                       })
                       .SkipWhile(x => x.weekFinish < jan1.AddDays(1))
                       .Select((x, y) => new
                       {
                           x.weekStart,
                           x.weekFinish,
                           weekNum = y + 1
                       }).Where(x => (x.weekStart.Month == j && x.weekStart.Year == interventionYear) ||
                           (x.weekFinish.Month == j && x.weekFinish.Year == interventionYear))
                  .Select(x => x.weekNum).ToList();

                    //Save the initial cell position where the month begins
                    var initialCell = _cellPosition[index] + "2";

                    //Writes every week in excel
                    foreach (var week in weeksOfMonth)
                    {
                        oRng = oSheet.Range[_cellPosition[index] + "3"].Cells[1, 1];
                        oRng.Cells.Value2 = week;
                        Header3Style(oRng);
                        index++;
                    }
                    //Save the final cell position where the month ends
                    finalCell = _cellPosition[index - 1] + "2";

                    //Merge the initial cell with the final, write the month name and apply the correspondent style
                    oRng = oSheet.Range[initialCell, finalCell];
                    oRng.Merge();
                    oRng.ColumnWidth = 2;
                    oRng.Cells[1, 1] = _months[j - 1];
                    Header2Style(oRng);




                }
                //Apply final style to the entire table
                finalCell = _cellPosition[index - 1] + (interventionsList.Count + 3);
                oRng = oSheet.Range["A1", finalCell];
                oRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //Style for the title of the table
                oRng = oSheet.Range["A1", _cellPosition[index - 1] + "1"];
                oRng.Merge();
                oRng.Cells[1, 1] = "PLANO DE INTERVENÇÃO PREVENTIVA SIE " + interventionYear;
                Header1Style(oRng);

                //The interventions strat column
                var startPosition = 4;

                //Bidimensional array that will store the periodicity of each intervention
                var periodicity = new int[interventionsList.Count, 1];

                var indexPeriodicity = 0;

                foreach (var intervention in interventionsList)
                {
                    var lastCell = 0;
                    var month = 0;

                    //Estimation of the  periodicity of an intervention
                    if (intervention.MarkedInterventionsList.Count == 1)
                    {
                        periodicity[indexPeriodicity, 0] = 365;
                    }
                    else
                    {
                        var list = intervention.MarkedInterventionsList.Take(2).ToList();
                        periodicity[indexPeriodicity, 0] = Math.Abs((int)(list.First() - list.Last()).TotalDays);

                    }
                    indexPeriodicity++;

                    foreach (var interventionDate in intervention.MarkedInterventionsList)
                    {

                        int week = GetWeek(interventionDate);
                        for (var j = lastCell; j < index; j++)
                        {
                            //Estimation of the  periodicity of an intervention
                            var value2 = oSheet.Range[_cellPosition[j] + "2"].Value2;
                            if (value2 != null)
                            {
                                month++;
                            }
                            if (oSheet.Range[_cellPosition[j] + "3"].Value2 != week || interventionDate.Month != month) continue;

                            //Write the specific day of the intervention in the correspondent week and style it
                            oSheet.Range[_cellPosition[j] + startPosition].Cells[1, 1] = interventionDate.Day;
                            oRng = oSheet.Range[_cellPosition[j] + startPosition];
                            oRng.Cells.Style = "Good";
                            oRng.Font.Name = "Calibri";
                            oRng.Font.Size = 10;
                            oRng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                            //Store the last cell position that the next iteration will start for the weeks
                            lastCell = j + 1;
                            break;
                        }
                    }
                    startPosition++;
                }


                FormatPeriodicity(interventionsList, oSheet, periodicity);

            }
            catch (Exception theException)
            {
                var errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, Resources.ExcelBuilder_GenerateExcel_Error);
            }
        }

        private static void FormatPeriodicity(List<Intervention> interventionsList, Excel._Worksheet oSheet, int[,] periodicity)
        {
            var oRng = oSheet.Range["C4", "C" + (interventionsList.Count + 3)];
            oRng.Value2 = periodicity;
            oRng.Font.Name = "Calibri";
            oRng.Font.Size = 10;
            oRng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        private static void Header1Style(Excel.Range oRng)
        {
            CommonHeadersStyle(oRng);
            oRng.Font.Color = Color.White;
            oRng.Cells.Interior.Color = Color.FromArgb(49, 134, 155);

        }


        private static void Header3Style(Excel.Range oRng)
        {

            CommonHeadersStyle(oRng);
            oRng.Cells.Interior.Color = Color.FromArgb(217, 217, 217);
        }

        private static void Header2Style(Excel.Range oRng)
        {
            CommonHeadersStyle(oRng);
            oRng.Font.Color = Color.FromArgb(74, 134, 168);
            oRng.Cells.Interior.Color = Color.FromArgb(183, 222, 232);
        }

        private static void CommonHeadersStyle(Excel.Range oRng)
        {
            CommonStyle(oRng);
            oRng.Font.Bold = true;
        }

        private static void CommonStyle(Excel.Range oRng)
        {
            oRng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oRng.Font.Name = "Calibri";
            oRng.Font.Size = 10;
        }

        private static Int16 GetWeek(DateTime date1)
        {
            var dfi = DateTimeFormatInfo.CurrentInfo;
            if (dfi == null) return -1;
            var cal = dfi.Calendar;

            return (short)cal.GetWeekOfYear(date1, dfi.CalendarWeekRule,
                dfi.FirstDayOfWeek);
        }
    }
}

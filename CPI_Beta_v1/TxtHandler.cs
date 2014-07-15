using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace CPI_Beta_v1
{
    public class TxtHandler
    {
        private CultureInfo _cultureInfo;
        private string _ddMmmYy;

        /// <summary>
        /// Builds an ordered list with all interventions and their date.
        /// </summary>
        /// <param name="lines">Contains all lines of the txt file export CPC.</param>
        /// <returns>A list of interventions sorted by ascending order.</returns>
        public List<Intervention> BuildInterventions(IEnumerable<string> lines)
        {
            var majorList = new List<Intervention>();

            //The first line of the file is the header.
            var first = true;

            //Split each line.
            foreach (var array in lines.Select(line => line.Split('\t')))
            {
                //To discard the header file.
                if (first || array[2] == string.Empty)
                {
                    first = false;
                    continue;
                }
                //Find an intervention by the identifier in majorList to add dates intervention. 
                //If the result is null then create a new one.
                var intervention = majorList.FirstOrDefault(x => x.Identifier.Equals(array[0]));
                if (intervention == null)
                {
                    intervention = new Intervention { Identifier = array[0], Description = array[2] };
                    try
                    {
                        //Selects the number contained in the identifier to serve in the ranking list.
                        intervention.NumberId = Int16.Parse(Regex.Match(array[0], @"\d+").Value);
                    }
                    catch (FormatException exception)
                    {

                        Console.WriteLine(exception.Message);
                    }

                    majorList.Add(intervention);
                }
                try
                {
                    //Parse the date in portuguese culture.
                    _cultureInfo = new CultureInfo("pt-PT");

                    //Date format pattern.
                    _ddMmmYy = "dd-MMM-yy";

                    intervention.MarkedInterventionsList.Add(DateTime.ParseExact(array[1], _ddMmmYy,
                     _cultureInfo));
                }
                catch (FormatException exception)
                {

                    Console.WriteLine(exception.Message);
                }
            }
            //Sorts the list in ascending order of NumberId.
            return majorList.OrderBy(x => x.NumberId).ToList();

        }
    }
}

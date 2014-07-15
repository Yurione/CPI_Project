using System;
using System.Collections.Generic;

namespace CPI_Beta_v1
{
    public class Intervention
    {

        private String _identifier;
        private String _description;
        private List<DateTime> _markedInterventionsList;
        private Int16 _numberId;


        public Intervention()
        {
            _identifier = string.Empty;
            _description = string.Empty;
            _markedInterventionsList = new List<DateTime>();
            _numberId = 0;


        }

        public string Identifier
        {
            get { return _identifier; }
            set { _identifier = value; }
        }

        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }

        public List<DateTime> MarkedInterventionsList
        {
            get { return _markedInterventionsList; }
            set { _markedInterventionsList = value; }
        }

        public short NumberId
        {
            get { return _numberId; }
            set { _numberId = value; }
        }

    }
}

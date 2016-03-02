using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class AutoNumber
    {

        #region Propriedades

        public string fullName { get; set; }
        public string description { get; set; }
        public string displayFormat { get; set; }
        public bool externalId { get; set; }
        public string inlineHelpText { get; set; }
        public string label { get; set; }
        public string type { get; set; }

        #endregion

        #region Construtores

        public AutoNumber() 
        {
        
        }

        #endregion

    }
}

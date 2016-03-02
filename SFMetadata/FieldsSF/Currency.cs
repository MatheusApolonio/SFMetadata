using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class Currency
    {

        #region Propriedades

        public string fullName { get; set; }
        public string defaultValue { get; set; }
        public string description { get; set; }
        public string inlineHelpText { get; set; }
        public string label { get; set; }
        public int precision { get; set; }
        public bool required { get; set; }
        public int scale { get; set; }
        public string type { get; set; }

        #endregion

        #region Construtores

        public Currency() 
        {
        
        }

        #endregion

    }
}

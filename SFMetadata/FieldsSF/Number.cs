using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class Number
    {

        #region Propriedades

        public string fullName { get; set; }
        public string defaultValue { get; set; }
        public string description { get; set; }
        public bool externalId { get; set; }
        public string inlineHelpText { get; set; }
        public string label { get; set; }
        public int precision { get; set; }
        public bool required { get; set; }
        public int scale { get; set; }
        public bool unique { get; set; }
        public string type { get; set; }

        #endregion

        #region Construtores

        public Number() { }

        #endregion

    }
}

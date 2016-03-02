using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class Percent
    {
        #region Propriedades

        public string fullname { get; set; }
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

        public Percent() { }

        #endregion
    }
}

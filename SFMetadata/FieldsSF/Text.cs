using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class Text
    {

        #region Propriedades

        public string fullName { get; set; }
        public bool caseSensitive { get; set; }
        public string defaultValue { get; set; }
        public string description { get; set; }
        public bool externalId { get; set; }
        public string inlineHelpText { get; set; }
        public string label { get; set; }
        public int length { get; set; }
        public bool required { get; set; }
        public bool unique { get; set; }
        public string type { get; set; }

        #endregion

        #region Construtores

        public Text() { }

        #endregion

    }
}

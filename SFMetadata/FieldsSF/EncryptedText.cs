using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class EncryptedText
    {

        #region Propriedades

        public string fullName { get; set; }
        public string label { get; set; }
        public int length { get; set; }
        public string maskChar { get; set; }
        public string maskType { get; set; }
        public bool required { get; set; }
        public string type { get; set; }

        #endregion

        #region Construtores

        public EncryptedText() { }

        #endregion

    }
}

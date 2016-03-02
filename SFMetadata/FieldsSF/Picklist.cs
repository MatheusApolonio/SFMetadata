using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class Picklist
    {

        #region Propriedades

        public string fullName { get; set; }
        public string description { get; set; }
        public string inlineHelpText { get; set; }
        public string label { get; set; }
        public string type { get; set; }
        public List<PickListValue> picklist { get; set; }
        public bool sorted { get; set; }

        #endregion

        #region Construtores

        public Picklist() { }

        #endregion

    }
}

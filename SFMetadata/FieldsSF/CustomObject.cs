using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class CustomObject
    {

        #region Propriedades

        public string deploymentStatus { get; set; }
        public string label { get; set; }
        public string gender { get; set; }
        public NewField nameField { get; set; }
        public string pluralLabel { get; set; }
        public string searchLayouts { get; set; }
        public string sharingModel { get; set; }
        public List<Object> fields { get; set; }

        #endregion

        #region Construtores

        public CustomObject() 
        {
        }

        #endregion

    }
}

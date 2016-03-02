using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class Email
    {

        #region Propriedades

        public string fullName { get; set; }
        public string defaultValue { get; set; }
        public bool externalId { get; set; }
        public string label { get; set; }
        public bool required { get; set; }
        public bool unique { get; set; }
        public string type { get; set; }

        #endregion

        #region Construtor

        public Email() 
        {
        
        }

        #endregion

    }
}

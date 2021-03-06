﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class LongTextArea
    {

        #region Propriedades

        public string fullName { get; set; }
        public string defaultValue { get; set; }
        public string description { get; set; }
        public string inlineHelpText { get; set; }
        public string label { get; set; }
        public int length { get; set; }
        public int visibleLines { get; set; }
        public string type { get; set; }

        #endregion

        #region Construtores

        public LongTextArea() { }

        #endregion

    }
}

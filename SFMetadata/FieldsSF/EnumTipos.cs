using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SFMetadata.FieldsSF
{
    public class EnumTipos
    {

        #region Propriedades

        public enumTipoDados tipoDados { get; set; }

        #endregion

        #region Construtor

        public EnumTipos()
        {

        }

        #endregion

        #region Enum

        public enum enumTipoDados
        {
            Checkbox,
            DateTime,
            Date,
            Email,
            Text,
            Location,
            Currency,
            AutoNumber,
            Number,
            MultiselectPicklist,
            Picklist,
            Percent,
            Html,
            Phone,
            LongTextArea,
            TextArea,
            EncryptedText,
            Url
        }
        #endregion
    }
}

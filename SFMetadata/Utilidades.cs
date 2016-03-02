using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Internal;
using SFMetadata.FieldsSF;
using System.Reflection;
using System.IO;

namespace SFMetadata
{
    public class Utilidades
    {
        //Para efeitos de erro durante o processamento, criaremos uma lista de String para guardar os erros que aconteceram durante a leitura dos campos do excel
        private List<string> errosProcesso = new List<string>();

        public delegate bool AsyncMethodCaller(string Arquivo);

        public bool LeArquivo(string Arquivo)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //Cria um objeto do tipo WorkBook com todos os elementos do Excel.
            Microsoft.Office.Interop.Excel.Workbook objWorkbook = app.Workbooks.Open(Arquivo,
            Type.Missing, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
            int numSheets = objWorkbook.Sheets.Count;

            //Criando a variavel lista que irá conter todas as relações de objetos para inserir no SF
            List<CustomObject> listaCustomObj = new List<CustomObject>();

            //Lista que irá conter o nome dos objetos que já processamos
            List<string> objetosProcessados = new List<string>();

            //esse loop vai percorrer todas as pastas de trabalho do excel.
            for (int sheetNum = 2; sheetNum < numSheets + 1; sheetNum++)
            {
                Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkbook.Sheets[sheetNum];
                int numColumns = objSheet.Columns.Count;
                int numRows = objSheet.Rows.Count;


                Microsoft.Office.Interop.Excel.Range excelRange = objSheet.UsedRange;
                //Pega todo conteúdo de uma linha e transforma e um array de objetos.
                object[,] Linha = (object[,])excelRange.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);

                List<string[]> Linhas = new List<string[]>();
                int cont_ant = 0;
                //Percorre todas as dimensões do array linha para pegar o conteúdo de cada célula.
                for (int i = 1; i <= Linha.GetUpperBound(0); i++)
                {
                    if (i > cont_ant)
                    {
                        cont_ant = i;
                        //pega o total de células preenchidas na linha e cria um array com o número exato de dimensões.
                        string[] celulas = new string[Linha.GetUpperBound(1)];
                        //percorre todas as células atribuindo preenchendo o array.
                        for (int j = 1; j <= Linha.GetUpperBound(1); j++)
                        {
                            if (Linha[i, j] != null)
                                celulas[j - 1] = Linha[i, j].ToString();
                        }

                        Linhas.Add(celulas);


                    }
                }

                //Nesse momento o List de string "Linhas" tem todo conteúdo  da pasta de trabalho.

                //Iremos considerar as seguintes colunas para recuperar a informação: 2,4,5,7,9 sendo que:
                /*  2 - Label(PT-BR)
                 *  4 - API Name
                 *  5 - FieldType
                 *  7 - PickList values(PT-BR)
                 *  9 - Field Security
                 */

                int countLinha = 0;

                string nomeCol = "";
                string nomeSheet = "";

                try
                {
                    //Primeiramente, iremos verificar se o item dessa Sheet já foi processado, se sim, iremos apenas fazer o merge dos campos...
                    if (objetosProcessados.Contains(Linhas[0][1]))
                    {
                        #region Objeto existente

                        //Percorreremos os objetos que já foram adicionados para procurar pelo objeto "novo"
                        foreach (CustomObject custObj in listaCustomObj)
                        {
                            if (custObj.label.Equals(Linhas[0][1]))
                            {
                                //Uma vez encontrado o objeto já tratado, iremos realizar a adição dos campos que ainda não foram adicionados

                                foreach (string[] cols in Linhas)
                                {

                                    nomeCol = cols[2];
                                    nomeSheet = objSheet.Name;

                                    if (countLinha >= 0 && countLinha <= 5)
                                    {
                                        //Pulando linhas desnecessárias
                                        countLinha++;
                                        continue;
                                    }

                                    //A condição de saida do foreach será o valor Lista Relacionada do documento ou a leitura de espaço vazio
                                    if (string.IsNullOrEmpty(cols[0]) || cols[0].Trim().Equals("ListaRelacionada"))
                                    {
                                        break;
                                    }

                                    if (!cols[0].Trim().Equals("Section:"))
                                    {

                                        //Flag que determina se o objeto existe ou não
                                        bool objetoExistente = false;

                                        //Validaremos se o campo do API Name está preenchido
                                        if (!string.IsNullOrEmpty(cols[4]))
                                        {
                                            //Uma vez iniciada a leitura, percorreremos os campos do objeto em questão para ver se ele já foi adicionado

                                            //Percorreremos todos os campos adicionados ao objeto
                                            foreach (Object campo in custObj.fields)
                                            {
                                                //Convertemos o objeto para Type
                                                Type typeObj = campo.GetType();

                                                //Recuperaremos a propriedade 'label' dos campos...
                                                PropertyInfo propInf = typeObj.GetProperty("label");

                                                //Flag que determina se o objeto ja foi adicionado
                                                bool jaAdicionado = false;

                                                //Uma vez encontrada, recuperaremos o value da propriedade 'label'
                                                string nomeCampo = (string)propInf.GetValue(campo, null);

                                                if (cols[2].Equals(nomeCampo.TrimEnd(' ')))
                                                {
                                                    //Se o valor que temos for igual a um objeto encontrado, entao setaremos a flag para true
                                                    jaAdicionado = true;
                                                    objetoExistente = true;
                                                }

                                                bool linhaProcessada = false;

                                                //Se a flag for true, processaremos de acordo com o valor da flag
                                                if (jaAdicionado)
                                                {
                                                    #region Campo existente

                                                    //Verificaremos se é do tipo Picklist ou MultiPicklist

                                                    //Recuperaremos a propriedade 'label' dos campos...
                                                    propInf = typeObj.GetProperty("type");

                                                    //Convertemos o objeto de Type para EnumTipos.enumTipoDados
                                                    string tipoCampo = (string)propInf.GetValue(campo, null);

                                                    //Verificamos se é algum dois dois tipos

                                                    /*Se for qualquer um dos dois tipos, realizaremos o cast do objeto para os objetos correspondentes 
                                                     * e realizaremos a tratativa dos picklistvalues*/
                                                    if (tipoCampo.Equals("Picklist"))
                                                    {
                                                        //Convertendo o Type para Object
                                                        Object obj = (Object)campo;

                                                        //Convertendo o Object para Picklist para facilitar o trabalho com a classe
                                                        Picklist pkList = (Picklist)obj;

                                                        if (!string.IsNullOrEmpty(cols[7]))
                                                        {
                                                            //Recuperamos os valores das celulas do documento excel para validar se existem valores que ja não adicionamos
                                                            string[] picklistValues = cols[7].Split('\n');

                                                            //Percorrendo a lista recuperada
                                                            foreach (string strValue in picklistValues)
                                                            {
                                                                //Flag que determina se iremos adicionar o novo valor no picklist
                                                                bool criarItem = false;

                                                                //Aqui, percorreremos os valores que já foram inseridos no picklist para verificar se o valor encontrado no excel já existe
                                                                foreach (PickListValue pkValue in pkList.picklist)
                                                                {
                                                                    //Validando de o registro existe na lista
                                                                    if (!pkValue.fullname.Equals(strValue.TrimEnd(' ')))
                                                                    {
                                                                        /*Se o nome não bater, inicialmente iremos considerar que o registro não existe e daremos 
                                                                         * OK para a criação do novo item*/
                                                                        criarItem = true;
                                                                    }
                                                                    else
                                                                    {
                                                                        //Se em algum item da lista, encontrarmos o item da planilha, então atribuiremos false a necessidade e sairemos do for
                                                                        criarItem = false;
                                                                        break;
                                                                    }
                                                                }

                                                                //Se a flag for true, então criaremos um novo PickListValue e atribuiremos o mesmo ao picklist já existente
                                                                if (criarItem)
                                                                {
                                                                    //Criando o novo PickListValue
                                                                    PickListValue pkNewVal = new PickListValue();
                                                                    pkNewVal.defaultValue = false;
                                                                    pkNewVal.fullname = strValue.TrimEnd(' ');

                                                                    pkList.picklist.Add(pkNewVal);
                                                                }
                                                            }

                                                            linhaProcessada = true;
                                                        }
                                                    }
                                                    else if (tipoCampo.Equals("MultiselectPicklist"))
                                                    {
                                                        //Convertendo o Type para Object
                                                        Object obj = (Object)campo;

                                                        //Convertendo o Object para MultiselectPicklist para facilitar a manipulação
                                                        MultiselectPicklist pkList = (MultiselectPicklist)obj;

                                                        if (!string.IsNullOrEmpty(cols[7]))
                                                        {
                                                            //Recuperando os valores do picklist no excel
                                                            string[] picklistValues = cols[7].Split('\n');

                                                            //Percorrendo esses valores...
                                                            foreach (string strValue in picklistValues)
                                                            {
                                                                //Flag que determina se iremos adicionar o novo valor no picklist
                                                                bool criarItem = false;

                                                                //Aqui, percorreremos os valores que já foram inseridos no picklist para verificar se o valor encontrado no excel já existe
                                                                foreach (PickListValue pkValue in pkList.picklist)
                                                                {
                                                                    //Validando de o registro existe na lista
                                                                    if (!pkValue.fullname.Equals(strValue.TrimEnd(' ')))
                                                                    {
                                                                        /*Se o nome não bater, inicialmente iremos considerar que o registro não existe e daremos 
                                                                         * OK para a criação do novo item*/
                                                                        criarItem = true;
                                                                    }
                                                                    else
                                                                    {
                                                                        //Se em algum item da lista, encontrarmos o item da planilha, então atribuiremos false a necessidade e sairemos do for
                                                                        criarItem = false;
                                                                        break;
                                                                    }
                                                                }

                                                                //Se a flag for true, então criaremos um novo PickListValue e atribuiremos o mesmo ao picklist já existente
                                                                if (criarItem)
                                                                {
                                                                    //Criando o novo PickListValue
                                                                    PickListValue pkNewVal = new PickListValue();
                                                                    pkNewVal.defaultValue = false;
                                                                    pkNewVal.fullname = strValue.TrimEnd(' ');

                                                                    pkList.picklist.Add(pkNewVal);
                                                                }
                                                            }

                                                            linhaProcessada = true;
                                                        }
                                                    }

                                                    #endregion
                                                }

                                                //Pulamos uma linha caso já tenhamos processado o valor
                                                if (linhaProcessada)
                                                {
                                                    break;
                                                }
                                            }

                                            //Caso o campo não tenha sido adicionado no objeto, iremos adiciona-lo aqui
                                            if (!objetoExistente)
                                            {
                                                #region Campo não existente

                                                //Se a label pt-br estiver vazia, o nome do campo será o mesmo da API name
                                                //TODO: Validar depois se essa é a melhor solução

                                                //Validamos se o campo FieldType está preenchido

                                                if (!string.IsNullOrEmpty(cols[5]))
                                                {
                                                    //Primeiro, determinamos qual o tipo do campo

                                                    string tipoCampo = cols[5];

                                                    //Inicialmente, iremos tratar o campo para tirar parenteses e valores numericos...

                                                    tipoCampo = tipoCampo.Replace("(", "").Replace(")", "").Replace("0", "").Replace("1", "").Replace("2", "").Replace("3", "").Replace("4", "").Replace("5", "");
                                                    tipoCampo = tipoCampo.Replace("6", "").Replace("7", "").Replace("8", "").Replace("9", "").Replace(".", "").Replace(" ", "").Replace(",", ".").ToLower();

                                                    //Agora, criaremos um switch / case para criar o objeto correto de acordo com o tipo do campo

                                                    switch (tipoCampo)
                                                    {
                                                        case "email":
                                                            Email email = new Email();
                                                            email.defaultValue = "";
                                                            email.externalId = false;
                                                            email.fullName = cols[4];
                                                            email.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                email.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            email.type = "Email";
                                                            email.unique = false;

                                                            custObj.fields.Add((Object)email);
                                                            break;

                                                        case "boolean":
                                                            Checkbox chk = new Checkbox();
                                                            chk.defaultValue = (cols[5].Contains("1") ? true : false);
                                                            chk.description = "";
                                                            chk.fullName = cols[4];
                                                            chk.inlineHelpText = "";
                                                            chk.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            chk.type = "Checkbox";

                                                            custObj.fields.Add((Object)chk);
                                                            break;

                                                        case "datetime":
                                                            DateTimeSF dtSF = new DateTimeSF();
                                                            dtSF.defaultvalue = "";
                                                            dtSF.description = "";
                                                            dtSF.fullName = cols[4];
                                                            dtSF.inlineHelpText = "";
                                                            dtSF.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                dtSF.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            dtSF.type = "DateTime";

                                                            custObj.fields.Add((Object)dtSF);
                                                            break;

                                                        case "date":
                                                            DateSF dSF = new DateSF();
                                                            dSF.defaultValue = "";
                                                            dSF.description = "";
                                                            dSF.fullName = cols[4];
                                                            dSF.inlineHelpText = "";
                                                            dSF.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                dSF.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            dSF.type = "Date";

                                                            custObj.fields.Add((Object)dSF);
                                                            break;

                                                        case "location":
                                                            Location lct = new Location();
                                                            lct.description = "";
                                                            lct.fullName = cols[4];
                                                            lct.inlineHelpText = "";
                                                            lct.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                lct.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            lct.scale = 3; //Por default, eu vou deixar valor 3. alterar depois se vier um valor descrito no excel...
                                                            lct.type = "Location";

                                                            custObj.fields.Add((Object)lct);
                                                            break;

                                                        case "currency":
                                                            Currency curr = new Currency();
                                                            curr.defaultValue = "";
                                                            curr.description = "";
                                                            curr.fullName = cols[4];
                                                            curr.inlineHelpText = "";
                                                            curr.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                            //Para recuperar o valor da precisão dos campos
                                                            string lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                            lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                            if (lenStr.Contains('.'))
                                                            {
                                                                curr.precision = Convert.ToInt32(lenStr.Split('.')[0]) + Convert.ToInt32(lenStr.Split('.')[1]);
                                                                curr.scale = Convert.ToInt32(lenStr.Split('.')[1]);
                                                            }
                                                            else
                                                            {
                                                                curr.precision = Convert.ToInt32(lenStr.Split('.')[0]);
                                                                curr.scale = 0;
                                                            }

                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                curr.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            curr.type = "Currency";

                                                            custObj.fields.Add((Object)curr);
                                                            break;

                                                        case "autonumber":
                                                            AutoNumber auto = new AutoNumber();
                                                            auto.description = "";
                                                            auto.displayFormat = "";
                                                            auto.externalId = false; //Por default, eu vou deixar os campos de externalId como false. alterar depois se vier um valor descrito no excel...
                                                            auto.fullName = cols[4];
                                                            auto.inlineHelpText = "";
                                                            auto.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            auto.type = "AutoNumber";

                                                            custObj.fields.Add((Object)auto);
                                                            break;

                                                        case "number":
                                                            Number nb = new Number();
                                                            nb.defaultValue = "";
                                                            nb.description = "";
                                                            nb.externalId = false;
                                                            nb.fullName = cols[4];
                                                            nb.inlineHelpText = "";
                                                            nb.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                            //Para recuperar o valor da precisão dos campos
                                                            lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                            lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                            if (lenStr.Contains('.'))
                                                            {
                                                                nb.precision = Convert.ToInt32(lenStr.Split('.')[0]) + Convert.ToInt32(lenStr.Split('.')[1]);
                                                                nb.scale = Convert.ToInt32(lenStr.Split('.')[1]);
                                                            }
                                                            else
                                                            {
                                                                nb.precision = Convert.ToInt32(lenStr.Split('.')[0]);
                                                                nb.scale = 0;
                                                            }


                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                nb.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            nb.type = "Number";
                                                            nb.unique = false;

                                                            custObj.fields.Add((Object)nb);
                                                            break;

                                                        case "multiselectpicklist":
                                                            MultiselectPicklist multiPick = new MultiselectPicklist();
                                                            multiPick.description = "";
                                                            multiPick.fullName = cols[4];
                                                            multiPick.inlineHelpText = "";
                                                            multiPick.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            multiPick.sorted = false;
                                                            multiPick.visibleLines = 4;
                                                            multiPick.picklist = new List<PickListValue>();

                                                            if (!string.IsNullOrEmpty(cols[7]))
                                                            {
                                                                string[] valoresPickListSingle = cols[7].Split('\n');

                                                                foreach (string valor in valoresPickListSingle)
                                                                {
                                                                    PickListValue pickValue = new PickListValue();
                                                                    pickValue.defaultValue = false;
                                                                    pickValue.fullname = valor.TrimEnd(' ');

                                                                    multiPick.picklist.Add(pickValue);
                                                                }

                                                                multiPick.type = "MultiselectPicklist";

                                                                custObj.fields.Add((Object)multiPick);
                                                            }
                                                            else
                                                            {
                                                                errosProcesso.Add("[" + DateTime.Now + "] Campo " + (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]) + " não será criado pois não possui items");
                                                            }


                                                            break;

                                                        case "picklist":
                                                            Picklist pick = new Picklist();
                                                            pick.description = "";
                                                            pick.fullName = cols[4];
                                                            pick.inlineHelpText = "";
                                                            pick.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            pick.picklist = new List<PickListValue>();

                                                            if (!string.IsNullOrEmpty(cols[7]))
                                                            {
                                                                string[] valoresPickListSingle = cols[7].Split('\n');

                                                                foreach (string valor in valoresPickListSingle)
                                                                {
                                                                    PickListValue pickValue = new PickListValue();
                                                                    pickValue.defaultValue = false;
                                                                    pickValue.fullname = valor.TrimEnd(' ');

                                                                    pick.picklist.Add(pickValue);
                                                                }

                                                                pick.sorted = false;
                                                                pick.type = "Picklist";

                                                                custObj.fields.Add((Object)pick);
                                                            }
                                                            else
                                                            {
                                                                errosProcesso.Add("[" + DateTime.Now + "] Campo " + (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]) + " não será criado pois não possui items");
                                                            }

                                                            break;

                                                        case "percent":
                                                            Percent prct = new Percent();
                                                            prct.defaultValue = "";
                                                            prct.description = "";
                                                            prct.fullname = cols[4];
                                                            prct.inlineHelpText = "";
                                                            prct.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                            //Para recuperar o valor da precisão dos campos
                                                            lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                            lenStr = lenStr.Replace("(", "").Replace(")", "");


                                                            if (lenStr.Contains('.'))
                                                            {
                                                                prct.precision = Convert.ToInt32(lenStr.Split('.')[0]) + Convert.ToInt32(lenStr.Split('.')[1]);
                                                                prct.scale = Convert.ToInt32(lenStr.Split('.')[1]);
                                                            }
                                                            else
                                                            {
                                                                prct.precision = Convert.ToInt32(lenStr.Split('.')[0]);
                                                                prct.scale = 0;
                                                            }

                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                prct.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            prct.type = "Percent";

                                                            custObj.fields.Add((Object)prct);
                                                            break;

                                                        case "html":
                                                            HTML html = new HTML();
                                                            html.description = "";
                                                            html.fullName = cols[4];
                                                            html.inlineHelpText = "";
                                                            html.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                            //Para recuperar o valor da precisão dos campos
                                                            lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                            lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                            html.length = Convert.ToInt32(lenStr); // Valor padrão do SF
                                                            html.visibleLines = 25; //Valor padrão do SF;
                                                            html.type = "Html";

                                                            custObj.fields.Add((Object)html);
                                                            break;

                                                        case "phone":
                                                            Phone phn = new Phone();
                                                            phn.defaultValue = "";
                                                            phn.description = "";
                                                            phn.fullName = cols[4];
                                                            phn.inlineHelpText = "";
                                                            phn.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                phn.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            phn.type = "Phone";

                                                            custObj.fields.Add((Object)phn);
                                                            break;

                                                        case "longtextarea":
                                                            LongTextArea lta = new LongTextArea();
                                                            lta.defaultValue = "";
                                                            lta.description = "";
                                                            lta.fullName = cols[4];
                                                            lta.inlineHelpText = "";
                                                            lta.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                            //Para recuperar o valor da precisão dos campos
                                                            lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                            lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                            lta.length = Convert.ToInt32(lenStr);
                                                            lta.visibleLines = 3;
                                                            lta.type = "LongTextArea";

                                                            custObj.fields.Add((Object)lta);
                                                            break;

                                                        case "textarea":
                                                            TextArea txAr = new TextArea();
                                                            txAr.defaultValue = "";
                                                            txAr.description = "";
                                                            txAr.fullName = cols[4];
                                                            txAr.inlineHelpText = "";
                                                            txAr.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                txAr.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            txAr.type = "TextArea";

                                                            custObj.fields.Add((Object)txAr);
                                                            break;

                                                        case "encryptedtext":
                                                            EncryptedText encText = new EncryptedText();
                                                            encText.fullName = cols[4];
                                                            encText.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                            //Para recuperar o valor da precisão dos campos
                                                            lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                            lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                            encText.length = Convert.ToInt32(lenStr);

                                                            //TODO: Definir qual campo irá conter o tipo de criptografia e qual campo irá conter o caractere de criptografia
                                                            string caractereCripto = "";
                                                            string tipoCripto = "";

                                                            //O SF aceita apenas valores mascarados por 'X' ou '*'
                                                            if (caractereCripto.Equals("X"))
                                                            {
                                                                encText.maskChar = "X";
                                                            }
                                                            else
                                                            {
                                                                encText.maskChar = "asterisk";
                                                            }

                                                            /*   O SF possui os seguintes tipos de mascara:
                                                                *  Mascarar todos os caracteres - all
                                                                *  Limpar quatro últimos caracteres - lastFour
                                                                *  Número do cartão de crédito - creditCard
                                                                *  Número do seguro nacional - nino
                                                                *  Número do CPF - ssn
                                                                *  Número do seguro social - sin
                                                                */

                                                            switch (tipoCripto)
                                                            {
                                                                case "Quatro ultimos":
                                                                    encText.maskType = "lastFour";
                                                                    break;
                                                                case "Cartão de crédito":
                                                                    encText.maskType = "creditCard";
                                                                    break;
                                                                case "Seguro nacional":
                                                                    encText.maskType = "nino";
                                                                    break;
                                                                case "CPF":
                                                                    encText.maskType = "ssn";
                                                                    break;
                                                                case "Seguro social":
                                                                    encText.maskType = "sin";
                                                                    break;
                                                                default:
                                                                    //Por padrão, se não for informado um tipo valido, iremos considerar a criptografia total da string
                                                                    encText.maskType = "all";
                                                                    break;
                                                            }


                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                encText.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            encText.type = "EncryptedText";

                                                            custObj.fields.Add((Object)encText);
                                                            break;

                                                        case "text":
                                                            Text tx = new Text();
                                                            tx.caseSensitive = false;
                                                            tx.defaultValue = "";
                                                            tx.description = "";
                                                            tx.externalId = false;
                                                            tx.fullName = cols[4];
                                                            tx.inlineHelpText = "";
                                                            tx.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                            //Para recuperar o valor da precisão dos campos
                                                            lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                            lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                            tx.length = Convert.ToInt32(lenStr);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                tx.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            tx.unique = false;
                                                            tx.type = "Text";

                                                            custObj.fields.Add((Object)tx);
                                                            break;

                                                        case "url":
                                                            URL url = new URL();
                                                            url.defaultValue = "";
                                                            url.description = "";
                                                            url.fullName = cols[4];
                                                            url.inlineHelpText = "";
                                                            url.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                            if (!string.IsNullOrEmpty(cols[9]))
                                                            {
                                                                url.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                            }
                                                            url.type = "Url";

                                                            custObj.fields.Add((Object)url);
                                                            break;

                                                        default:
                                                            //Caso seja diferente de qualquer um dos valores acima, lançaremos o erro no log
                                                            errosProcesso.Add("[" + DateTime.Now + "] Tipo de campo não identificado - " + tipoCampo + "");
                                                            break;
                                                    }
                                                }
                                                else
                                                {
                                                    errosProcesso.Add("[" + DateTime.Now + "] Tipo de campo não informado. Campo - " + cols[2]);
                                                }

                                                #endregion 
                                            }
                                        }
                                        countLinha++;
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        //Caso contrario, passaremos a criação dele no objeto para depois gerarmos o xml
                        #region Novo Objeto

                        //Objeto que irá conter as informações que necessitamos para criar o xml
                        CustomObject custom = new CustomObject();

                        //Criando o array dos objetos de campos
                        custom.fields = new List<Object>();

                        foreach (string[] cols in Linhas)
                        {
                            if (countLinha >= 1 && countLinha <= 5)
                            {
                                //Pulando linhas desnecessárias
                                countLinha++;
                                continue;
                            }

                            if (countLinha == 0)
                            {
                                //Adicionando o tipo de objeto na lista de objetos processados...
                                objetosProcessados.Add(cols[1]);
                            }

                            //A condição de saida do foreach será o valor Lista Relacionada do documento ou a leitura de espaço vazio
                            if (string.IsNullOrEmpty(cols[0]) || cols[0].Trim().Equals("Lista Relacionada"))
                            {
                                break;
                            }

                            if (!cols[0].Trim().Equals("Section:"))
                            {
                                nomeCol = cols[2];
                                nomeSheet = objSheet.Name;

                                //Validaremos se o campo do API Name está preenchido
                                if (!string.IsNullOrEmpty(cols[4]))
                                {
                                    //Se a label pt-br estiver vazia, o nome do campo será o mesmo da API name


                                    //Validaremos se o FieldType está preenchido

                                    if (!string.IsNullOrEmpty(cols[5]))
                                    {
                                        //Primeiro, determinamos qual o tipo do campo

                                        string tipoCampo = cols[5];

                                        //Inicialmente, iremos tratar o campo para tirar parenteses e valores numericos...

                                        tipoCampo = tipoCampo.Replace("(", "").Replace(")", "").Replace("0", "").Replace("1", "").Replace("2", "").Replace("3", "").Replace("4", "").Replace("5", "");
                                        tipoCampo = tipoCampo.Replace("6", "").Replace("7", "").Replace("8", "").Replace("9", "").Replace(".", "").Replace(" ", "").Replace(",", ".").ToLower();

                                        //Agora, criaremos um switch / case para criar o objeto correto de acordo com o tipo do campo

                                        switch (tipoCampo)
                                        {

                                            case "email":
                                                Email email = new Email();
                                                email.defaultValue = "";
                                                email.externalId = false;
                                                email.fullName = cols[4];
                                                email.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    email.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                email.type = "Email";
                                                email.unique = false;

                                                custom.fields.Add((Object)email);
                                                break;

                                            case "boolean":
                                                Checkbox chk = new Checkbox();
                                                chk.defaultValue = (cols[5].Contains("1") ? true : false);
                                                chk.description = "";
                                                chk.fullName = cols[4];
                                                chk.inlineHelpText = "";
                                                chk.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                chk.type = "Checkbox";

                                                custom.fields.Add((Object)chk);
                                                break;

                                            case "datetime":
                                                DateTimeSF dtSF = new DateTimeSF();
                                                dtSF.defaultvalue = "";
                                                dtSF.description = "";
                                                dtSF.fullName = cols[4];
                                                dtSF.inlineHelpText = "";
                                                dtSF.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    dtSF.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                dtSF.type = "DateTime";

                                                custom.fields.Add((Object)dtSF);
                                                break;

                                            case "date":
                                                DateSF dSF = new DateSF();
                                                dSF.defaultValue = "";
                                                dSF.description = "";
                                                dSF.fullName = cols[4];
                                                dSF.inlineHelpText = "";
                                                dSF.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    dSF.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                dSF.type = "Date";

                                                custom.fields.Add((Object)dSF);
                                                break;

                                            case "location":
                                                Location lct = new Location();
                                                lct.description = "";
                                                lct.fullName = cols[4];
                                                lct.inlineHelpText = "";
                                                lct.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    lct.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                lct.scale = 3; //Por default, eu vou deixar valor 3. alterar depois se vier um valor descrito no excel...
                                                lct.type = "Location";

                                                custom.fields.Add((Object)lct);
                                                break;

                                            case "currency":
                                                Currency curr = new Currency();
                                                curr.defaultValue = "";
                                                curr.description = "";
                                                curr.fullName = cols[4];
                                                curr.inlineHelpText = "";
                                                curr.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                //Para recuperar o valor da precisão dos campos
                                                string lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                if (lenStr.Contains('.'))
                                                {
                                                    curr.precision = Convert.ToInt32(lenStr.Split('.')[0]) + Convert.ToInt32(lenStr.Split('.')[1]);
                                                    curr.scale = Convert.ToInt32(lenStr.Split('.')[1]);
                                                }
                                                else
                                                {
                                                    curr.precision = Convert.ToInt32(lenStr.Split('.')[0]);
                                                    curr.scale = 0;
                                                }

                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    curr.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                curr.type = "Currency";

                                                custom.fields.Add((Object)curr);
                                                break;

                                            case "autonumber":
                                                AutoNumber auto = new AutoNumber();
                                                auto.description = "";
                                                auto.displayFormat = "";
                                                auto.externalId = false; //Por default, eu vou deixar os campos de externalId como false. alterar depois se vier um valor descrito no excel...
                                                auto.fullName = cols[4];
                                                auto.inlineHelpText = "";
                                                auto.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                auto.type = "AutoNumber";

                                                custom.fields.Add((Object)auto);
                                                break;

                                            case "number":
                                                Number nb = new Number();
                                                nb.defaultValue = "";
                                                nb.description = "";
                                                nb.externalId = false;
                                                nb.fullName = cols[4];
                                                nb.inlineHelpText = "";
                                                nb.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                //Para recuperar o valor da precisão dos campos
                                                lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                if (lenStr.Contains('.'))
                                                {
                                                    nb.precision = Convert.ToInt32(lenStr.Split('.')[0]) + Convert.ToInt32(lenStr.Split('.')[1]);
                                                    nb.scale = Convert.ToInt32(lenStr.Split('.')[1]);
                                                }
                                                else
                                                {
                                                    nb.precision = Convert.ToInt32(lenStr.Split('.')[0]);
                                                    nb.scale = 0;
                                                }


                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    nb.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                nb.type = "Number";
                                                nb.unique = false;

                                                custom.fields.Add((Object)nb);
                                                break;

                                            case "multiselectpicklist":
                                                MultiselectPicklist multiPick = new MultiselectPicklist();
                                                multiPick.description = "";
                                                multiPick.fullName = cols[4];
                                                multiPick.inlineHelpText = "";
                                                multiPick.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                multiPick.sorted = false;
                                                multiPick.visibleLines = 4;
                                                multiPick.picklist = new List<PickListValue>();

                                                if (!string.IsNullOrEmpty(cols[7]))
                                                {
                                                    string[] valoresPickListSingle = cols[7].Split('\n');

                                                    foreach (string valor in valoresPickListSingle)
                                                    {
                                                        PickListValue pickValue = new PickListValue();
                                                        pickValue.defaultValue = false;
                                                        pickValue.fullname = valor.TrimEnd(' ');

                                                        multiPick.picklist.Add(pickValue);
                                                    }

                                                    multiPick.type = "MultiselectPicklist";

                                                    custom.fields.Add((Object)multiPick);
                                                }
                                                else
                                                {
                                                    errosProcesso.Add("[" + DateTime.Now + "] Campo " + (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]) + " não será criado pois não possui items");
                                                }


                                                break;

                                            case "picklist":
                                                Picklist pick = new Picklist();
                                                pick.description = "";
                                                pick.fullName = cols[4];
                                                pick.inlineHelpText = "";
                                                pick.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                pick.picklist = new List<PickListValue>();

                                                if (!string.IsNullOrEmpty(cols[7]))
                                                {
                                                    string[] valoresPickListSingle = cols[7].Split('\n');

                                                    foreach (string valor in valoresPickListSingle)
                                                    {
                                                        PickListValue pickValue = new PickListValue();
                                                        pickValue.defaultValue = false;
                                                        pickValue.fullname = valor.TrimEnd(' ');

                                                        pick.picklist.Add(pickValue);
                                                    }

                                                    pick.sorted = false;
                                                    pick.type = "Picklist";

                                                    custom.fields.Add((Object)pick);
                                                }
                                                else
                                                {
                                                    errosProcesso.Add("[" + DateTime.Now + "] Campo " + (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]) + " não será criado pois não possui items");
                                                }


                                                break;

                                            case "percent":
                                                Percent prct = new Percent();
                                                prct.defaultValue = "";
                                                prct.description = "";
                                                prct.fullname = cols[4];
                                                prct.inlineHelpText = "";
                                                prct.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                //Para recuperar o valor da precisão dos campos
                                                lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                lenStr = lenStr.Replace("(", "").Replace(")", "");


                                                if (lenStr.Contains('.'))
                                                {
                                                    prct.precision = Convert.ToInt32(lenStr.Split('.')[0]) + Convert.ToInt32(lenStr.Split('.')[1]);
                                                    prct.scale = Convert.ToInt32(lenStr.Split('.')[1]);
                                                }
                                                else
                                                {
                                                    prct.precision = Convert.ToInt32(lenStr.Split('.')[0]);
                                                    prct.scale = 0;
                                                }

                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    prct.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                prct.type = "Percent";

                                                custom.fields.Add((Object)prct);
                                                break;

                                            case "html":
                                                HTML html = new HTML();
                                                html.description = "";
                                                html.fullName = cols[4];
                                                html.inlineHelpText = "";
                                                html.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                //Para recuperar o valor da precisão dos campos
                                                lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                html.length = Convert.ToInt32(lenStr); // Valor padrão do SF
                                                html.visibleLines = 25; //Valor padrão do SF;
                                                html.type = "Html";

                                                custom.fields.Add((Object)html);
                                                break;

                                            case "phone":
                                                Phone phn = new Phone();
                                                phn.defaultValue = "";
                                                phn.description = "";
                                                phn.fullName = cols[4];
                                                phn.inlineHelpText = "";
                                                phn.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    phn.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                phn.type = "Phone";

                                                custom.fields.Add((Object)phn);
                                                break;

                                            case "longtextarea":
                                                LongTextArea lta = new LongTextArea();
                                                lta.defaultValue = "";
                                                lta.description = "";
                                                lta.fullName = cols[4];
                                                lta.inlineHelpText = "";
                                                lta.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                //Para recuperar o valor da precisão dos campos
                                                lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                lta.length = Convert.ToInt32(lenStr);
                                                lta.visibleLines = 3;
                                                lta.type = "LongTextArea";

                                                custom.fields.Add((Object)lta);
                                                break;

                                            case "textarea":
                                                TextArea txAr = new TextArea();
                                                txAr.defaultValue = "";
                                                txAr.description = "";
                                                txAr.fullName = cols[4];
                                                txAr.inlineHelpText = "";
                                                txAr.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    txAr.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                txAr.type = "TextArea";

                                                custom.fields.Add((Object)txAr);
                                                break;

                                            case "encryptedtext":
                                                EncryptedText encText = new EncryptedText();
                                                encText.fullName = cols[4];
                                                encText.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                //Para recuperar o valor da precisão dos campos
                                                lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                encText.length = Convert.ToInt32(lenStr);

                                                //TODO: Definir qual campo irá conter o tipo de criptografia e qual campo irá conter o caractere de criptografia
                                                string caractereCripto = "";
                                                string tipoCripto = "";

                                                //O SF aceita apenas valores mascarados por 'X' ou '*'
                                                if (caractereCripto.Equals("X"))
                                                {
                                                    encText.maskChar = "X";
                                                }
                                                else
                                                {
                                                    encText.maskChar = "asterisk";
                                                }

                                                /*   O SF possui os seguintes tipos de mascara:
                                                 *  Mascarar todos os caracteres - all
                                                 *  Limpar quatro últimos caracteres - lastFour
                                                 *  Número do cartão de crédito - creditCard
                                                 *  Número do seguro nacional - nino
                                                 *  Número do CPF - ssn
                                                 *  Número do seguro social - sin
                                                 */

                                                switch (tipoCripto)
                                                {
                                                    case "Quatro ultimos":
                                                        encText.maskType = "lastFour";
                                                        break;
                                                    case "Cartão de crédito":
                                                        encText.maskType = "creditCard";
                                                        break;
                                                    case "Seguro nacional":
                                                        encText.maskType = "nino";
                                                        break;
                                                    case "CPF":
                                                        encText.maskType = "ssn";
                                                        break;
                                                    case "Seguro social":
                                                        encText.maskType = "sin";
                                                        break;
                                                    default:
                                                        //Por padrão, se não for informado um tipo valido, iremos considerar a criptografia total da string
                                                        encText.maskType = "all";
                                                        break;
                                                }


                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    encText.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                encText.type = "EncryptedText";

                                                custom.fields.Add((Object)encText);
                                                break;

                                            case "text":
                                                Text tx = new Text();
                                                tx.caseSensitive = false;
                                                tx.defaultValue = "";
                                                tx.description = "";
                                                tx.externalId = false;
                                                tx.fullName = cols[4];
                                                tx.inlineHelpText = "";
                                                tx.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);

                                                //Para recuperar o valor da precisão dos campos
                                                lenStr = cols[5].Substring(cols[5].IndexOf("("));
                                                lenStr = lenStr.Replace("(", "").Replace(")", "");

                                                tx.length = Convert.ToInt32(lenStr);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    tx.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                tx.unique = false;
                                                tx.type = "Text";

                                                custom.fields.Add((Object)tx);
                                                break;

                                            case "url":
                                                URL url = new URL();
                                                url.defaultValue = "";
                                                url.description = "";
                                                url.fullName = cols[4];
                                                url.inlineHelpText = "";
                                                url.label = (string.IsNullOrEmpty(cols[2]) ? cols[4] : cols[2]);
                                                if (!string.IsNullOrEmpty(cols[9]))
                                                {
                                                    url.required = (cols[9].Trim().Equals("Mandatory") ? true : false);
                                                }
                                                url.type = "Url";

                                                custom.fields.Add((Object)url);
                                                break;

                                            default:
                                                //Caso seja diferente de qualquer um dos valores acima, lancaremos um erro no log
                                                errosProcesso.Add("[" + DateTime.Now + "] Tipo de campo não identificado - " + tipoCampo + "");
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        //Se estiver vazio, lançaremos o erro no log...
                                        errosProcesso.Add("[" + DateTime.Now + "] Tipo de campo não informado. Campo - " + cols[2]);
                                    }
                                }
                            }
                            countLinha++;
                        }

                        //Para determinar quais são objetos customizados e os objetos padroes, iremos criar uma lista com os objetos mais comuns. Se o nome do objeto constar na lista,
                        //será considerada somente a adição dos fields
                        List<string> objetosPadraoSF = new List<string>();

                        objetosPadraoSF.Add("Account");
                        objetosPadraoSF.Add("Attachment");
                        objetosPadraoSF.Add("Case");
                        objetosPadraoSF.Add("Contact");
                        objetosPadraoSF.Add("Campaign");
                        objetosPadraoSF.Add("CampaignMember");
                        objetosPadraoSF.Add("Lead");
                        objetosPadraoSF.Add("Opportunity");
                        objetosPadraoSF.Add("OpportunityLineItem");
                        objetosPadraoSF.Add("Product2");
                        objetosPadraoSF.Add("Task");
                        objetosPadraoSF.Add("User");

                        if (!objetosPadraoSF.Contains(objetosProcessados[objetosProcessados.Count - 1]))
                        {
                            //Uma vez os objetos preenchidos, iremos colocar as informações sobre o novo objeto do SF antes de adiciona-lo à lista
                            custom.deploymentStatus = "Deployed";

                            //custom.gender = (objetosProcessados[objetosProcessados.Count - 1].Substring(objetosProcessados.Count - 1, 1).Equals("a") ? "Feminine" : "Masculine");

                            custom.pluralLabel = objetosProcessados[objetosProcessados.Count - 1] + "s";
                            custom.searchLayouts = null;
                            custom.sharingModel = "ReadWrite";
                            custom.nameField = new NewField();
                            custom.nameField.displayFormat = "{0}";
                            custom.nameField.label = objetosProcessados[objetosProcessados.Count - 1] + "Id";
                            custom.nameField.type = "AutoNumber";
                        }

                        //Caso seja um objeto padrão, prosseguiremos à adição dele na lista de objetos sem as informações adicionais

                        //OBS.: Para efeito de identificação, estaremos guardando apenas a Label do objeto...
                        custom.label = objetosProcessados[objetosProcessados.Count - 1];

                        //Adicionando o elemento na lista
                        listaCustomObj.Add(custom);

                        #endregion
                    }
                }
                catch (Exception e)
                {
                    errosProcesso.Add("[" + DateTime.Now + "] - Erro ao processar o arquivo do excel. Exception: " + e.Message + ". Campo: " + nomeCol + " Nome Sheet: " + nomeSheet + " StackTrace: " + e.StackTrace);
                }
            }

            //Apartir daqui, iremos criar o xml para criarmos o objeto via Eclipse
            criaXmlSF(listaCustomObj);

            objWorkbook.Close();

            //Validaremos se houve algum erro durante o processamento...
            if (errosProcesso.Count == 0)
            {
                //Se não houver, retornaremos true
                return true;
            }
            else
            {
                //Caso contrario, criaremos um arquivo com todos os erros descritos durante o processaemnto do documento xls

                //variavel com o local para a gravação e o nome do arquivo
                string fileName = @"C:\Temp\LogErrosSFXML\log.txt";

                try
                {
                    //Verificamos se o arquivo já existe
                    if (File.Exists(fileName))
                    {
                        //apagamos ele do sistema
                        File.Delete(fileName);
                    }

                    //Criamos o novo arquivo

                    using (StreamWriter sw = File.CreateText(fileName))
                    {
                        //Aqui, percorreremos a lista de string para criar o arquivo

                        foreach (string str in errosProcesso)
                        {
                            sw.WriteLine(str);
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("Erro ao criar o arquivo com o log de erros. Exception : " + e.Message);
                }

                //retornaremos false no final da criação...

                return false;
            }
        }

        /// <summary>
        /// Método que cria a definição do xml de acordo com o objeto recuperado na sheet
        /// </summary>
        /// <param name="objetoCustom">Lista de objetos customizados para criar o XML</param>
        private void criaXmlSF(List<CustomObject> objetoCustom)
        {

            List<string> linhasArquivo = new List<string>();
            //Percorreremos o objeto para criar os xmls respectivos de cada Sheet que o documento possuia
            foreach (CustomObject custObj in objetoCustom)
            {
                try
                {
                    //Variavel que vai conter o XML
                    string xmlObjSF = string.Empty;

                    //De antemão, atribuiremos a tag padrão do xml para a variavel juntamente com a tag base do CustomObject
                    xmlObjSF = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
                    xmlObjSF += "<CustomObject xmlns=\"http://soap.sforce.com/2006/04/metadata\">";

                    //Primeiramente, iremos diferir se o xml da iteração será um xml de um objeto padrão ou customizado do Salesforce

                    /*Validaremos pela propriedade 'deploymentStatus' pois ela só é presente em objetos customizados. Se ela for nula ou vazia então seguiremos com o objeto padrão.
                     Caso contrário, criaremos um XML de objeto customizado*/
                    if (string.IsNullOrEmpty(custObj.deploymentStatus))
                    {
                        #region Objeto Padrão

                        //Percorreremos a lista dos tipos com os valores dos campos para criar a tag 'fields'

                        foreach (Object obj in custObj.fields)
                        {

                            //Antes de mais nada, atribuiremos as tags de 'fields' no começo da iteração...
                            xmlObjSF += "<fields>";

                            //Primeiramente, iremos determinar com qual objeto estamos tratando...
                            Type t = obj.GetType();

                            PropertyInfo propSingle = t.GetProperty("type");

                            string tipoPropriedade = (string)propSingle.GetValue(obj, null);

                            //Agora daremos o cast para string...

                            //Como apenas os objetos Picklist e MultiselectPicklist tem objetos dentro deles alem dos tipos primitivos, iremos tratar os dois como exceção...
                            if (tipoPropriedade.Equals("Picklist"))
                            {
                                #region Picklist

                                //Object para Picklist
                                Picklist pk = (Picklist)obj;

                                string fullNameVal = pk.fullName.Replace(" ", "");
                                fullNameVal = fullNameVal + "__c";

                                //Tag FullName
                                xmlObjSF += "<fullName>" + fullNameVal + "</fullName>";

                                //Tag Description
                                xmlObjSF += "<description>" + pk.description + "</description>";

                                //Tag inlineHelpText
                                xmlObjSF += "<inlineHelpText>" + pk.inlineHelpText + "</inlineHelpText>";

                                //Tag Label
                                xmlObjSF += "<label>" + pk.label + "</label>";

                                //Aqui iniciaremos a criação das tags dos picklistvalues

                                //Tag de abertura...
                                xmlObjSF += "<picklist>";

                                foreach (PickListValue pkValue in pk.picklist)
                                {
                                    //Tag de abertura do objeto
                                    xmlObjSF += "<picklistValues>";

                                    //Valores...
                                    xmlObjSF += "<fullName>" + pkValue.fullname + "</fullName>";
                                    xmlObjSF += "<default>" + pkValue.defaultValue + "</default>";

                                    //Tag de fechamento do objeto
                                    xmlObjSF += "</picklistValues>";
                                }

                                //Tag de Sorted
                                xmlObjSF += "<sorted>" + pk.sorted + "</sorted>";

                                //...tag de fechamento
                                xmlObjSF += "</picklist>";

                                xmlObjSF += "<type>" + pk.type.ToString() + "</type>";

                                #endregion
                            }
                            else if (tipoPropriedade.Equals("MultiselectPicklist"))
                            {
                                #region MultiselectPicklist

                                //Object para MultiselectPicklist
                                MultiselectPicklist pk = (MultiselectPicklist)obj;

                                string fullNameVal = pk.fullName.Replace(" ", "");
                                fullNameVal = fullNameVal + "__c";

                                //Tag FullName
                                xmlObjSF += "<fullName>" + fullNameVal + "</fullName>";

                                //Tag Description
                                xmlObjSF += "<description>" + pk.description + "</description>";

                                //Tag inlineHelpText
                                xmlObjSF += "<inlineHelpText>" + pk.inlineHelpText + "</inlineHelpText>";

                                //Tag Label
                                xmlObjSF += "<label>" + pk.label + "</label>";

                                //Aqui iniciaremos a criação das tags dos picklistvalues

                                //Tag de abertura...
                                xmlObjSF += "<picklist>";

                                foreach (PickListValue pkValue in pk.picklist)
                                {
                                    //Tag de abertura do objeto
                                    xmlObjSF += "<picklistValues>";

                                    //Valores...
                                    xmlObjSF += "<fullName>" + pkValue.fullname + "</fullName>";
                                    xmlObjSF += "<default>" + pkValue.defaultValue + "</default>";

                                    //Tag de fechamento do objeto
                                    xmlObjSF += "</picklistValues>";
                                }

                                //Tag de Sorted
                                xmlObjSF += "<sorted>" + pk.sorted + "</sorted>";

                                //...tag de fechamento
                                xmlObjSF += "</picklist>";

                                xmlObjSF += "<type>" + pk.type.ToString() + "</type>";

                                //Tag de visibleLines
                                xmlObjSF += "<visibleLines>" + pk.visibleLines + "</visibleLines>";

                                #endregion
                            }
                            else
                            {
                                #region Outros objetos

                                //...todos os outros seguirão uma mesma forma de construção

                                //Declaramos um array de propriedades que receberá os dados do Type
                                PropertyInfo[] propriedades = t.GetProperties();

                                //Percorreremos a lista de propriedades e criaremos o xml referente a cada objeto
                                foreach (PropertyInfo propFields in propriedades)
                                {
                                    //Para não pegarmos uma propriedade do sistema...
                                    if (propFields.CanRead)
                                    {
                                        if (propFields.CanWrite)
                                        {
                                            //Todas os valores do objeto entenderemos como string para montar o XML
                                            var valorCampo = propFields.GetValue(obj, null);


                                            if (propFields.Name.Equals("fullName"))
                                            {
                                                //Caso seja a tag FullName - para concatenar o __c
                                                string valCam = valorCampo.ToString().Replace(" ", "");
                                                valCam = valCam + "__c";

                                                //Montando a tag dinamicamente...
                                                xmlObjSF += "<" + propFields.Name + ">" + valCam + "</" + propFields.Name + ">";
                                            }
                                            else
                                            {
                                                //Montando a tag dinamicamente...
                                                xmlObjSF += "<" + propFields.Name + ">" + valorCampo + "</" + propFields.Name + ">";
                                            }
                                        }
                                    }
                                }

                                #endregion
                            }

                            //...e no final da iteração...
                            xmlObjSF += "</fields>";
                        }

                        #endregion
                    }
                    else
                    {
                        #region Objeto Customizado

                        //Antes de criarmos as tags fields, criamos a tag do deploymentStatus
                        xmlObjSF += "<deploymentStatus>" + custObj.deploymentStatus + "</deploymentStatus>";

                        //Percorreremos a lista dos tipos com os valores dos campos para criar a tag 'fields'

                        foreach (Object obj in custObj.fields)
                        {
                            //Antes de mais nada, atribuiremos as tags de 'fields' no começo da iteração...
                            xmlObjSF += "<fields>";

                            //Primeiramente, iremos determinar com qual objeto estamos tratando...

                            Type t = obj.GetType();

                            PropertyInfo propSingle = t.GetProperty("type");

                            string tipoPropriedade = (string)propSingle.GetValue(obj, null);

                            //Como apenas os objetos Picklist e MultiselectPicklist tem objetos dentro deles alem dos tipos primitivos, iremos tratar os dois como exceção...

                            if (tipoPropriedade.Equals("Picklist"))
                            {
                                #region Picklist

                                //Object para Picklist
                                Picklist pk = (Picklist)obj;

                                string pkFullName = pk.fullName.Replace(" ", "");
                                pkFullName = pkFullName + "__c";

                                //Tag FullName
                                xmlObjSF += "<fullName>" + pkFullName + "</fullName>";

                                //Tag Description
                                xmlObjSF += "<description>" + pk.description + "</description>";

                                //Tag inlineHelpText
                                xmlObjSF += "<inlineHelpText>" + pk.inlineHelpText + "</inlineHelpText>";

                                //Tag Label
                                xmlObjSF += "<label>" + pk.label + "</label>";

                                //Aqui iniciaremos a criação das tags dos picklistvalues

                                //Tag de abertura...
                                xmlObjSF += "<picklist>";

                                foreach (PickListValue pkValue in pk.picklist)
                                {
                                    //Tag de abertura do objeto
                                    xmlObjSF += "<picklistValues>";

                                    //Valores...
                                    xmlObjSF += "<fullName>" + pkValue.fullname + "</fullName>";
                                    xmlObjSF += "<default>" + pkValue.defaultValue + "</default>";

                                    //Tag de fechamento do objeto
                                    xmlObjSF += "</picklistValues>";
                                }

                                //Tag de Sorted
                                xmlObjSF += "<sorted>" + pk.sorted + "</sorted>";

                                //...tag de fechamento
                                xmlObjSF += "</picklist>";

                                xmlObjSF += "<type>" + pk.type.ToString() + "</type>";

                                #endregion
                            }
                            else if (tipoPropriedade.Equals("MultiselectPicklist"))
                            {
                                #region MultiselectPicklist

                                //Object para Picklist
                                MultiselectPicklist pk = (MultiselectPicklist)obj;

                                string pkFullName = pk.fullName.Replace(" ", "");
                                pkFullName = pkFullName + "__c";

                                //Tag FullName
                                xmlObjSF += "<fullName>" + pkFullName + "</fullName>";

                                //Tag Description
                                xmlObjSF += "<description>" + pk.description + "</description>";

                                //Tag inlineHelpText
                                xmlObjSF += "<inlineHelpText>" + pk.inlineHelpText + "</inlineHelpText>";

                                //Tag Label
                                xmlObjSF += "<label>" + pk.label + "</label>";

                                //Aqui iniciaremos a criação das tags dos picklistvalues

                                //Tag de abertura...
                                xmlObjSF += "<picklist>";

                                foreach (PickListValue pkValue in pk.picklist)
                                {
                                    //Tag de abertura do objeto
                                    xmlObjSF += "<picklistValues>";

                                    //Valores...
                                    xmlObjSF += "<fullName>" + pkValue.fullname + "</fullName>";
                                    xmlObjSF += "<default>" + pkValue.defaultValue + "</default>";

                                    //Tag de fechamento do objeto
                                    xmlObjSF += "</picklistValues>";
                                }

                                //Tag de Sorted
                                xmlObjSF += "<sorted>" + pk.sorted + "</sorted>";

                                //...tag de fechamento
                                xmlObjSF += "</picklist>";

                                xmlObjSF += "<type>" + pk.type.ToString() + "</type>";

                                //Tag de visibleLines
                                xmlObjSF += "<visibleLines>" + pk.visibleLines + "</visibleLines>";

                                #endregion
                            }
                            else
                            {
                                #region Outros objetos

                                //...todos os outros seguirão uma mesma forma de construção

                                //Declaramos um array de propriedades que receberá os dados do Type
                                PropertyInfo[] propriedades = t.GetProperties();

                                //Percorreremos a lista de propriedades e criaremos o xml referente a cada objeto
                                foreach (PropertyInfo propFields in propriedades)
                                {
                                    //Para não pegarmos uma propriedade do sistema...
                                    if (propFields.CanRead)
                                    {
                                        if (propFields.CanWrite)
                                        {
                                            //Todas os valores do objeto entenderemos como string para montar o XML
                                            var valorCampo = propFields.GetValue(obj, null);

                                            if (propFields.Name.Equals("fullName"))
                                            {
                                                //Caso seja a tag FullName - para concatenar o __c
                                                string valCam = valorCampo.ToString().Replace(" ", "");
                                                valCam = valCam + "__c";

                                                //Montando a tag dinamicamente...
                                                xmlObjSF += "<" + propFields.Name + ">" + valCam + "</" + propFields.Name + ">";
                                            }
                                            else
                                            {
                                                //Montando a tag dinamicamente...
                                                xmlObjSF += "<" + propFields.Name + ">" + valorCampo + "</" + propFields.Name + ">";
                                            }
                                        }
                                    }
                                }

                                #endregion
                            }

                            //...e no final da iteração...
                            xmlObjSF += "</fields>";
                        }

                        #endregion
                    }

                    //Fechando o XML
                    xmlObjSF += "</CustomObject>";

                    //Aqui, iremos processar a data/hora que o arquivo foi criado, o nome do objeto e o conteudo XML do mesmo

                    /*A criação das linhas seguirá a seguinte ordem:
                     *Linha 1 - Data/Hora====================NomeDoObjeto=====================
                     *Linha 2 - Conteudo XML
                     *Linha 3 - ==============================================================
                     *Linha 4 - Espaço vazio
                     * */

                    //Primeira linha - data/hora, nome do objeto e separador
                    string linha1 = DateTime.Now.ToString() + "===============================================" + custObj.label + "===================================";
                    linhasArquivo.Add(linha1);

                    //Segunda linha - XML
                    linhasArquivo.Add(xmlObjSF);

                    //Terceira linha - Separador
                    linhasArquivo.Add("============================================================================================");

                    //Quarta linha - Espaço vazio
                    linhasArquivo.Add(string.Empty);
                }
                catch (Exception e)
                {
                    errosProcesso.Add("[" + DateTime.Now + "] Erro ao criar o XML dos objetos. Exception: " + e.Message);
                }
            }

            criaTxt(linhasArquivo);
        }

        /// <summary>
        /// Método que cria o arquivo TXT com os XMLS gerados no processo
        /// </summary>
        /// <param name="linhasArquivo"></param>
        private void criaTxt(List<string> linhasArquivo)
        {
            //variavel com o local para a gravação e o nome do arquivo
            string fileName = @"C:\Temp\GeradorObjetosSF.txt";

            try
            {
                //Verificamos se o arquivo já existe
                if (File.Exists(fileName))
                {
                    //apagamos ele do sistema
                    File.Delete(fileName);
                }

                //Criamos o novo arquivo

                using (StreamWriter sw = File.CreateText(fileName))
                {
                    //Aqui, percorreremos a lista de string para criar o arquivo

                    foreach (string str in linhasArquivo)
                    {
                        sw.WriteLine(str);
                    }
                }
            }
            catch (Exception e)
            {
                errosProcesso.Add("[" + DateTime.Now + "] Erro ao criar o arquivo XML. Exception: " + e.Message);
            }
        }

    }
}

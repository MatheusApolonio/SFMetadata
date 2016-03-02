using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SFMetadata
{
    public partial class SFMetadata : Form
    {
        public SFMetadata()
        {
            InitializeComponent();
        }

        private void SFMetadata_Load(object sender, EventArgs e)
        {

        }

        private void btnProcessar_Click(object sender, EventArgs e)
        {
            Utilidades util = new Utilidades();
            Utilidades.AsyncMethodCaller caller = new Utilidades.AsyncMethodCaller(util.LeArquivo);

            IAsyncResult result = caller.BeginInvoke(txtCaminho.Text, null, null);

            LoadingForm loadFrm = new LoadingForm();

            bool execucaoOK = true;
            loadFrm.Show();
            execucaoOK = caller.EndInvoke(result);

            loadFrm.Hide();

            if (!execucaoOK)
            {
                string pathFile = @"C:\Temp\LogErrosSFXML";
                MessageBox.Show("Foram encontrados erros durante o processaemento do arquivo. Favor verificar o arquivo log.txt no seguinte caminho: " + pathFile + "");

            }
        }

        private void btnSelecionar_Click(object sender, EventArgs e)
        {
            {
                //define as propriedades do controle 
                this.openFileDialog1.Multiselect = false;
                this.openFileDialog1.Title = "Selecionar Excel";
                openFileDialog1.InitialDirectory = @"C:\";
                openFileDialog1.Filter = "Microsoft Excel (*.XLS;*.XLSX;*.xls;*.xlsx) | *.XLS;*.XLSX;*.xls;*.xlsx";
                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = false;
                openFileDialog1.ReadOnlyChecked = true;
                openFileDialog1.ShowReadOnly = true;
                this.openFileDialog1.ShowDialog();

                string strPath = null;
                strPath = this.openFileDialog1.FileName;

                if ((strPath != null && strPath != ""))
                    if (File.Exists(strPath))
                        txtCaminho.Text = strPath;
                    else
                        MessageBox.Show("Arquivo não encontrado!");
                else
                    MessageBox.Show("Nenhum arquivo selecionado!");
            }
        }
    }
}

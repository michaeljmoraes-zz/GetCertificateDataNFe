using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


/// <summary>
/// CÓDIGO DESENVOLVIDO POR MICHAEL MORAES PARA FACILITAR EXTRAÇÃO DE DADOS DE 
/// CERTIFICADOS NFe, a Distribuição desse utilitário e livre e pode ser usado, 
/// Copiada a vontade e usado como quiser
/// 
/// Mas se quiser ajuda pode me contatar
/// Tel: +55 11 97956-3231
/// michaeljmoraes@gmail.com
/// fb.com/michaeljmoraes
/// </summary>
namespace GetCertificateData
{
    /// <summary>
    /// Interação lógica para MainWindow.xam
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnCarregarCertificados_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var cert = ListareObterDoRepositorio();
                if (!cert.Subject.Split(',')[2].Split('=')[1].Contains("RFB"))
                    throw new Exception("Esse certificado é do tipo: " + cert.Subject.Split(',')[2].Split('=')[1] + "\nPara NFe precisa ser RFB");


                this.txtSerialNumber.Text = cert.SerialNumber;
                this.txtCNPJ.Text = cert.Subject.Split(',')[0].Split(':')[1];
                this.txtRazaoSocial.Text = cert.Subject.Split(',')[0].Split(':')[0].Split('=')[1];

                this.lblDescUF.Content = cert.Subject.Split(',')[6].Split('=')[1];
                String codigoUF  = Convert.ToString((Int32)Enum.Parse(typeof(enmUF), cert.Subject.Split(',')[6].Split('=')[1]));

                this.txtUF.Text = codigoUF;

                this.txtCidade.Text = cert.Subject.Split(',')[5].Split('=')[1];
                this.txtTipoCertificado.Text = cert.Subject.Split(',')[2].Split('=')[1];

            }catch(Exception ex)
            {
                MessageBox.Show("Não é um certificado para NFe Válido\n\n" + ex.Message);
                
            }

        }


        #region Certificado Digital
        /// <summary>
        /// Exibe a lista de certificados instalados no PC e devolve o certificado selecionado
        /// </summary>
        /// <returns></returns>
        public X509Certificate2 ListareObterDoRepositorio()
        {
            var store = ObterX509Store(OpenFlags.OpenExistingOnly | OpenFlags.ReadOnly);
            var collection = store.Certificates;
            var fcollection = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, true);
            var scollection = X509Certificate2UI.SelectFromCollection(fcollection, "Certificados válidos:", "Selecione o certificado que deseja usar",
                X509SelectionFlag.SingleSelection);

            if (scollection.Count == 0)
            {
                throw new Exception("Nenhum certificado foi selecionado!");
            }

            store.Close();
            return scollection[0];
        }

        /// <summary>
        /// Cria e devolve um objeto <see cref="X509Store"/>
        /// </summary>
        /// <param name="openFlags"></param>
        /// <returns></returns>
        private X509Store ObterX509Store(OpenFlags openFlags)
        {
            var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(openFlags);
            return store;
        }




        public enum enmUF
        {
            NA = 0,
            RO = 11, // Rondônia 
            AC = 12, // Acre 
            AM = 13, // Amazonas 
            RR = 14, // Roraima 
            PA = 15, // Pará 
            AP = 16, // Amapá 
            TO = 17, // Tocantins 
            MA = 21, // Manaus 
            PI = 22, // Piauí 
            CE = 23, // Ceará 
            RN = 24, // Rio Grande do Norte 
            PB = 25, // Paraíba 
            PE = 26, // Pernambuco 
            AL = 27, // Alagoas 
            SE = 28, // Sergipe 
            BA = 29, // Bahia 
            MG = 31, // Minas Gerais 
            ES = 32, // Espírito Santo 
            RJ = 33, // Rio de Janeiro 
            SP = 35, // São Paulo 
            PR = 41, // Paraná 
            SC = 42, // Santa Catarina 
            RS = 43, // Rio Grande do Sul (*) 
            MS = 50, // Mato Grosso do Sul 
            MT = 51, // Mato Grosso 
            GO = 52, // Goiás 
            DF = 53 // Distrito Federal 
        }
        
        #endregion

    }
}

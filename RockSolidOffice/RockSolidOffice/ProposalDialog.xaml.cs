using System;
using System.Windows;
using System.Windows.Input;
using log4net;
using Wd = Microsoft.Office.Interop.Word;

namespace RockSolidOffice
{
    public partial class ProposalDialog : OfficeDialog
    {
        static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType); //See http://logging.apache.org/log4net/index.html

        readonly Wd.Document Doc;

        public ProposalDialog()
        {
            if (log.IsInfoEnabled) log.Info(System.Reflection.MethodBase.GetCurrentMethod().Name);
            InitializeComponent();
            txtProposalDate.Text = DateTime.Now.ToString("d MMMM yyyy");
            txtAcceptanceDate.Text = DateTime.Now.AddMonths(1).ToString("d MMMM yyyy");
            Loaded += (sender, e) => MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }
        public ProposalDialog(Wd.Document Doc)
            : this()
        {
            this.Doc = Doc;
        }

        void OK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (log.IsInfoEnabled) log.Info(System.Reflection.MethodBase.GetCurrentMethod().Name);

                var part = this.Doc.CustomXMLParts.SelectByNamespace("http://schemas.rocksolid.com.au/office")[1];
                part.SelectSingleNode("/ns0:root/ns0:clientname").Text = txtClientName.Text;
                part.SelectSingleNode("/ns0:root/ns0:clientabbreviatedname").Text = txtClientAbbreviatedName.Text;
                part.SelectSingleNode("/ns0:root/ns0:clientaddress").Text = txtClientAddress.Text;
                part.SelectSingleNode("/ns0:root/ns0:proposaldate").Text = txtProposalDate.Text;
                part.SelectSingleNode("/ns0:root/ns0:acceptancedate").Text = txtAcceptanceDate.Text;
                InvestmentSchedule.Update(this.Doc);

                this.DialogResult = true;
                this.Close();
            }
            catch (Exception ex)
            {
                log.Error(System.Reflection.MethodBase.GetCurrentMethod().Name, ex);
                MessageBox.Show(ex.Message, String.Format("{0} {1}", Settings.Caption, Settings.Version));
            }
        }

        void Cancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (log.IsInfoEnabled) log.Info(System.Reflection.MethodBase.GetCurrentMethod().Name);
                this.DialogResult = false;
                this.Close();
            }
            catch (Exception ex)
            {
                log.Error(System.Reflection.MethodBase.GetCurrentMethod().Name, ex);
                MessageBox.Show(ex.Message, String.Format("{0} {1}", Settings.Caption, Settings.Version));
            }
        }
    }
}

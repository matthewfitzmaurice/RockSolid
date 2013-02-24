using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Wd = Microsoft.Office.Interop.Word;
using log4net;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;

namespace RockSolidOffice
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType); //See http://logging.apache.org/log4net/index.html

        Office.IRibbonUI ribbon;
        Wd.Application app;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RockSolidOffice.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                if (log.IsInfoEnabled) log.Info(System.Reflection.MethodBase.GetCurrentMethod().Name);
                this.ribbon = ribbonUI;
                this.app = Globals.ThisAddIn.Application;
            }
            catch (Exception ex)
            {
                log.Error(System.Reflection.MethodBase.GetCurrentMethod().Name, ex);
                MessageBox.Show(ex.Message, String.Format("{0} {1}", Settings.Caption, Settings.Version));
            }
        }

        public void NewProposal_Click(Office.IRibbonControl control)
        {
            try
            {
                var doc = this.app.Documents.Add(GetProposalPath());
                var dialog = new ProposalDialog(doc);
                dialog.ShowDialog();
            }
            catch (FileNotFoundException ex)
            {
                if (log.IsWarnEnabled) log.Warn(System.Reflection.MethodBase.GetCurrentMethod().Name, ex);
                MessageBox.Show(ex.Message, String.Format("{0} {1}", Settings.Caption, Settings.Version), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Process.Start(GetTemplatesFolder().FullName);
            }
            catch (Exception ex)
            {
                log.Error(System.Reflection.MethodBase.GetCurrentMethod().Name, ex);
                MessageBox.Show(ex.Message, String.Format("{0} {1}", Settings.Caption, Settings.Version), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public bool IsProposal(Office.IRibbonControl control)
        {
            if (this.app.Documents.Count > 0)
            {
                Wd.Template template = (Wd.Template)this.app.ActiveDocument.get_AttachedTemplate();
                if (template.Name == "RockSolid CMS - Proposal template 2013.dotx")
                    return true;
            }
            return false;
        }

        public void UpdateInvestmentSchedule_Click(Office.IRibbonControl control)
        {
            try
            {
                InvestmentSchedule.Update(this.app.ActiveDocument);
            }
            catch (Exception ex)
            {
                log.Error(System.Reflection.MethodBase.GetCurrentMethod().Name, ex);
                MessageBox.Show(ex.Message, String.Format("{0} {1}", Settings.Caption, Settings.Version), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        static string GetProposalPath()
        {
            if (AtRockSolid())
            {
                string path = ConfigurationManager.AppSettings["ProposalTemplatePath"]; //TODO: Why is this returning '' for Sean? Replace with a Settings file
                if (!File.Exists(path))
                    path = @"S:\Proposals\10 Templates\00 Presales\RockSolid CMS - Proposal template 2013.dotx";
                if (!File.Exists(path))
                    throw new FileNotFoundException(string.Format("Unable to find template at\n'{0}'.\n\nIf this is the first time you have run the macro, you will need to copy 'RockSolid CMS - Proposal template 2013.dotx' to your network drive.\n\nNote: you must use the template provided because it contains Content Controls and Formulas used by the system code.", path));

                FileInfo file = new FileInfo(path);
                return file.FullName;
            }
            else // User is developer
            {
                var file = GetTemplatesFolder().GetFiles("RockSolid CMS - Proposal template 2013.dotx")[0];
                return file.FullName;
            }
        }

        static DirectoryInfo GetAssemblyFolder()
        {
            var uri = new Uri(Assembly.GetExecutingAssembly().CodeBase);
            var file = new FileInfo(uri.LocalPath);
            return file.Directory;
        }

        static DirectoryInfo GetTemplatesFolder()
        {
            var folder = GetAssemblyFolder();
            if (Directory.Exists(folder.FullName + "\\Templates"))
                return folder.GetDirectories("Templates")[0]; // This will be the location when installed via the MSI
            else
                return folder.Parent.GetDirectories("Templates")[0]; // This will be the location when run from Visual Studio
        }

        static bool AtRockSolid()
        {
            return Environment.UserName != "mjf" && Environment.UserName != "Moof";
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}

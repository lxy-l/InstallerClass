using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace AdvInstaller
{
    [RunInstaller(true)]
    public class AdvInstaller : Installer
    {
        private string SoftPath;
        private string SoftName;
        private int SoftEdition;
        private string SoftClass;
        private string configPath;
        public override void Install(IDictionary stateSaver)
        {
            string SourcePath = this.Context.Parameters["targetdir"];
          
            string[] arr = SourcePath.Split('\\');
            SoftPath = arr[2];
            configPath = $@"{SourcePath.Substring(0, SourcePath.Length - 1)}Config\App.config";
           
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(configPath);
            XmlNodeList nodeList = xmldoc.DocumentElement.ChildNodes;
            SoftName = nodeList[0].Attributes["value"].Value;
            SoftClass = nodeList[2].Attributes["value"].Value;
            SoftEdition = Convert.ToInt32(nodeList[1].Attributes["value"].Value);

            for (int i = 2014; i <= SoftEdition; i++)
            {
                string path = $@"C:\ProgramData\Autodesk\Revit\Addins\{i}";
                DirectoryInfo directoryInfo = new DirectoryInfo(path);
                if (!directoryInfo.Exists)
                {
                    directoryInfo.Create();
                }
                XDocument document = new XDocument();
                document.Declaration = new XDeclaration("1.0", "UTF-8", "");
                XElement root = new XElement("RevitAddIns");
                XElement element = new XElement("AddIn",
                        new XElement("Name", SoftName),
                        new XElement("Assembly", $@"C:\Program Files\{SoftPath}\Bin\{SoftName}_{i}.dll"),
                        new XElement("ClientId", Guid.NewGuid()),
                        new XElement("FullClassName", $"{SoftName}.{SoftClass}"),
                        new XElement("VendorId", "ADSK"),
                        new XElement("VendorDescription", "Autodesk, www.autodesk.com"));
                element.SetAttributeValue("Type", "Application");
                root.Add(element);
                document.Add(root);
                document.Save($@"{path}\{SoftName}_{i}.addin");
            }
            base.Install(stateSaver);
        }


        public override void Uninstall(IDictionary savedState)
        {
            string path = this.Context.Parameters["targetdir"];
            string[] arr = path.Split('\\');
            SoftPath = arr[2];
            configPath = $@"{path.Substring(0, path.Length - 1)}Config\App.config";
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(configPath);
            XmlNodeList nodeList = xmldoc.DocumentElement.ChildNodes;
            SoftName = nodeList[0].Attributes["value"].Value;
            SoftEdition = Convert.ToInt32(nodeList[1].Attributes["value"].Value);
            for (int i = 2014; i <= SoftEdition; i++)
            {
                string file = $@"C:\ProgramData\Autodesk\Revit\Addins\{i}\{SoftName}_{i}.addin";
                FileInfo fileInfo = new FileInfo(file);
                if (fileInfo.Exists)
                {
                    fileInfo.Delete();
                }

            }
            base.Uninstall(savedState);
        }
    }
}

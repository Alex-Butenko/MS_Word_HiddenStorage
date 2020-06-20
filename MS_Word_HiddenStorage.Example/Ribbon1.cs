using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System.Windows.Forms;

namespace OfficeTools.Examples {
    public partial class Ribbon1 {
        void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        void button1_Click(object sender, RibbonControlEventArgs e) {
            string key = "test_key";
            HiddenStorage storage = new HiddenStorage(Globals.ThisAddIn.Application.ActiveDocument, key);

            string textBeforeEdit = "";
            try {
                textBeforeEdit = storage.Read();
            }
            catch (KeyNotFoundException) {
                // means nothing saved yet
            }

            MetadataEditForm editForm = new MetadataEditForm(textBeforeEdit);
            if (editForm.ShowDialog() == DialogResult.OK) {
                storage.Write(editForm.Text);
            }
        }
    }
}
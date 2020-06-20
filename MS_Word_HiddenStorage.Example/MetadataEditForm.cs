using System;
using System.Windows.Forms;

namespace OfficeTools.Examples {
    public partial class MetadataEditForm : Form {
        public MetadataEditForm() {
            InitializeComponent();
        }

        public MetadataEditForm(string text) : this() {
            textBox1.Text = Text = text;
        }

        public string Text { get; private set; }

        void ButtonSave_Click(object sender, EventArgs e) {
            Text = textBox1.Text;
            const string message = "The text is saved as metadata in the file."
                + "You can save the file, close and reopen, open the same dialog and ensure the text is still the same.";
            MessageBox.Show(message, "Save metadata information", MessageBoxButtons.OK);
            DialogResult = DialogResult.OK;
            Close();
        }

        void ButtonCancel_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
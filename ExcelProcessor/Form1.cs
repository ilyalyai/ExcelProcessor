using EpPlus;

namespace ExcelProcessor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.FilterIndex = 1;
                //openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    var epPlus = new EpPlusClass(filePath);
                    MessageBox.Show(epPlus.Run(), "New file at path: ", MessageBoxButtons.OK);
                }
            }
        }
    }
}
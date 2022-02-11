using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToolConfrontoExcel
{
    public partial class Form1 : Form
    {

        private DataTable DT_Excel1;
        private DataTable DT_Excel2;

        public Form1()
        {
            InitializeComponent();

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        public void button1_Click(object sender, EventArgs e)
        {
            //Viene letto il path del primo excel
            string varpath1_temp = textBox1.Text;
            string varpath1 = varpath1_temp.Replace("\"", "");
                        
            // Viene convertito il primo excel in DataTable
            DT_Excel1 = TrattaExcel.Excel_To_DataTable(varpath1, 0);
            Utility.Scrivi_Log(listBox1, "Excel 1 aggiunto");

            
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Viene letto il path del secondo excel
            string varpath2_temp = textBox2.Text;
            string varpath2 = varpath2_temp.Replace("\"", "");
            
            // Viene convertito il secondo excel in DataTable
            DT_Excel2 = TrattaExcel.Excel_To_DataTable(varpath2, 0);
            Utility.Scrivi_Log(listBox1, "Excel 2 aggiunto");
            
            
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        public void button3_Click(object sender, EventArgs e)
        {
            int test = Controlli.Cfr_Numero_Record(DT_Excel1, DT_Excel2);
            
            switch (test)
            {
                case 0:
                Utility.Scrivi_Log(listBox1, "Numero record uguale");
                break;
                case 1:
                Utility.Scrivi_Log(listBox1, "Numero record diverso");
                break;
                case 2:
                Utility.Scrivi_Log(listBox1, "Errore lettura Excel 1");
                break;
                case 3:
                Utility.Scrivi_Log(listBox1, "Errore lettura Excel 2");
                break;
            }
            //Viene scritto il log 
            if (Controlli.Count_Numero_Record(DT_Excel1).flag == true)
            {
                int num_record = Controlli.Count_Numero_Record(DT_Excel1).result + 1; // +1 perché parte da 0
            Utility.Scrivi_Log(listBox1, "Numero record Excel 1: " + num_record);
            }
            else Utility.Scrivi_Log(listBox1, "Errore lettura Excel 1");

            if (Controlli.Count_Numero_Record(DT_Excel2).flag == true) { 
            int num_record2 = Controlli.Count_Numero_Record(DT_Excel2).result + 1; // +1 perché parte da 0
            Utility.Scrivi_Log(listBox1, "Numero record Excel 2: " + num_record2);
            }
            else Utility.Scrivi_Log(listBox1, "Errore lettura Excel 2");


        }

        private void button4_Click(object sender, EventArgs e)
        {
            Utility.Scrivi_Log(listBox1, "Inizio confronto dati...");
            Controlli.Cfr_Record(DT_Excel1, DT_Excel2);
        }
    }
}

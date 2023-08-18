using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using LiteDB;


namespace RisXpert_App
{
    public partial class Form1 : Form
    {
        BindingList<RiskAnalysis> ListaRiesgos = new BindingList<RiskAnalysis>();
        [AttributeUsage(AttributeTargets.Property)]
        public class RequiresValidationAttribute : Attribute { } //Creacion de Atributo 

        public Form1()
        {
            InitializeComponent();
            lblDate.Text = DateTime.Today.ToString();
            dtgvFase1.DataSource = ListaRiesgos;    //Configuracion dtgv Fase 1
            dtgvFase1.Columns["_id"].Visible = false;
            dtgvFase1.Columns["S"].Visible = false;
            dtgvFase1.Columns["F"].Visible = false;
            dtgvFase1.Columns["P"].Visible = false;
            dtgvFase1.Columns["A"].Visible = false;
            dtgvFase1.Columns["V"].Visible = false;
            dtgvFase1.Columns["E"].Visible = false;
            dtgvFase1.Columns["CR"].Visible = false;
            dtgvFase1.Columns["Pb"].Visible = false;
            dtgvFase1.Columns["Er"].Visible = false;
            dtgvFase1.Columns["Clasificacion"].Visible = false;

            dtgvFase1.Columns["Analista"].ReadOnly = true;
            dtgvFase1.Columns["Activo"].ReadOnly = true;

            dtgvFase2.DataSource = ListaRiesgos;    //Configuracion dtgv Fase 2
            dtgvFase2.Columns["CR"].Visible = false;
            dtgvFase2.Columns["Pb"].Visible = false;
            dtgvFase2.Columns["Er"].Visible = false;
            dtgvFase2.Columns["_id"].Visible = false;
            dtgvFase2.Columns["Clasificacion"].Visible = false;
            dtgvFase2.Columns["Analista"].Visible = false;
            dtgvFase2.Columns["Activo"].Visible = false;

            dtgvFase2.Columns["Riesgo"].ReadOnly = true;
            dtgvFase2.Columns["Dano"].ReadOnly = true;

            dtgvFase3.DataSource = ListaRiesgos;    //Configuracion dtgv Fase 3
            dtgvFase3.Columns["_id"].Visible = false;
            dtgvFase3.Columns["S"].Visible = false;
            dtgvFase3.Columns["F"].Visible = false;
            dtgvFase3.Columns["P"].Visible = false;
            dtgvFase3.Columns["A"].Visible = false;
            dtgvFase3.Columns["V"].Visible = false;
            dtgvFase3.Columns["E"].Visible = false;
            dtgvFase3.Columns["Clasificacion"].Visible = false;
            dtgvFase3.Columns["Analista"].Visible = false;
            dtgvFase3.Columns["Activo"].Visible = false;

            dtgvFase3.Columns["Riesgo"].ReadOnly = true;
            dtgvFase3.Columns["Dano"].ReadOnly = true;
            dtgvFase3.Columns["CR"].ReadOnly = true;
            dtgvFase3.Columns["ER"].ReadOnly = true;
            dtgvFase3.Columns["Pb"].ReadOnly = true;

            dtgvFase4.DataSource = ListaRiesgos;    //Configuracion dtgv Fase4

            dtgvFase4.Columns["_id"].Visible = false;
            dtgvFase4.Columns["S"].Visible = false;
            dtgvFase4.Columns["F"].Visible = false;
            dtgvFase4.Columns["P"].Visible = false;
            dtgvFase4.Columns["A"].Visible = false;
            dtgvFase4.Columns["V"].Visible = false;
            dtgvFase4.Columns["E"].Visible = false;
            dtgvFase4.Columns["Analista"].Visible = false;
            dtgvFase4.Columns["Activo"].Visible = false;
            dtgvFase4.Columns["CR"].Visible = false;
            dtgvFase4.Columns["Pb"].Visible = false;

            dtgvFase4.Columns["Riesgo"].ReadOnly = true;
            dtgvFase4.Columns["Dano"].ReadOnly = true;
            dtgvFase4.Columns["ER"].ReadOnly = true;
            dtgvFase4.Columns["Clasificacion"].ReadOnly = true;
        }

        private void btnNewRisk_Click(object sender, EventArgs e)
        {
            RiskAnalysis X = new RiskAnalysis
            {
                Analista = txtAnalystName.Text,
                Activo = txtActive.Text
            };
            ListaRiesgos.Add(X);
            UpdateData(sender, e);
        }
        private void dgtvValues_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            DataGridViewColumn column = dtgvFase2.Columns[e.ColumnIndex];

            PropertyInfo propertyInfo = typeof(RiskAnalysis).GetProperty(column.DataPropertyName);

            if (propertyInfo != null && Attribute.IsDefined(propertyInfo, typeof(RequiresValidationAttribute)))
            {
                if (!string.IsNullOrEmpty(e.FormattedValue.ToString()))
                {
                    if (!int.TryParse(e.FormattedValue.ToString(), out int numericValue))
                    {
                        e.Cancel = true;
                        MessageBox.Show("Inserte valor numérico.", "Valor no Válido", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (numericValue < 1 || numericValue > 5)
                    {
                        e.Cancel = true;
                        MessageBox.Show("Inserte valor entre 1 y 5.", "Valor no Válido", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        private void dtgvFase2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cellValues = dtgvFase2.Rows[e.RowIndex].Cells[e.ColumnIndex];

            if (e.ColumnIndex >= 3)
            {
                if (int.TryParse(cellValues.Value?.ToString(), out int numericValue))
                {
                    switch (numericValue)
                    {
                        case 1:
                            cellValues.Style.BackColor = Color.Lime;
                            break;
                        case 2:
                            cellValues.Style.BackColor = Color.LimeGreen;
                            break;
                        case 3:
                            cellValues.Style.BackColor = Color.Yellow;
                            break;
                        case 4:
                            cellValues.Style.BackColor = Color.Orange;
                            break;
                        case 5:
                            cellValues.Style.BackColor = Color.Red;
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        private void tabControl_Selected(object sender, TabControlEventArgs e)
        {
            for (int i = 0; i < dtgvFase1.Rows.Count; i++)
            {
                var CurrentRow = ListaRiesgos[i];

                int I = CurrentRow.F * CurrentRow.S;
                int D = CurrentRow.P * CurrentRow.E;
                CurrentRow.Pb = CurrentRow.A * CurrentRow.V;
                CurrentRow.CR = I + D;
                CurrentRow.ER = CurrentRow.Pb * CurrentRow.CR;
                CurrentRow.Clasificacion = Classify(i);
                //ListaRiesgos.OrderBy()
            }
        }
        private void UpdateData(object sender, EventArgs e)
        {

            for (int i = 0; i < dtgvFase1.Rows.Count; i++)
            {
                var Fase1Rows = dtgvFase1.Rows[i];

                Fase1Rows.Cells[0].Value = txtAnalystName.Text;
                Fase1Rows.Cells[1].Value = txtActive.Text;
            }
        }
        public class RiskAnalysis
        {
            public string Analista { get; set; }
            public string Activo { get; set; }
            public string Riesgo { get; set; }
            public string Dano { get; set; }
            public int CR { get; set; }
            public int Pb { get; set; }
            public int ER { get; set; }
            public string Clasificacion { get; set; }
            public int _id { get; set; }
            [RequiresValidation] public int S { get; set; }
            [RequiresValidation] public int F { get; set; }
            [RequiresValidation] public int P { get; set; }
            [RequiresValidation] public int A { get; set; }
            [RequiresValidation] public int V { get; set; }
            [RequiresValidation] public int E { get; set; }   
        }

        private void SaveData()
        {
            using (var db = new LiteDatabase(@"C:\Users\HP\Desktop\Test.db"))
            {
                string ColName = txtActive.Text + "_" + txtID.Text;
                var col = db.GetCollection<RiskAnalysis>(ColName);
                if (!db.CollectionExists(ColName))
                {
                    col.InsertBulk(ListaRiesgos);
                }
                else
                {
                    col.Update(ListaRiesgos);
                };
            }
        }
        
        private string Classify(int i)
        {
            var ERValue = ListaRiesgos[i].ER;

            if (ERValue >= 2 && ERValue <= 250)
            {
                return "Muy Pequeño";
            }
            else if (ERValue >= 251 && ERValue <= 500)
            {
                return "Pequeño";
            }
            else if (ERValue >= 501 && ERValue <= 750)
            {
                return "Normal";
            }
            else if (ERValue >= 751 && ERValue <= 1000)
            {
                return "Grande";
            }
            else if (ERValue >= 1001 && ERValue <= 1250)
            {
                return "Elevado";
            }
            else return "";
        }

        private void btnSave2_Click(object sender, EventArgs e)
        {
            SaveData();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CreateCatalog
{
    public partial class frmSelectProjectName : Form
    {
        int scu;
        public static string sOtherProjectName;
        public frmSelectProjectName()
        {
            InitializeComponent();
        }

        private void frmSelectProjectName_Load(object sender, EventArgs e)
        {
            DataSet ds = PIMS_ClassLib.DataAccess.GetDataSet(frmMain.sConstrBase, "select project_no,project_name from tb_other_project_name where project_no = '" + frmMain.sProjectNo + "'", out scu);
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                cboProjectName.Items.Add(row[1].ToString());
            }
        }

        private void cboProjectName_SelectedIndexChanged(object sender, EventArgs e)
        {
            sOtherProjectName = cboProjectName.Text;
            this.Close();
        }
    }
}

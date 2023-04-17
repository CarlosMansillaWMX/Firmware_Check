using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Traceability;
namespace Maxim_verification
{
    public partial class Form2 : Form
    {
        public static string user = "";
        public static string userid = "";
        SFC_System wmx_ts = new SFC_System();
        public Form2()
        {
            InitializeComponent();
        }

        private void tblogin_Click(object sender, EventArgs e)
        {
            string usuario="";
            usuario = tbuser.Text.ToUpper();
            user = tbuser.Text.ToUpper();
            string[] subs = new string[6] { "", "", "", "", "", "" };
            string response = "";
            try
            {
                response = wmx_ts.AuthenticateEmployee(user);
                subs = response.Split(',');

            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            if (subs[0] == "OK")
            {
                user = subs[2];
                userid = usuario;
                tbuser.Text = "";
                Form1 frm = new Form1(this);
                frm.Show();
                this.Hide();
            }
            else
                MessageBox.Show(response);
            tbuser.Text = "";

        }

        private void tbclose_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea Salir?", "Salir", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
                Application.Exit();
        }
    }
}

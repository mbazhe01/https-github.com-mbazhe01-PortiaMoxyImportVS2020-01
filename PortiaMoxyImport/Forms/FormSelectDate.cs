using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PortiaMoxyImport.Forms
{
    public partial class FormSelectDate : Form
    {
        public FormSelectDate()
        {
            InitializeComponent();
        }

        private void FormSelectDate_Load(object sender, EventArgs e)
        {
            // set initial date
            dateTimePicker1.Value = DateTime.Today;
        }

        public DateTime getSelectedDate()
        {
            return dateTimePicker1.Value;
        }
    }
}

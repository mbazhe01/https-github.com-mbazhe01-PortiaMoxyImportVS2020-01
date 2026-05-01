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
        private Boolean pickToday;
        public FormSelectDate(bool pickToday)
        {
            InitializeComponent();
            this.pickToday = pickToday;
        }

        private void FormSelectDate_Load(object sender, EventArgs e)
        {
          

        }

        private static DateTime GetPreviousBusinessDay(DateTime date)
        {
            date = date.Date.AddDays(-1);

            while (date.DayOfWeek == DayOfWeek.Saturday ||
                   date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(-1);
            }

            return date;
        }

        public DateTime getSelectedDate()
        {
            return dateTimePicker1.Value;
        }

        public DateTime getTodayDate()
        {
            dateTimePicker1.Value = DateTime.Today;
            return dateTimePicker1.Value;
        }

        private void FormSelectDate_Shown(object sender, EventArgs e)
        {
            // set initial date to previous business day
            if (pickToday)
            {
                dateTimePicker1.Value = DateTime.Today;
            }
            else
                dateTimePicker1.Value = GetPreviousBusinessDay(DateTime.Today);

        }
    }
}

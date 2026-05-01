using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PortiaMoxyImport.Redesign
{
    public sealed class RichTextBoxConversionReporter : IConversionReporter
    {
        private readonly RichTextBox _output;
        private readonly Label _status;

        public RichTextBoxConversionReporter(RichTextBox output, Label status)
        {
            _output = output ?? throw new ArgumentNullException(nameof(output));
            _status = status ?? throw new ArgumentNullException(nameof(status));
        }

        public void Info(string message)
        {
            AppendLine(message);
        }

        public void Success(string message)
        {
            // You can keep your green text behavior here
            AppendLine(message);
            // Optionally color or use a helper
        }

        public void Warning(string message)
        {
            AppendLine(message);
        }

        public void Error(string message)
        {
            
            AppendLine(message, System.Drawing.Color.Red);
        }

        public void SetStatus(string message)
        {
            _status.Text = message;
            _status.Refresh();
        }

        public void Clear()
        {
            _output.Clear();
            _status.Text = string.Empty;
        }

        //private void AppendLine(string message)
        //{
        //    if (string.IsNullOrEmpty(message))
        //        return;

        //    _output.AppendText(message + Environment.NewLine);
        //    _output.ScrollToCaret();
        //    _output.Refresh();
        //}

        private void AppendLine(string message, System.Drawing.Color? color = null)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            _output.SelectionStart = _output.TextLength;
            _output.SelectionLength = 0;

            if (color.HasValue)
                _output.SelectionColor = color.Value;
            else
                _output.SelectionColor = _output.ForeColor;

            _output.AppendText(message + Environment.NewLine);

            // Reset color back to default
            _output.SelectionColor = _output.ForeColor;

            _output.ScrollToCaret();
        }


    }

}

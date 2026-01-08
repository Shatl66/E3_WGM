using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    public partial class E3CommonControl : UserControl
    {
        public event EventHandler SynchronizeClicked;

        public E3CommonControl()
        {
            InitializeComponent();
        }

        private void buttonSynch_Click(object sender, EventArgs e)
        {
            SynchronizeClicked?.Invoke(this, EventArgs.Empty); // Генерируем событие на которое подписана E3WGMForm
        }

    }
}

using e3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    public partial class E3WGMForm : Form
    {

        public static e3Application app;
        public static e3Job job;

        // Start Делем недоступным закрытие формы по "крестику"
        /*
        private const int CP_NOCLOSE_BUTTON = 0x200;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CP_NOCLOSE_BUTTON;
                return cp;
            }
        }
        */
        // End Делем недоступным закрытие формы по "крестику"

        public E3WGMForm()
        {
            InitializeComponent();        
        }




        private void E3WGMForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            DisconnectFromE3Series();
        }

        public void DisconnectFromE3Series()
        {
            if (app == null)
                return;

            // разрываем соединение с E3.Series.
            Marshal.ReleaseComObject(app);
            Marshal.ReleaseComObject(job);

            // Обнуляем поля, ссылавшиеся на COM-объекты
            app = null;
            job = null;
            //dev = null;
            //Cab = null;
            //Pin = null;
            //Core = null;

            // принудительно удаляем пустые (null) ссылки
            GC.Collect();
        }

    }
}

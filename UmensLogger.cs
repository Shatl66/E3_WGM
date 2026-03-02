using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    public static class UmensLogger // не нужно создавать экземпляр класса. Можно обращаться напрямую: Logger.Log("сообщение")
    {
        public static TextBox LogControl { get; set; } // Хранит ссылку на RichTextBox (поле E3Log из вашей формы). с доступом на чтение и запись

        // Windows Forms элементы управления "живут" в основном потоке (UI thread)
        public static void Log(string message)
        {
            if (LogControl == null) return;

            if (LogControl.InvokeRequired) // проверяет потоки
            {
                LogControl.Invoke(new Action<string>(Log), message); // метод Log вызовет сам себя, но уже в правильном потоке
            }
            else
            {
                LogControl.AppendText($"{message}\r\n");
                LogControl.SelectionStart = LogControl.Text.Length;
                LogControl.ScrollToCaret(); // прокручивает RichTextBox вниз, чтобы было видно последнее сообщение
            }
        }

    }
}

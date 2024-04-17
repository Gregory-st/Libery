using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ControlsWorker
{
    public static class ControlsApplication
    {

        public static T CloneElement<T>(T prototype) where T : Control
        {
            T instance = Activator.CreateInstance<T>();
            PropertyInfo[] info = typeof(T).GetProperties();

            foreach(var i in info)
            {
                if(i.CanWrite && i.Name != "WindowTarget")
                {
                    i.SetValue(instance, i.GetValue(prototype, null), null);
                }
            }

            return instance;
        }

        public static T CloneChild<T>(T prototype) where T : Control
        {
            T instance = CloneElement(prototype);

            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            bool continer = false;

            foreach(var i in propertyInfos)
            {
                continer = i.Name == "Controls";
                if (continer) break;
            }

            if(!continer) return instance;

            foreach(Control i in prototype.Controls)
            {
                if (i is Button button)
                    instance.Controls.Add(CloneElement(button));
                else if (i is TextBox textBox)
                    instance.Controls.Add(CloneElement(textBox));
            }

            return instance;
        }
    }
}

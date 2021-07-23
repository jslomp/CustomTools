using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CustomTools
{
    
    class ListBoxSearch
    {
        TextBox textBox1 = null;
        ListBox listBox1 = null;
        CheckedListBox checkbox = null;
        List<string> removed = new List<string>();

        internal void setListbox(CheckedListBox listBox1)
        {
            // todo: If item is checked, don't remove them.

            this.listBox1 = listBox1;
            this.checkbox = listBox1;
        }

        internal void setListbox(ListBox listBox1)
        {
            this.listBox1 = listBox1;
        }

        internal void setTextField(TextBox textBox1)
        {
            this.textBox1 = textBox1;

            this.textBox1.TextChanged += delegate
            {
                changed();
            };
        }


        public void changed()
        {


            string search = textBox1.Text.ToLower();
            string[] search_or = search.Split(',');
            for(int i = listBox1.Items.Count-1; i >= 0; i--)
            {
                var item = listBox1.Items[i];

                Boolean found = false;
                foreach (string s in search_or)
                {
                    if (item.ToString().ToLower().Contains(s))
                    {
                        found = true;
                    }
                }
                if (!found)
                {
                    listBox1.Items.Remove(item);
                    removed.Add(item.ToString());
                }
            }

            for (int i = removed.Count - 1; i >= 0; i--)
            {
                var item = removed[i];

                Boolean found = false;
                foreach (string s in search_or)
                {
                    if (item.ToString().ToLower().Contains(s))
                    {
                        found = true;
                    }
                }

                if (found)
                {
                    
                    listBox1.Items.Add(item);
                    removed.Remove(item);
                }
            }

            listBox1.Sorted = true;

        }
    }
}

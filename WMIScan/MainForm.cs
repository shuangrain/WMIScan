using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;

namespace WMIScan
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("本程式提供簡單的掃描WMI是否含有HAO123的入侵，並提供刪除功能 !!");
            sb.AppendFormat("GitHub: https://github.com/shuangrain/WMIScan");

            MessageBox.Show(sb.ToString());
        }

        private void btnScan_Click(object sender, EventArgs e)
        {
            ManagementObjectCollection queryCollection = getList();

            action(() =>
            {
                Dictionary<string, string> dty = new Dictionary<string, string>();
                int idx = 0;
                foreach (ManagementObject mo in queryCollection)
                {
                    string name = mo["Name"].ToString();
                    string scriptText = mo["ScriptText"].ToString();
                    string url = string.Empty;

                    if (scriptText.Contains("http"))
                    {
                        int first = scriptText.IndexOf("http") - 1;
                        if (scriptText[first] != '"')
                        {
                            //跳出
                            break;
                        }
                        int end = first;
                        while (end < (scriptText.Length - 1))
                        {
                            if (scriptText[end] == '"')
                            {
                                //跳出
                                break;
                            }

                            end++;
                        }
                        if (scriptText[end] != '"')
                        {
                            //跳出
                            break;
                        }
                        first++;
                        end--;

                        url = scriptText.Substring(first, end);
                    }
                    if (!string.IsNullOrWhiteSpace(url))
                    {
                        string id = string.Join("", ((byte[])mo["CreatorSID"]).Select(x => x.ToString()));
                        idx++;
                        dty.Add(id, $"{idx}: {name} | {url}");
                    }
                }

                if (dty.Count > 0)
                {
                    list.DataSource = new BindingSource(dty, null); ;

                    list.DisplayMember = "Value";
                    list.ValueMember = "Key";
                }
                else
                {
                    list.DataSource = null;
                }
            }, "Not Found !!");
        }

        private ManagementObjectCollection getList()
        {
            ManagementScope ms = new ManagementScope(@"\root\cimv2");
            ms.Connect();

            ObjectQuery oq = new ObjectQuery("SELECT * FROM ActiveScriptEventConsumer");
            ManagementObjectSearcher query = new ManagementObjectSearcher(ms, oq);
            ManagementObjectCollection queryCollection = query.Get();

            return queryCollection;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (list.Items.Count > 0 && list.SelectionMode == SelectionMode.One)
            {
                string id = ((KeyValuePair<string, string>)list.SelectedItem).Key;
                ManagementObjectCollection queryCollection = getList();
                action(() =>
                {
                    foreach (ManagementObject mo in queryCollection)
                    {
                        string tmp = string.Join("", ((byte[])mo["CreatorSID"]).Select(x => x.ToString()));
                        MessageBox.Show(tmp + "\r\n" + id);
                        if (tmp == id)
                        {
                            mo.Delete();
                        }
                    }

                    MessageBox.Show("Success !!");
                }, "Error !!");
            }
        }

        private void action(Action action, string errorMsg)
        {
            try
            {
                action();
            }
            catch
            {
                MessageBox.Show(errorMsg);
            }
        }
    }
}

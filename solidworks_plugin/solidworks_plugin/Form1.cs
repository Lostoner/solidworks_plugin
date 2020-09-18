﻿using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace solidworks_plugin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ISldWorks SwApp = doc_class.ConnectToSolidWorks();
            if(SwApp != null)
            {
                string msg = "This message from C#. solidworks version is " + SwApp.RevisionNumber();
                SwApp.SendMsgToUser(msg);
            }
        }
    }
}

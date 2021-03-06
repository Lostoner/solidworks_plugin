﻿using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace solidworks_plugin
{

    public partial class Form1 : Form
    {
        string path = @"D:\F\三维模型库\Parts_WithFeature";
        string license = @"Y6T54-H6G5N-H65GG-6YT5M-7H65B";
        public enum SwDmDocumentVersion_e
        {
            SwDmDocumentVersionSOLIDWORKS_2000 = 1500,
            SwDmDocumentVersionSOLIDWORKS_2001 = 1750,
            SwDmDocumentVersionSOLIDWORKS_2001Plus = 1950,
            SwDmDocumentVersionSOLIDWORKS_2003 = 2200,
            SwDmDocumentVersionSOLIDWORKS_2004 = 2500,
            SwDmDocumentVersionSOLIDWORKS_2005 = 2800,
            SwDmDocumentVersionSOLIDWORKS_2006 = 3100,
            SwDmDocumentVersionSOLIDWORKS_2007 = 3400,
            SwDmDocumentVersionSOLIDWORKS_2008 = 3800,
            SwDmDocumentVersionSOLIDWORKS_2009 = 4100,
            SwDmDocumentVersionSOLIDWORKS_2010 = 4400,
            SwDmDocumentVersionSOLIDWORKS_2011 = 4700,
            SwDmDocumentVersionSOLIDWORKS_2012 = 5000,
            SwDmDocumentVersionSOLIDWORKS_2013 = 6000,
            SwDmDocumentVersionSOLIDWORKS_2014 = 7000,
            SwDmDocumentVersionSOLIDWORKS_2015 = 8000,
            SwDmDocumentVersionSOLIDWORKS_2016 = 9000,
            SwDmDocumentVersionSOLIDWORKS_2017 = 10000,
            SwDmDocumentVersionSOLIDWORKS_2018 = 11000
        }


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

        private void button2_Click(object sender, EventArgs e)
        {
            doc_class.OpenDocument();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ISldWorks swApp = doc_class.ConnectToSolidWorks();

            if (swApp != null)
            {
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;

                Feature swFeat = (Feature)swModel.FirstFeature();

                doc_class.TraverseFeature(swFeat, true);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string[] files;

            files = doc_class.GetAllFile(path);

            foreach(string f in files)
            {
                Debug.Print(f);
                Debug.Print("\n");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            doc_class.DifferenceToFeatures();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            doc_class.TrueFeature_NameOnly();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string[] files;

            files = doc_class.GetAllFile(path);

            foreach (string f in files)
            {
                doc_class.OpenAndClose(f);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string license = @"Y6T54-H6G5N-H65GG-6YT5M-7H65B";
            string path = @"D:\F\三维模型库\Parts_WithFeature\11.SLDPRT";

            doc_class.DirectDocument(path, license);
        }
    }
}

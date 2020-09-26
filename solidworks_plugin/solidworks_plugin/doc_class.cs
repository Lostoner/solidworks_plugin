using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace solidworks_plugin
{
    public class doc_class
    {
        public static ISldWorks SwApp { get; private set; }

        public static ISldWorks ConnectToSolidWorks()
        {
            if(SwApp != null)
            {
                return SwApp;
            }
            Debug.Print("Connect to solidworks...");
            try
            {
                SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch (COMException)
            {
                try
                {
                    SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.23");//2015
                }
                catch (COMException)
                {
                    try
                    {
                        SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.26");//2018
                    }
                    catch (COMException)
                    {
                        MessageBox.Show("Could not connect to SolidWorks.", "SolidWorks", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        SwApp = null;
                    }
                }
            }
            if(SwApp != null)
            {
                Debug.Print("Connection succeed.");
                return SwApp;
            }
            else
            {
                Debug.Print("Connection failed.");
                return null;
            }
        }

        public static void OpenDocument()
        {
            string DocPath = @"D:\F\三维模型库\CAD模型库\[130616]Previous_model_set\26_Wheel\DEFAULT_12190-NY-Wheel(NY-150-50-60-20).sldprt";

            //string partDefaultTemplate = SwApp.GetDocumentTemplate((int)swDocumentTypes_e.swDocPART, "", 0, 0, 0);

            SwApp.OpenDoc(DocPath, (int)swDocumentTypes_e.swDocPART);
            SwApp.SendMsgToUser("Open file complete.");

        }

        public static void TraverseFeature(Feature Fea, bool TopLevel)
        {
            Feature curFea = default(Feature);
            curFea = Fea;

            while(curFea != null)
            {
                Debug.Print(curFea.Name);

                Feature subfeat = default(Feature);
                subfeat = (Feature)curFea.GetFirstSubFeature();

                while(subfeat != null)
                {
                    TraverseFeature(subfeat, false);
                    Feature nextSubFeat = default(Feature);
                    nextSubFeat = (Feature)subfeat.GetNextSubFeature();
                    subfeat = nextSubFeat;
                    nextSubFeat = null;
                }

                subfeat = null;

                Feature nextFeat = default(Feature);

                if(TopLevel)
                {
                    nextFeat = (Feature)curFea.GetNextFeature();
                }
                else
                {
                    nextFeat = null;
                }

                curFea = nextFeat;
                nextFeat = null;
            }
        }

        public static string[] GetAllFile(string path)
        {
            string[] files = System.IO.Directory.GetFiles(path, "*", System.IO.SearchOption.AllDirectories);

            return files;
        }

        public static void AutoFeature()
        {
            ISldWorks swApp = ConnectToSolidWorks();
            var swModel = (ModelDoc2)swApp.ActiveDoc;


        }

        public static void DifferenceToFeatures()
        {
            ISldWorks swApp = ConnectToSolidWorks();
            var swModel = (ModelDoc2)swApp.ActiveDoc;

            Feature swFeat = (Feature)swModel.FirstFeature();
            //Feature subFeat = (Feature)swFeat.GetFirstSubFeature();

            Feature nextFeat = default(Feature);
            while(swFeat != null)
            {
                Debug.Print(swFeat.Name);
                nextFeat = swFeat.GetNextFeature();
                swFeat = nextFeat;
                nextFeat = null;
            }
        }

        public static void TrueFeature_NameOnly()
        {
            ISldWorks swApp = ConnectToSolidWorks();
            var swModel = (ModelDoc2)swApp.ActiveDoc;

            Feature swFeat = (Feature)swModel.FirstFeature();
            //int i = 0;
            while(swFeat != null)
            {
                Feature SubFeature = swFeat.GetFirstSubFeature();
                while(SubFeature != null)
                {
                    if(SubFeature.GetTypeName2() == "ProfileFeature")
                    {
                        Sketch swSketch = SubFeature.GetSpecificFeature2();
                        //i++;
                        //Debug.Print("\: ", i);

                        //string skePrint = swSketch.ToString();
                        string skePrint = SubFeature.Name;
                        Debug.Print(skePrint);
                    }

                    Feature NextSubFeat = SubFeature.GetNextSubFeature();
                    SubFeature = NextSubFeat;
                    NextSubFeat = null;
                }

                Feature NextFea = swFeat.GetNextFeature();
                swFeat = NextFea;
                NextFea = null;
            }
        }

        public static void OpenAndClose(string path)
        {
            Debug.Print(path);
            ISldWorks SwApp = ConnectToSolidWorks();
            //string partDefaultTemplate = SwApp.GetDocumentTemplate((int)swDocumentTypes_e.swDocPART, "", 0, 0, 0);
            SwApp.OpenDoc(path, (int)swDocumentTypes_e.swDocPART);

            TrueFeature_NameOnly();

            SwApp.CloseDoc(path);
        }
    }
}

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdocumentmgr;
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

        public static ISldWorks ConnectToSolidWorks()               //连接函数，连接solidworks，任何与solidworks有关的操作必须在经过连接之后才能进行
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

        public static void OpenDocument()                               //使solidworks打开指定文件，仅作测试
        {
            //注意：本函数需要先使用连接按钮或事先调用连接函数之后才能使用

            string DocPath = @"D:\F\三维模型库\CAD模型库\[130616]Previous_model_set\26_Wheel\DEFAULT_12190-NY-Wheel(NY-150-50-60-20).sldprt";    //指定一文件路径

            //string partDefaultTemplate = SwApp.GetDocumentTemplate((int)swDocumentTypes_e.swDocPART, "", 0, 0, 0);

            SwApp.OpenDoc(DocPath, (int)swDocumentTypes_e.swDocPART);       //SwApp为连接函数中定义的类，调用其OpenDoc方法打开文件
            SwApp.SendMsgToUser("Open file complete.");

        }

        public static void TraverseFeature(Feature Fea, bool TopLevel)          //输入单一模型文件的首特征，自动地遍历该文件的所有特征
        {
            //本函数需要在调用时输入一个特征树中的特征，而后自动地遍历剩余的特征

            Feature curFea = default(Feature);          //创建Feature类实例
            curFea = Fea;

            while(curFea != null)                                       //持续遍历
            {
                Debug.Print(curFea.Name);                       //在Debug窗口输出该特征的名字

                Feature subfeat = default(Feature);         //创建另一个Feature类实例，用以遍历某一特征的子特征
                subfeat = (Feature)curFea.GetFirstSubFeature();         //获得当前curFea的首个子特征

                while(subfeat != null)                              //持续遍历curFea的子特征
                {
                    TraverseFeature(subfeat, false);
                    Feature nextSubFeat = default(Feature);
                    nextSubFeat = (Feature)subfeat.GetNextSubFeature();     //获取下一个子特征
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

        public static string[] GetAllFile(string path)                                        //获取指定路径下所有文件，返回各文件路径组成的的字符串数组，全部使用C#函数
        {
            string[] files = System.IO.Directory.GetFiles(path, "*", System.IO.SearchOption.AllDirectories);

            return files;
        }

        public static void AutoFeature()                                    //废弃，不必查看
        {
            ISldWorks swApp = ConnectToSolidWorks();
            var swModel = (ModelDoc2)swApp.ActiveDoc;


        }

        public static void DifferenceToFeatures()                       //特征的测试函数，与主要功能无关
        {
            //本函数最大目的是弄清楚了Feature与subFeature的区别，与主要功能无关，不必查看

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

        public static void TrueFeature_NameOnly()                   //遍历solidworks中已打开的文件的所有草图
        {
            //本函数在之前的遍历特征函数基础上增加了对特征的甄别，使其只获取所有草图类（草图一般作为某些特征的subFeature）

            ISldWorks swApp = ConnectToSolidWorks();        //首先连接solidworks
            var swModel = (ModelDoc2)swApp.ActiveDoc;     //获取当前文件有关信息

            Feature swFeat = (Feature)swModel.FirstFeature();   //获取当前文件首个特征
            //int i = 0;
            while(swFeat != null)                                              //持续遍历特征
            {
                Feature SubFeature = swFeat.GetFirstSubFeature();       //获取当前特征的subFeature
                while(SubFeature != null)
                {
                    if(SubFeature.GetTypeName2() == "ProfileFeature")   //若该subFeature为草图
                    {
                        Sketch swSketch = SubFeature.GetSpecificFeature2();     //获取该草图作为实例swSketch
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

        public static void OpenAndClose(string path)                //遍历指定目录下所有文件的所有草图
        {
            //本质上为上述功能的整合

            Debug.Print(path);
            ISldWorks SwApp = ConnectToSolidWorks();
            //string partDefaultTemplate = SwApp.GetDocumentTemplate((int)swDocumentTypes_e.swDocPART, "", 0, 0, 0);
            SwApp.OpenDoc(path, (int)swDocumentTypes_e.swDocPART);

            TrueFeature_NameOnly();

            SwApp.CloseDoc(path);
        }

        public static void DirectDocument(string path, string license)
        {
            //setDocType();
            SwDmDocumentType dmDocType;

            SwDMClassFactory dmClassFact = new SwDMClassFactory();
            SwDMApplication4 dmDocMgr = (SwDMApplication4)dmClassFact.GetApplication(license);
            //SwDMDocument18 dmDoc = (SwDMDocument18)dmDocMgr.GetDocument(path, true)
            if(dmDocMgr != null)
            {
                Debug.Print("Yes!");
            }

        }

        public static void DirectDocument1(string sDocFileName, string license)
        {
            SwDMClassFactory swClassFact = default(SwDMClassFactory);
            SwDMApplication swDocMgr = default(SwDMApplication);
            SwDMDocument10 swDoc = default(SwDMDocument10);
            SwDMConfigurationMgr swCfgMgr = default(SwDMConfigurationMgr);
            string[] vCfgNameArr = null;
            SwDMConfiguration2 swCfg = default(SwDMConfiguration2);
            SwDmDocumentType nDocType = 0;
            SwDmPreviewError nError = 0;
            SwDmBodyError nBodyError = 0;
            SwDmDocumentOpenError nRetVal = 0;
            int i = 0;
            string sFileName = "multi-inter";

            if (sDocFileName.EndsWith("sldprt"))
            {
                nDocType = SwDmDocumentType.swDmDocumentPart;
            }
            else if (sDocFileName.EndsWith("sldasm"))
            {
                nDocType = SwDmDocumentType.swDmDocumentAssembly;
            }
            else if (sDocFileName.EndsWith("slddrw"))
            {
                nDocType = SwDmDocumentType.swDmDocumentDrawing;
            }
            else
            {
                // Not a SOLIDWORKS file
                nDocType = SwDmDocumentType.swDmDocumentUnknown;
                return;
            }

            if((nDocType != SwDmDocumentType.swDmDocumentDrawing))
            {
                SwDMComponent2 swDocMgr2;

                swClassFact = new SwDMClassFactory();
                swDocMgr = (SwDMApplication)swClassFact.GetApplication(license);
                swDocMgr2 = (SwDMComponent2)swClassFact.GetApplication(license);
                //swDoc = (SwDMDocument10)swDocMgr.GetDocument(sDocFileName, nDocType, true, out nRetVal);
                swDoc = (SwDMDocument10)swDocMgr.GetDocument(sDocFileName, nDocType, true, out nRetVal);
                
                

                swCfgMgr = swDoc.ConfigurationManager;

                
            }
        }
    }
}

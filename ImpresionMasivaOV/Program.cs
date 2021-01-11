using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace ImpresionMasivaOV
{
    class Program
    {
        public static SAPbobsCOM.Company oCompany = null;



        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);

                Conexion.Conectar_Aplicacion();
                oCompany = Conexion.oCompany;

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }


        static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.DataTable oDataTable = null;

            try
            {
                // ------------------------------------------------------------------------------------------------------------------------------------------------
                //   ESTOS EVENTO MANEJA LA CONDICION MODAL DE LAS PANTALLAS DONDE ReportType = "Modal"
                // ------------------------------------------------------------------------------------------------------------------------------------------------
                if (((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                            && (pVal.BeforeAction == true) && pVal.FormTypeEx == "ImpresionMasivaOV.Form1"))
                {
                    try
                    {
                        oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        if (pVal.ItemUID == "Grid" && pVal.ColUID == "N° Documento")
                        {
                            Form1.OpendocumenLink(pVal.Row);

                            //oForm.Freeze(true);
                            //string sDocnum = oGrid.DataTable.GetValue("N° Documento", pVal.Row).ToString().Trim();
                            //s = @"SELECT ""DocEntry"" FROM ""ORDR"" WHERE ""DocNum"" = {0}";
                            //s = String.Format(s, sDocnum);
                            //oRecordSet.DoQuery(s);
                            //TempDocNumLink = sDocnum;
                            //string docEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                            //oGrid.DataTable.SetValue("N° Documento", pVal.Row, docEntry);
                        }


                    }
                    catch (Exception) { }

                }

                if (((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) && (pVal.BeforeAction == false) && pVal.FormTypeEx == "ImpresionMasivaOV.Form1"))
                {
                    try
                    {
                        if (pVal.ItemUID == "Grid" && pVal.ColUID == "N° Documento")
                        {
                            Form1.CloseDocumentLink(pVal.Row);
                        }
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception) { }
        }

    }
}

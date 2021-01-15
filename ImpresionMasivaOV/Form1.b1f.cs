using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using Newtonsoft.Json;
using Stimulsoft.Report;
using Stimulsoft.Base;
using System.IO;
using System.Xml.Serialization;
using System.Data;

namespace ImpresionMasivaOV
{
    [FormAttribute("ImpresionMasivaOV.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.CheckBox CheckBoxImpresion;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditTextDocDesde;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditTextDocHasta;
        private SAPbouiCOM.ComboBox CBoxTpoEstado;
        private SAPbouiCOM.ComboBox CBoxtpoRuta;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditTextFechaDesde;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.EditText EditTextFechaHasta;
        private static SAPbouiCOM.Form oForm = null;
        private static SAPbouiCOM.DBDataSource oDBDataSource = null;
        private string s;
        private static SAPbobsCOM.Recordset oRecordSet = null;
        private SAPbouiCOM.DataTable oDataTable;
        private static SAPbobsCOM.Company oCompany = Program.oCompany;
        private  SAPbouiCOM.Grid oGrid;
        private static SAPbouiCOM.Grid oGridstatic;
        private SAPbouiCOM.Button ButtonBuscar;
        private SAPbouiCOM.Button ButtonImprimir;
        private SAPbouiCOM.Button ButtonCancelar;
        private SAPbouiCOM.CheckBox CheckBox1;
        private SAPbouiCOM.StaticText StaticText0;
        public Log log;
        private static string TempDocNumLink;


        public Form1()
        {
            log = new Log();
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.CheckBoxImpresion = ((SAPbouiCOM.CheckBox)(this.GetItem("chkImp").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditTextDocDesde = ((SAPbouiCOM.EditText)(this.GetItem("docDesde").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.EditTextDocHasta = ((SAPbouiCOM.EditText)(this.GetItem("docHasta").Specific));
            this.CBoxTpoEstado = ((SAPbouiCOM.ComboBox)(this.GetItem("tpoEstado").Specific));
            this.CBoxtpoRuta = ((SAPbouiCOM.ComboBox)(this.GetItem("tpoRuta").Specific));
            this.CBoxtpoRuta.ClickBefore += new SAPbouiCOM._IComboBoxEvents_ClickBeforeEventHandler(this.CBoxtpoRuta_ClickBefore);
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            this.EditTextFechaDesde = ((SAPbouiCOM.EditText)(this.GetItem("FechaD").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.EditTextFechaHasta = ((SAPbouiCOM.EditText)(this.GetItem("FechaH").Specific));
            this.oGrid = ((SAPbouiCOM.Grid)(this.GetItem("Grid").Specific));
            this.oGrid.LinkPressedBefore += new SAPbouiCOM._IGridEvents_LinkPressedBeforeEventHandler(this.oGrid_LinkPressedBefore);
            this.oGrid.LinkPressedAfter += new SAPbouiCOM._IGridEvents_LinkPressedAfterEventHandler(this.oGrid_LinkPressedAfter);
            this.oGrid.DoubleClickBefore += new SAPbouiCOM._IGridEvents_DoubleClickBeforeEventHandler(this.oGrid_DoubleClickBefore);
            this.oGrid.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.oGrid_DoubleClickAfter);
            this.ButtonBuscar = ((SAPbouiCOM.Button)(this.GetItem("bBuscar").Specific));
            this.ButtonBuscar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.ButtonBuscar_ClickBefore);
            this.ButtonImprimir = ((SAPbouiCOM.Button)(this.GetItem("bImprimir").Specific));
            this.ButtonImprimir.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.ButtonImprimir_ClickAfter);
            this.ButtonImprimir.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.ButtonImprimir_ClickBefore);
            this.ButtonCancelar = ((SAPbouiCOM.Button)(this.GetItem("bCancelar").Specific));
            this.ButtonCancelar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.ButtonCancelar_ClickAfter);
            this.ButtonCancelar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.ButtonCancelar_ClickBefore);
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("chk_sel").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.RightClickAfter += new SAPbouiCOM.Framework.FormBase.RightClickAfterHandler(this.Form_RightClickAfter);
            this.ActivateAfter += new ActivateAfterHandler(this.Form_ActivateAfter);

        }



        private void OnCustomInitialize()
        {
            try
            {

                oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);

                oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oGridstatic = (SAPbouiCOM.Grid)oForm.Items.Item("Grid").Specific;
                
                //((SAPbouiCOM.Grid)(this.GetItem("Grid").Specific));
                

                //oDBDataSource = oForm.DataSources.DBDataSources.Item("ORDR");
                //SAPbouiCOM.DBDataSource oDBDSORDR = oForm.DataSources.DBDataSources.Item("ORDR");

                oForm.DataSources.UserDataSources.Add("FechaD", SAPbouiCOM.BoDataType.dt_DATE, 10);
                EditTextFechaDesde.DataBind.SetBound(true, "", "FechaD");

                oForm.DataSources.UserDataSources.Add("FechaH", SAPbouiCOM.BoDataType.dt_DATE, 10);
                EditTextFechaHasta.DataBind.SetBound(true, "", "FechaH");


                oForm.DataSources.UserDataSources.Add("tpoEstado", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                CBoxTpoEstado.DataBind.SetBound(true, "", "tpoEstado");
                CBoxTpoEstado.ValidValues.Add("O", "Abierto");
                CBoxTpoEstado.ValidValues.Add("C", "Cerrado");
                CBoxTpoEstado.ValidValues.Add("T", "Todos");


                oForm.DataSources.UserDataSources.Add("tpoRuta", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                CBoxtpoRuta.DataBind.SetBound(true, "", "tpoRuta");

                //s = @"SELECT DISTINCT T0.""U_RUTA"" ""Code"" , T0.""U_RUTA"" ""Name"" FROM ""ORDR"" T0 
                //     WHERE T0.""U_RUTA"" IS NOT NULL";

                //oRecordSet.DoQuery(s);
                //oRecordSet.MoveFirst();

                //while (!oRecordSet.EoF)
                //{
                //    string code = (String)(oRecordSet.Fields.Item("Code").Value);
                //    string name = (String)(oRecordSet.Fields.Item("Name").Value);
                //    CBoxtpoRuta.ValidValues.Add(code, name);
                //    oRecordSet.MoveNext();
                //}
                //int a = oRecordSet.RecordCount;

                oForm.DataSources.UserDataSources.Add("chkImp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                CheckBoxImpresion.DataBind.SetBound(true, "", "chkImp");
                CheckBoxImpresion.ValOn = "Y";
                CheckBoxImpresion.ValOff = "N";

                oForm.DataSources.UserDataSources.Add("chk_sel", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                CheckBox1.DataBind.SetBound(true, "", "chk_sel");
                CheckBox1.ValOn = "Y";
                CheckBox1.ValOff = "N";

                oForm.DataSources.UserDataSources.Add("docDesde", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10);
                EditTextDocDesde.DataBind.SetBound(true, "", "docDesde");

                oForm.DataSources.UserDataSources.Add("docHasta", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10);
                EditTextDocHasta.DataBind.SetBound(true, "", "docHasta");


                oForm.Items.Item("FechaD").Click(); //para asignar ese campo como primero por llenar

                oDataTable = oForm.DataSources.DataTables.Add("dt");
                oGrid.DataTable = oDataTable;
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                buscarDatos();

            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }



        }

        private void cargarRuta()
        {
            String fechaD, fechaH;

            try
            {
                oForm.Freeze(true);

                if (CBoxtpoRuta.ValidValues.Count > 0)
                {
                    for (int i = CBoxtpoRuta.ValidValues.Count - 1; i >= 0; i--)
                        CBoxtpoRuta.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                fechaD = ((System.String)((SAPbouiCOM.EditText)oForm.Items.Item("FechaD").Specific).Value);
                fechaH = ((System.String)((SAPbouiCOM.EditText)oForm.Items.Item("FechaH").Specific).Value);

                s = @"SELECT T0.""U_RUTA"" ""Code"" , T0.""U_RUTA"" ""Name""
                        FROM ""ORDR"" T0
                        WHERE T0.""U_RUTA"" IS NOT NULL
                        AND T0.""DocDueDate"" BETWEEN '{0}' AND  '{1}'
                        GROUP BY T0.""U_RUTA""
                        ORDER BY T0.""U_RUTA""";
                s = String.Format(s, fechaD, fechaH);

                oRecordSet.DoQuery(s);
                oRecordSet.MoveFirst();

                int a = oRecordSet.RecordCount;

                while (!oRecordSet.EoF)
                {
                    string code = (String)(oRecordSet.Fields.Item("Code").Value);
                    string name = (String)(oRecordSet.Fields.Item("Name").Value);
                    CBoxtpoRuta.ValidValues.Add(code, name);
                    oRecordSet.MoveNext();

                    ;
                }
                if (CBoxtpoRuta.ValidValues.Count > 0)
                {
                    CBoxtpoRuta.ValidValues.Add("Todos", "Todos");
                    //oRecordSet.MoveNext();
                }

            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        private SAPbobsCOM.Recordset buscarDatosDetalle(int docEntry)
        {
            try
            {

                s = @"SELECT
                    T1.""DocEntry"",
                    T1.""LineNum"" + 1 ""LineNum"",
                    T1.""ItemCode"",
                    T1.""Dscription"",
                    T1.""Quantity"",
                    T1.""Price"",
                    T1.""LineTotal""
                    FROM ""RDR1"" T1
                    WHERE
                    T1.""DocEntry"" = {0}";
                s = String.Format(s, docEntry);

                oRecordSet.DoQuery(s);
                if (oRecordSet.RecordCount > 0)
                    return oRecordSet;
                else
                    return null;

            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
                return null;
            }
        }
        private void buscarDatos()
        {
            String fechaD, fechaH, Ruta, Estado, DocnumDesde, DocnumHasta;

            try
            {
                oForm.Freeze(true);
                fechaD = ((System.String)((SAPbouiCOM.EditText)oForm.Items.Item("FechaD").Specific).Value);
                fechaH = ((System.String)((SAPbouiCOM.EditText)oForm.Items.Item("FechaH").Specific).Value);
                Ruta = ((System.String)((SAPbouiCOM.ComboBox)oForm.Items.Item("tpoRuta").Specific).Value);
                Estado = ((System.String)((SAPbouiCOM.ComboBox)oForm.Items.Item("tpoEstado").Specific).Value);
                DocnumDesde = ((System.String)((SAPbouiCOM.EditText)oForm.Items.Item("docDesde").Specific).Value);
                DocnumHasta = ((System.String)((SAPbouiCOM.EditText)oForm.Items.Item("docHasta").Specific).Value);

                s = @"SELECT 
                    'N' ""Selec."",
                    T0.""DocNum"" ""N° Documento"" ,
                    T0.""U_CorrelativoERP"" ""Pre-Orden"",
                    T0.""U_RUTA"" ""Ruta"", 
                    T1.""SlpName"" ""Vendedor"",
                    T0.""DocDueDate"" ""Fecha Despacho"", 
                    T0.""CardCode"" ""Código Cliente/Proveedor"", 
                    T0.""CardName"" ""Nombre Cliente/Proveedor"", 
                    T0.""DocTotal"" ""Total Documento"",
                    T0.""DocEntry"" ""DocEntry"",
                    T12.""StreetS"" ""DirDespacho"",
                    T12.""CityS"" ""Ciudad"",
                    T12.""CountyS"" ""Comuna"",
                    T0.""U_CorrelativoERP"" ""CorrelativoERP""
                    FROM ""ORDR"" T0
                    JOIN ""RDR12"" T12 ON T12.""DocEntry"" = T0.""DocEntry""
                    LEFT JOIN ""OSLP"" T1 ON T0.""SlpCode"" = T1.""SlpCode""
                    WHERE 1=1
                    AND T0.""DocDueDate"" BETWEEN '{0}' AND  '{1}'";

                s = String.Format(s, fechaD, fechaH);

                if (Ruta != "Todos")
                {
                    s = s + @"AND T0.""U_RUTA"" = '{0}'";
                    s = String.Format(s, Ruta);
                }

                if (Estado != "T")
                {
                    s = s + @"AND T0.""DocStatus"" = '{0}'";
                    s = String.Format(s, Estado);
                }

                if (((SAPbouiCOM.CheckBox)oForm.Items.Item("chkImp").Specific).Checked)
                {
                    s = s + @"AND T0.""Printed"" = '{0}'";
                    s = String.Format(s, 'N');
                }

                if (DocnumDesde != "" && DocnumHasta != "")
                {
                    s = s + @"AND T0.""DocNum"" BETWEEN '{0}' AND '{1}'";
                    s = String.Format(s, DocnumDesde, DocnumHasta);
                }


                s = s + @"ORDER BY T0.""DocDueDate"" , T0.""U_RUTA"", T0.""CardCode"" , T0.""DocTotal"" ";

                oDataTable.ExecuteQuery(s);

                oGrid.Columns.Item("Selec.").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                var ocheckColumns = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Selec."));
                var ocheckColumn = (SAPbouiCOM.CheckBoxColumn)(ocheckColumns);
                ocheckColumn.Editable = true;

                oGrid.Columns.Item("N° Documento").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                var oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("N° Documento"));
                var oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;
                oEditColumn.LinkedObjectType = "17";

                oGrid.Columns.Item("Pre-Orden").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Pre-Orden"));
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;

                oGrid.Columns.Item("Ruta").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Ruta"));
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;

                oGrid.Columns.Item("Vendedor").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Vendedor"));
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;

                oGrid.Columns.Item("Fecha Despacho").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Fecha Despacho"));
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;

                oGrid.Columns.Item("Código Cliente/Proveedor").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Código Cliente/Proveedor"));
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;

                oGrid.Columns.Item("Nombre Cliente/Proveedor").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Nombre Cliente/Proveedor"));
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;

                oGrid.Columns.Item("Total Documento").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Total Documento"));
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Editable = false;

                oGrid.Columns.Item("DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("DocEntry"));
                oColumn.Visible = false;
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Visible = false;

                oGrid.Columns.Item("DirDespacho").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("DirDespacho"));
                oColumn.Visible = false;
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Visible = false;

                oGrid.Columns.Item("Comuna").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Comuna"));
                oColumn.Visible = false;
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Visible = false;

                oGrid.Columns.Item("CorrelativoERP").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("CorrelativoERP"));
                oColumn.Visible = false;
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Visible = false;

                oGrid.Columns.Item("Ciudad").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oColumn = (SAPbouiCOM.GridColumn)(oGrid.Columns.Item("Ciudad"));
                oColumn.Visible = false;
                oEditColumn = (SAPbouiCOM.EditTextColumn)(oColumn);
                oEditColumn.Visible = false;

                SAPbouiCOM.RowHeaders oHeader = null;
                for (int i = 0; i <= oGrid.DataTable.Rows.Count - 1; i++)
                {
                    //Enumera Fila
                    oHeader = oGrid.RowHeaders;
                    oHeader.SetText(i, Convert.ToString(i + 1));
                }
            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }
            finally
            {
                oForm.Freeze(false);
            }
        }



        private void ButtonBuscar_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //validar 

            buscarDatos();
            BubbleEvent = true;
            // throw new System.NotImplementedException();

        }

        private void ButtonImprimir_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();

        }

        private void oGrid_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {


        }

        private void generarPDF()
        {
            try
            {
                SAPbobsCOM.Recordset oRecordSetDetalle = null;
                oGrid = ((SAPbouiCOM.Grid)oForm.Items.Item("Grid").Specific);
                List<OV> listOV = new List<OV>();
                string xml = oGrid.DataTable.GetAsXML();
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    var ocheckColumn = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Selec.");
                    if (ocheckColumn.IsChecked(i))
                    {
                        OV ov = new OV();
                        ov.Encabezado = new oEncabezado();
                        ov.Encabezado.ruta = ((System.String)oGrid.DataTable.GetValue("Ruta", i)).Trim();
                        ov.Encabezado.cardName = ((System.String)oGrid.DataTable.GetValue("Nombre Cliente/Proveedor", i)).Trim();
                        ov.Encabezado.cardCode = ((System.String)oGrid.DataTable.GetValue("Código Cliente/Proveedor", i)).Trim();
                        ov.Encabezado.docentry = ((int)oGrid.DataTable.GetValue("DocEntry", i));
                        ov.Encabezado.docnum = ((int)oGrid.DataTable.GetValue("N° Documento", i));
                        ov.Encabezado.dirDespacho = ((System.String)oGrid.DataTable.GetValue("DirDespacho", i)).Trim();
                        ov.Encabezado.comuna = ((System.String)oGrid.DataTable.GetValue("Comuna", i)).Trim();
                        ov.Encabezado.ciudad = ((System.String)oGrid.DataTable.GetValue("Ciudad", i)).Trim();
                        ov.Encabezado.docTotal = ((System.Double)oGrid.DataTable.GetValue("Total Documento", i));
                        ov.Encabezado.fechaDespacho = ((System.DateTime)oGrid.DataTable.GetValue("Fecha Despacho", i));
                        ov.Encabezado.CorrelativoERP = ((System.String)oGrid.DataTable.GetValue("CorrelativoERP", i)).Trim();

                        oRecordSetDetalle = buscarDatosDetalle(ov.Encabezado.docentry);
                        if (oRecordSetDetalle != null)
                        {
                            ov.Detalle = new List<oDetalle>();
                            while (!oRecordSetDetalle.EoF)
                            {
                                oDetalle detalle = new oDetalle();
                                detalle.codigo = ((System.String)oRecordSetDetalle.Fields.Item("ItemCode").Value);
                                detalle.descripcion = ((System.String)oRecordSetDetalle.Fields.Item("Dscription").Value);
                                detalle.cantiad = ((System.Double)oRecordSetDetalle.Fields.Item("Quantity").Value);
                                detalle.docentry = ((System.Int32)oRecordSetDetalle.Fields.Item("DocEntry").Value);
                                detalle.lineNum = ((System.Int32)oRecordSetDetalle.Fields.Item("LineNum").Value);
                                detalle.precio = ((System.Double)oRecordSetDetalle.Fields.Item("Price").Value);
                                detalle.total = ((System.Double)oRecordSetDetalle.Fields.Item("LineTotal").Value);
                                ov.Detalle.Add(detalle);
                                oRecordSetDetalle.MoveNext();
                            }
                        }
                        listOV.Add(ov);

                    }
                }

                if (listOV.Count > 0)
                {
                    //string jsonData = JsonConvert.SerializeObject(listOV);
                    string xmlString = ConvertObjectToXMLString(listOV);
                    GetReportStimulsoft(xmlString);
                }

            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }

        }

        static string ConvertObjectToXMLString(object classObject)
        {
            string xmlString = null;
            XmlSerializer xmlSerializer = new XmlSerializer(classObject.GetType());
            using (MemoryStream memoryStream = new MemoryStream())
            {
                xmlSerializer.Serialize(memoryStream, classObject);
                memoryStream.Position = 0;
                xmlString = new StreamReader(memoryStream).ReadToEnd();
            }
            return xmlString;
        }


        public DataSet ConvertXMLToDataSet(string xmlData)
        {
            StringReader stream = null;
            XmlTextReader reader = null;
            try
            {
                DataSet xmlDS = new DataSet();
                stream = new StringReader(xmlData);
                // Load the XmlTextReader from the stream
                reader = new XmlTextReader(stream);
                xmlDS.ReadXml(reader);
                return xmlDS;
            }
            catch (Exception e)
            {
                log.AddLog(e.Message + " ** Trace: ");
                return null;
            }
            finally
            {
                if (reader != null) reader.Close();
            }
        }


        private void GetReportStimulsoft(string xmltext)
        {

            try
            {

                global::Stimulsoft.Base.StiLicense.Key = "6vJhGtLLLz2GNviWmUTrhSqnOItdDwjBylQzQcAOiHn6T1QyRLNg9ob5/AoMlKpfD06YlnbaK+apLpkPGy58/hwEVP" +
                                "JFLu2ahVXhoRuQ6rqqr2dmiE1sVk+HoFVWz15idNVym7+T9lWeQUbd8FI/gJCJVd9zEPTA3yfhJpZx1s2ZXumj8n0P" +
                                "FAahfNUT8qlOCjmeZ2admzNVdRlTcH/uN3Ms51HIix2g7C0cuupRUJOYBM36vuEOSXp1B07rV6NwU0iACQHiUQ/Y4c" +
                                "Gx2SVSiZdVGKY4hVgfWDeHCTr5MaqXWo6p6EOSVB0bM3Y421Tv2qitJ3Utj/zcYDVbW5nSwhahuygT3ZCY5iftNvzw" +
                                "gwIEjS2LnGME3QghFEWnC04Vld/zxSQyxGcMyK7/03VkqfHlBN8jIVHEjFT0YQUhPAbiC2pfFKa6MIgJqvXTJDNQgn" +
                                "6y8c9RwfwPdC6PJjL/9c0kEpaG198A2R0mVZNzjvXHpG/mEUIeWN2zmWJMJNm5fgySzlV9BLUwKlM1jpv4rQcf5MR/" +
                                "/ZONmx6qqmjYcSASNmW/ICM72fwSsJE7F7chh1Q0VMkOe6sriXsdhkqC3lV5yTifwCK3JYM9i08XF1HXDMeNF6/tss" +
                                "wdMaaCVQDGJJp3stA8KlSyAeLFvRo5uMFl/5vvCuK3lV275SRgStTvS4uAu2yWIkUMnxMey6mZ";



                String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
                sPath = sPath + @"/Report/OVReportMasivo.mrt";

                XmlDocument xml = new XmlDocument();
                xml.LoadXml(xmltext);

                StiReport rep = StiReport.CreateNewReport();

                MemoryStream stream = new MemoryStream();

                rep.Load(sPath);
                rep.Dictionary.Databases.Clear();
                var ds = ConvertXMLToDataSet(xml.InnerXml);
                rep.Dictionary.DataSources.Clear(); //refrescar reporte con campo nuevo 
                rep.RegData(ds);
                rep.Dictionary.Synchronize();     //refrescar reporte con campo nuevo 
                rep.Render();


                DateTime now = DateTime.Now;
                string pdfName = now.ToString("yyyyMMddTHHmmssZ") + "OVReportMasivo.pdf";


                //rep.ExportDocument(StiExportFormat.Pdf, stream); //intento de abrirlo pdf en memoria, pero no funciono
                
                rep.ExportDocument(StiExportFormat.Pdf, "C:\\Windows\\Temp\\" + pdfName);
                sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
                //string filename = "OVReportMasivo.pdf";
                System.Diagnostics.Process.Start("C:\\Windows\\Temp\\" + pdfName);

            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }
            finally
            {

            }




        }

        private void oGrid_DoubleClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (pVal.Row == -1 && pVal.ColUID == "Selec.")
            {
                if (((SAPbouiCOM.CheckBox)oForm.Items.Item("chk_sel").Specific).Checked)
                    Seleccion(false);
                else
                    Seleccion(true);
            }

            BubbleEvent = true;
            //throw new System.NotImplementedException();

        }

        private void Seleccion(Boolean bvalor)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Grid").Specific;
                SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("dt");


                for (Int32 i = 0; i <= DT_GRID.Rows.Count - 1; i++)
                {
                    if (oGrid.CommonSetting.GetCellEditable(i + 1, 1))
                        if (bvalor)
                            DT_GRID.SetValue("Selec.", i, "Y");
                        else
                            DT_GRID.SetValue("Selec.", i, "N");
                }
                CheckBox1.Item.Visible = true; //debe estar "visible" para camibar el checked

                if (((SAPbouiCOM.CheckBox)oForm.Items.Item("chk_sel").Specific).Checked)
                    ((SAPbouiCOM.CheckBox)oForm.Items.Item("chk_sel").Specific).Checked = false;
                else
                    ((SAPbouiCOM.CheckBox)oForm.Items.Item("chk_sel").Specific).Checked = true;
                CheckBox1.Item.Visible = false;
            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }
            oForm.Freeze(false);
        }



        private void CBoxtpoRuta_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            try
            {
                cargarRuta();
            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }
            finally
            {
                BubbleEvent = true;
            }



        }

        private void ButtonCancelar_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();

        }

        private void ButtonCancelar_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                oForm.Close();
            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                log.AddLog(e.Message + " ** Trace: ");
            }
            finally
            {

            }


        }

        private void ButtonImprimir_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            generarPDF();


        }

        private void Form_RightClickAfter(ref SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            //throw new System.NotImplementedException();

        }

        private void oGrid_LinkPressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //if (pVal.ItemUID == "Grid" && pVal.ColUID == "N° Documento")
            //{
            //    oGrid.DataTable.SetValue("N° Documento", pVal.Row, TempDocNumLink);
            //}

            //oForm.Freeze(false);
        }

        public static void CloseDocumentLink(int row)
        {

           oGridstatic.DataTable.SetValue("N° Documento", row, TempDocNumLink);

            oForm.Freeze(false);
        }


        public static void OpendocumenLink (int row )
        {
            string s;
            string sDocnum;
            oForm.Freeze(true);
            sDocnum = oGridstatic.DataTable.GetValue("N° Documento", row).ToString().Trim();
            s = @"SELECT ""DocEntry"" FROM ""ORDR"" WHERE ""DocNum"" = {0}";
            s = String.Format(s, sDocnum);
            oRecordSet.DoQuery(s);
            TempDocNumLink = sDocnum;
            string docEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
            oGridstatic.DataTable.SetValue("N° Documento", row, docEntry);
        }
        private void oGrid_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //try
            //{
            //    if (pVal.ItemUID == "Grid" && pVal.ColUID == "N° Documento")
            //    {
            //        oForm.Freeze(true);
            //        string sDocnum = oGrid.DataTable.GetValue("N° Documento", pVal.Row).ToString().Trim();
            //        s = @"SELECT ""DocEntry"" FROM ""ORDR"" WHERE ""DocNum"" = {0}";
            //        s = String.Format(s, sDocnum);
            //        oRecordSet.DoQuery(s);
            //        TempDocNumLink = sDocnum;
            //        string docEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
            //        oGrid.DataTable.SetValue("N° Documento", pVal.Row, docEntry);
            //    }

            //}
            //catch (Exception e)
            //{
            //    //oForm.Freeze(false);
            //}

            BubbleEvent = true;
        }

        private void Form_ActivateAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (1 ==1 )
            {

            }

            //throw new System.NotImplementedException();

        }
    }
}

using AddOnRevalorizacion.Comunes;
using AddOnRevalorizacion.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRevalorizacion.Class
{
    class Frm_143
    {
        public const string csFormType = "143";
        public void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            int resp = 0;
            string s_Message = "";
            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                            if (pVal.FormTypeEx == "143")
                            {
                                DubujarBoton(FormUID);
                            }
                            
                            break;
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.FormTypeEx == "143")
                            {
                                if (pVal.ItemUID== "btnProceso")
                                {
                                    s_Message = "¿Esta seguro de hacer la revalozarización de inventario? \n";
                                    s_Message += "Se generará una entrada/salida de inventario en función de la entrada de mercancias, \n";
                                    s_Message += "luego se generará una revalorización de inventario \n";

                                    resp = Conexion.Conexion_SBO.m_SBO_Appl.MessageBox(s_Message, 2, "Procesar", "Cancelar");

                                    if(resp == 1)
                                        GenerarInventyDoc(FormUID);
                                }
                               
                            }
                            break;
                    }
                }
                
            }
            catch (Exception ex)
            {
                ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
            }
        }

        public void m_SBO_Appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                validateMenu();
            }
            catch (Exception ex)
            {
                Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Properties.Resources.NombreAddon + " Error: Frm_134.cs > m_SBO_Appl_MenuEvent():"
                    + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void validateMenu()
        {
            string s_FormTypeEx = "";
            string s_FormUID = "";
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oItem = null;

            try
            {
                s_FormTypeEx = Conexion.Conexion_SBO.m_SBO_Appl.Forms.ActiveForm.TypeEx.ToString();
                if (s_FormTypeEx == "143")
                {
                    s_FormUID = Conexion.Conexion_SBO.m_SBO_Appl.Forms.ActiveForm.UniqueID.ToString();
                    oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item(s_FormUID);
                    oItem = oForm.Items.Item("btnProceso");

                    switch (oForm.Mode)
                    {
                        
                        case SAPbouiCOM.BoFormMode.fm_EDIT_MODE:
                        case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                            oItem.Enabled = true;
                            break;
                        case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                        case SAPbouiCOM.BoFormMode.fm_OK_MODE:
                        case SAPbouiCOM.BoFormMode.fm_PRINT_MODE:
                        case SAPbouiCOM.BoFormMode.fm_VIEW_MODE:
                        case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                            oItem.Enabled = false;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Properties.Resources.NombreAddon + " Error: Frm_143.cs > ValidarMenu(): "
                    + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        private void DubujarBoton(string s_FormUID)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item(s_FormUID);

                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Button oButton = null;

                oItem = oForm.Items.Add("btnProceso", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oButton = oItem.Specific;
                oButton.Caption = "Revalorizar";

                
                oItem.Top = 100;
                oItem.Left = (oItem.Width ) ;

                switch (oForm.Mode)
                {
                    
                    case SAPbouiCOM.BoFormMode.fm_EDIT_MODE:
                    case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                        oItem.Enabled = true;
                        break;
                    case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                    case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                    case SAPbouiCOM.BoFormMode.fm_OK_MODE:
                    case SAPbouiCOM.BoFormMode.fm_PRINT_MODE:
                    case SAPbouiCOM.BoFormMode.fm_VIEW_MODE:
                        oItem.Enabled = false;
                        break;
                }

            }
            catch (Exception ex)
            {
                ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
            }
        }


        private void GenerarInventyDoc(string s_FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbobsCOM.Recordset oRS = null;
            


            try
            {
                oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item(s_FormUID);
                oRS = (SAPbobsCOM.Recordset)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                
                oItem = oForm.Items.Item("8");
                oEditText = oItem.Specific;

                string query = "SELECT IFNULL(B.\"U_SMF_RINV\",'0') AS \"Revalorizacion\", IFNULL(B.\"U_SMF_SINV\",'0') AS \"Salida\", IFNULL(B.\"U_SMF_EINV\",'0') AS \"Entrada\", B.\"LineNum\", B.\"ItemCode\", A.\"DocEntry\", A.\"DocCur\",A.\"DocRate\", CASE WHEN A.\"DocCur\"='USD' THEN B.\"TotalFrgn\" ELSE B.\"LineTotal\" END AS \"TotalLine\", ";
                query = query + " A.\"DocDate\" AS \"DocDate\", A.\"TaxDate\" AS \"TaxDate\", ";
                query = query + " B.\"Quantity\", IFNULL(B.\"U_SMF_CREAL\",0) AS \"QuantityReal\", (SELECT \"U_SMF_CCSA\" FROM \"@SMF_REVA\" WHERE \"Code\" = '001') AS \"AcctCodeS\", (SELECT \"U_SMF_CCEN\" FROM \"@SMF_REVA\" WHERE \"Code\" = '001') AS \"AcctCodeE\", B.\"WhsCode\", ";
                query = query + " (SELECT MAX(\"BatchNum\") FROM OIBT WHERE \"BaseType\"='20' AND \"BaseEntry\"=A.\"DocEntry\" AND \"BaseLinNum\"= B.\"LineNum\" ) AS \"BatchNum\", ";
                query = query + " ( select MAX(T2.\"AbsEntry\") from OITW T0 inner join OITM T1 on T0.\"ItemCode\" = T1.\"ItemCode\" inner join OIBQ T3 on T0.\"ItemCode\" = T3.\"ItemCode\" and T0.\"WhsCode\" = T3.\"WhsCode\" ";
                query = query + " inner join OBIN T2 on T2.\"AbsEntry\" = T3.\"BinAbs\" WHERE T0.\"ItemCode\" = B.\"ItemCode\" AND T0.\"WhsCode\" = B.\"WhsCode\" ) AS \"Location\", ";
                query = query + " B.\"OcrCode\", B.\"OcrCode2\", B.\"OcrCode3\", B.\"OcrCode4\", B.\"OcrCode5\", IFNULL((SELECT MAX(\"DocRate\") FROM OPCH WHERE \"DocEntry\"=B.\"BaseEntry\"),1) AS \"tc_base\" ";
                query = query + " FROM OPDN A INNER JOIN PDN1 B ON A.\"DocEntry\"=B.\"DocEntry\" WHERE A.\"DocNum\" = '" + oEditText.Value + "' ORDER BY 4 ";

                oRS.DoQuery(query);

                if (oRS.RecordCount > 0)
                {

                    oRS.MoveFirst();
                    List<string> lista = new List<string>();

                    while (!oRS.EoF)
                    {
                        Receipt receipt = new Receipt();

                        receipt.DocDate = Convert.ToDateTime(oRS.Fields.Item("DocDate").Value);
                        receipt.TaxDate = Convert.ToDateTime(oRS.Fields.Item("TaxDate").Value);
                        receipt.DocCur = oRS.Fields.Item("DocCur").Value;
                        receipt.DocEntry = oRS.Fields.Item("DocEntry").Value;
                        receipt.LineNum = oRS.Fields.Item("LineNum").Value;
                        receipt.Itemcode = oRS.Fields.Item("ItemCode").Value;
                        receipt.Quantity = oRS.Fields.Item("Quantity").Value;
                        receipt.QuantityReal = oRS.Fields.Item("QuantityReal").Value;
                        receipt.TotalLine = oRS.Fields.Item("TotalLine").Value;
                        receipt.AccountCodeSalida = oRS.Fields.Item("AcctCodeS").Value;
                        receipt.AccountCodeEntrada = oRS.Fields.Item("AcctCodeE").Value;
                        receipt.BatchNum = oRS.Fields.Item("BatchNum").Value;
                        receipt.Location = oRS.Fields.Item("Location").Value == null ? 0: oRS.Fields.Item("Location").Value;
                        receipt.WarehouseCode = oRS.Fields.Item("WhsCode").Value; 
                        receipt.CostingCode = oRS.Fields.Item("OcrCode").Value; 
                        receipt.CostingCode2 = oRS.Fields.Item("OcrCode2").Value;
                        receipt.CostingCode3 = oRS.Fields.Item("OcrCode3").Value;
                        receipt.CostingCode4 = oRS.Fields.Item("OcrCode4").Value;
                        receipt.CostingCode5 = oRS.Fields.Item("OcrCode5").Value;
                        receipt.Revalorizacion = oRS.Fields.Item("Revalorizacion").Value;
                        receipt.Salida = oRS.Fields.Item("Salida").Value;
                        receipt.Entrada = oRS.Fields.Item("Entrada").Value;
                        receipt.TcBase = oRS.Fields.Item("tc_base").Value;

                        lista.Add("Linea (" + (receipt.LineNum + 1) + ") :");


                        if (receipt.Salida == 0 && receipt.Entrada == 0 && receipt.Revalorizacion == 0)
                        {
                            
                            var salida_entrada = GenerarInventoryDocument(receipt);
                            lista.Add(" - " + salida_entrada.Item2.ToString());
                            if (salida_entrada.Item1)
                            {
                                var revalorizacionInventario = GenerarInventoryRevaluation(receipt);
                                lista.Add(" - " + revalorizacionInventario.Item2.ToString());

                            }
                            


                        }
                        else if ( (receipt.Salida == 0 || receipt.Entrada == 0) && receipt.Revalorizacion == 0)
                        {
                            var revalorizacionInventario = GenerarInventoryRevaluation(receipt);
                            lista.Add(" - La salida/entrada de inv, ya fue realizado en otro momento; solo faltó la revalorizacion");
                            lista.Add(" - " + revalorizacionInventario.Item2.ToString());
                        }
                        else
                        {
                            lista.Add(" - Ya se realizó la revalorización de inventario anteriormente");
                           // Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText("Ya se realizó la revalorización para la linea (" + (receipt.LineNum+1) + ") ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }


                        oRS.MoveNext();
                    }

                    string aaaa = "";
                    foreach (string aaa in lista)
                    {
                        aaaa = aaaa + aaa + "\n";
                    }

                    Conexion.Conexion_SBO.m_SBO_Appl.MessageBox(aaaa, 1, "", "", "");
                }

            }
            catch (Exception ex)
            {
                ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
            }
        }


        private Tuple<bool, string> GenerarInventoryDocument(Receipt receipt)
        {
            string sErrMsg = "";
            bool result = false;
            SAPbobsCOM.Documents oDocument = null;
            bool esSalida = false;

            try
            {
                

                if (receipt.Quantity != receipt.QuantityReal)
                {
                    string account_final = "";
                    if (receipt.Quantity > receipt.QuantityReal)
                    {
                        oDocument = Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                        account_final = receipt.AccountCodeSalida;
                        esSalida = true;
                    }
                    else if (receipt.Quantity < receipt.QuantityReal )
                    {
                        oDocument = Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                        account_final = receipt.AccountCodeEntrada;
                        esSalida = false;
                    }

                    oDocument.DocDate = receipt.DocDate;
                    oDocument.TaxDate = receipt.TaxDate;
                    //oDocument.DocCurrency = receipt.DocCur;


                    oDocument.Lines.Quantity = Math.Abs(receipt.Quantity - receipt.QuantityReal);
                    oDocument.Lines.AccountCode = account_final;
                    oDocument.Lines.ItemCode = receipt.Itemcode;

                    var BatchNum = receipt.BatchNum;
                    var WhsCode = receipt.WarehouseCode;
                    var sadsdfasd = receipt.Location.ToString();

                    if (!String.IsNullOrEmpty(BatchNum))
                    {
                        oDocument.Lines.BatchNumbers.BatchNumber = BatchNum;
                        oDocument.Lines.BatchNumbers.Quantity = Math.Abs(receipt.Quantity - receipt.QuantityReal);                       
                        oDocument.Lines.BatchNumbers.Add();

                        if (receipt.Location == 0)
                        {
                            oDocument.Lines.BinAllocations.BaseLineNumber = 0;
                            oDocument.Lines.BinAllocations.BinAbsEntry = receipt.Location;
                            oDocument.Lines.BinAllocations.Quantity = Math.Abs(receipt.Quantity - receipt.QuantityReal); 
                            oDocument.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
                        }
                            
                    }

                   


                    if (!String.IsNullOrEmpty(WhsCode))
                        oDocument.Lines.WarehouseCode = WhsCode;

                    var OcrCode = receipt.CostingCode;
                    if (!String.IsNullOrEmpty(OcrCode))
                        oDocument.Lines.CostingCode = OcrCode;

                    var OcrCode2 = receipt.CostingCode2;
                    if (!String.IsNullOrEmpty(OcrCode2))
                        oDocument.Lines.CostingCode2 = OcrCode2;

                    var OcrCode3 = receipt.CostingCode3;
                    if (!String.IsNullOrEmpty(OcrCode3))
                        oDocument.Lines.CostingCode3 = OcrCode3;

                    var OcrCode4 = receipt.CostingCode4;
                    if (!String.IsNullOrEmpty(OcrCode4))
                        oDocument.Lines.CostingCode4 = OcrCode4;

                    var OcrCode5 = receipt.CostingCode5;
                    if (!String.IsNullOrEmpty(OcrCode5))
                        oDocument.Lines.CostingCode5 = OcrCode5;


                    var res = oDocument.Add();
                    if (res != 0)
                    {
                        result = false;
                        Conexion.Conexion_SBO.m_oCompany.GetLastError(out res, out sErrMsg);
                        Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnRevalorizacion.Properties.Resources.NombreAddon + " Error: Revalorizacion de inventario, linea("+(receipt.LineNum+1)+"):"
                        + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        result = false;
                    }
                    else
                    {
                        result = true;
                        sErrMsg = "Se genero con éxito la " + (esSalida ? "salida":"entrada") ;
                        ActualizarEntradaMercancia(receipt, Conexion.Conexion_SBO.m_oCompany.GetNewObjectKey(), esSalida ? 1: 2) ;
                        

                    }

                }

            }
            catch (Exception ex)
            {
                result = false;
                ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
                sErrMsg = "Ocurrio un error al registrar la " + (esSalida ? "salida" : "entrada");
            }

            return new Tuple<bool, string>(result, sErrMsg);
        }


        private Tuple<bool, string> GenerarInventoryRevaluation(Receipt re)
        {
            string sErrMsg = "";
            SAPbobsCOM.Recordset oRS = null;
            bool result = false;

            try
            {
                oRS = (SAPbobsCOM.Recordset)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = " SELECT B.\"MdAbsEntry\" AS \"AbsEntry\", (SELECT \"U_SMF_IACC\" FROM \"@SMF_REVA\" WHERE \"Code\" = '001') AS \"AumentarCuenta\", (SELECT \"U_SMF_DACC\" FROM \"@SMF_REVA\" WHERE \"Code\" = '001') AS \"DisminuirCuenta\" ";
                query = query + " FROM OITL A  INNER JOIN ITL1 B ON A.\"LogEntry\" = B.\"LogEntry\" ";
                query = query + " WHERE A.\"DocEntry\" = '" + re.DocEntry.ToString() + "'  AND A.\"DocType\" = '20'  AND A.\"DocLine\"='" + re.LineNum.ToString() + "' ";

                oRS.DoQuery(query);

                if (oRS.RecordCount > 0)
                {
                    SAPbobsCOM.MaterialRevaluation oMaterialRevaluation = default(SAPbobsCOM.MaterialRevaluation);
                    SAPbobsCOM.SNBLines oMaterialRevaluationSNBLines = default(SAPbobsCOM.SNBLines);
                    oMaterialRevaluation = Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation);
                    oMaterialRevaluation.DocDate = System.DateTime.Now;
                    oMaterialRevaluation.RevalType = "P";
                    oMaterialRevaluation.Comments = "Added through DIAPI";
                    oMaterialRevaluation.Lines.ItemCode = re.Itemcode.ToString();
                    oMaterialRevaluation.Lines.WarehouseCode = re.WarehouseCode.ToString();
                    var OcrCode = re.CostingCode;
                    if (!String.IsNullOrEmpty(OcrCode))
                        oMaterialRevaluation.Lines.DistributionRule = OcrCode;
                    var OcrCode2 = re.CostingCode2;
                    if (!String.IsNullOrEmpty(OcrCode2))
                        oMaterialRevaluation.Lines.DistributionRule2 = OcrCode2;
                    var OcrCode3 = re.CostingCode3;
                    if (!String.IsNullOrEmpty(OcrCode3))
                        oMaterialRevaluation.Lines.DistributionRule3 = OcrCode3;
                    var OcrCode4 = re.CostingCode4;
                    if (!String.IsNullOrEmpty(OcrCode4))
                        oMaterialRevaluation.Lines.DistributionRule4 = OcrCode4;
                    var OcrCode5 = re.CostingCode5;
                    if (!String.IsNullOrEmpty(OcrCode5))
                        oMaterialRevaluation.Lines.DistributionRule5 = OcrCode5;

                    oMaterialRevaluation.Lines.RevaluationDecrementAccount = oRS.Fields.Item("DisminuirCuenta").Value;
                    oMaterialRevaluation.Lines.RevaluationIncrementAccount = oRS.Fields.Item("AumentarCuenta").Value;
                   

                    oRS.MoveFirst();
                    int cont = 0;
                    while (!oRS.EoF)
                    {
                        oMaterialRevaluationSNBLines = oMaterialRevaluation.Lines.SNBLines;
                        oMaterialRevaluationSNBLines.SetCurrentLine(cont);
                        oMaterialRevaluationSNBLines.SnbAbsEntry = oRS.Fields.Item("AbsEntry").Value;  //AbsEntry from OBTN Table
                        oMaterialRevaluationSNBLines.NewCost = (re.TotalLine/re.QuantityReal)*re.TcBase;
                        oMaterialRevaluationSNBLines.Add();
                        cont++;

                        oRS.MoveNext();
                    }

                    int RetVal = oMaterialRevaluation.Add();
                    if (RetVal != 0)
                    {
                        Conexion.Conexion_SBO.m_oCompany.GetLastError(out RetVal, out sErrMsg);
                        Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnRevalorizacion.Properties.Resources.NombreAddon + " Error: Revalorizacion de inventario, linea ("+ (re.LineNum+1) +"):"
                        + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        result = false;
                    }
                    else
                    {
                        ActualizarEntradaMercancia(re, Conexion.Conexion_SBO.m_oCompany.GetNewObjectKey(), 3);
                        sErrMsg = "Se genero con éxito la revalorizacion de inventario";
                        result = true;
                    }
                }



                
            }
            catch (Exception ex)
            {
                result = false;
                ExceptionLogging.HandleExcepcion(ex, this.GetType(), MethodBase.GetCurrentMethod().Name);
                sErrMsg = "Ocurrio un error al registrar la revalorizacion de inventario";
            }

            return new Tuple<bool, string>(result, sErrMsg);
        }



        private void ActualizarEntradaMercancia(Receipt receipt, string VINV, int type)
        {
            SAPbouiCOM.Form oForm = null;
            //SAPbobsCOM.JournalEntries oJE = null;

            SAPbobsCOM.Recordset oRecordSet = null;
            string sScript = string.Empty;
            string errMsg = string.Empty;
            try
            {

                oRecordSet = (SAPbobsCOM.Recordset)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                try
                {
                    sScript = Comunes.Consultas.ActualizarDocumento(Conexion.Conexion_SBO.m_oCompany.DbServerType, receipt.DocEntry, receipt.LineNum , VINV, type);
                    oRecordSet.DoQuery(sScript);
                }
                catch (Exception ex)
                {
                    errMsg = ex.Message;
                }

            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnRevalorizacion.Properties.Resources.NombreAddon + " Error: Entrega de mercancias:"
                    + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);//Mensaje de error
            }
        }


    }
}

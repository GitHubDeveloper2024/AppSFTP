using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;
using GemBox.Spreadsheet;
using Ionic.Zip;
using MDDBCDataAccess.Maestros;
using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;
using Renci.SshNet;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks; 



namespace AppSFTP
{
    public static class Program
    {
        static string sBD = "INTERNET_PERU.dbo.";
        static string sRutaExcel = System.Configuration.ConfigurationManager.AppSettings["DWH_STOCK"];
        //static string sRutaCentros = System.Configuration.ConfigurationManager.AppSettings["RUTA_CENTROS"];
        static string sConexion = ConfigurationManager.ConnectionStrings["cnxOracle"].ConnectionString;
        static string sEsquema = System.Configuration.ConfigurationManager.AppSettings["_Esquema"];
        static string sRutaDestino = System.Configuration.ConfigurationManager.AppSettings["_RutaDestino"];

        static void Main(string[] args)
        {

            fn_GenerarGarantiaExcel();

            // fn_PA_STOCKVALORIZADO_LISTARXFECHA(DateTime.Now.ToShortDateString());
            //fn_GenerarEECCExcel();

            //fn_GeneraPedidoExcel();

            /*
              fn_SFTP("(serie)", "Atlas_Entel_Reporte_Stock_SB",
                  "Pa_Atlas_Entel_Reporte_Stock_ObtenerColumnas_SB",8,1);
              */
            //fn_GenerarExcel("07/02/2023", "0");

            /*
            fn_SFTP("recepción", "Atlas_Entel_Reporte_Ingreso_SB",
            "Pa_Atlas_Entel_Reporte_Ingreso_SB_ObtenerColumnas_SB",5,2);
            */


            // SALIDA POR SERIE

            /*
            fn_SFTP("(detalle_serie)", "Atlas_Entel_Reporte_Salidas_SB",
            "Pa_Atlas_Entel_Reporte_Salidas_SB_ObtenerColumnas_SB",7,3);
            */
            //(general) = LOTES

            /*
            fn_SFTP("(general)", "Atlas_Entel_Reporte_Salidas_SB",
            "Pa_Atlas_Entel_Reporte_Salidas_SB_ObtenerColumnas_SB",7,4);
            */

            /*
            fn_SFTP_ConstruirExcel("(serie)", "Atlas_Entel_Reporte_Stock_SB",
                "Pa_Atlas_Entel_Reporte_Stock_ObtenerColumnas_SB", 7, 1);
            */

            /*
            fn_SFTP_ConstruirExcelNativo("(serie)", "Atlas_Entel_Reporte_Stock_SB",
              "Pa_Atlas_Entel_Reporte_Stock_ObtenerColumnas_SB", 7, 1);
            */

            //fn_ObtenerValorStock();
            //fn_Obtener_DWH();
            //fn_ObtenerValorCentrosBulkCopy();
            //fn_Obtener_DWH_BulkCopy();
            //fn_Obtener_CentrosSQL();
            //fn_ConvertirDataTableToCSV();
        }

        /*
         
        De servidor "Shadow": 
Tabla: PLC_INF.HOM_INAR
Hacia servidor: "Pods_Lm"
Tabla: CI_INAR_HIST
         
         */

        static void fn_TransferirShadow_PODS()
        {
            foreach(DataRow oRows in fn_ObtenerResultado("select NUMPERIODO,\r\nFECFECHAPROCESO,\r\nFECFECHAACTIVACION,\r\nVCHRAZONSOCIAL,\r\nVCHC_CONTRATO,\r\nVCHIMEI,\r\nVCHIMEI_BSCS,\r\nVCHTELEFONO,\r\nVCHMODELOEQUIPO,\r\nVCHC_PLAN,\r\nVCHN_PLAN,\r\nVCHESTADOINAR,\r\nVCHMOVIMIENTOS,\r\nNUMGROSS,\r\nVCHVENDEDOR,\r\nVCHTIPODOCUMENTO,\r\nVCHDOCUMENTO,\r\nNUMNRO_ORDEN,\r\nVCHPRODUCT,\r\nNUMRENTABASICA,\r\nNUMRENTAIGV,\r\nVCHSEGMENTO,\r\nVCHMODOPAGO,\r\nVCHTECNOLOGIA,\r\nVCHCLASIFICACIONRENTA,\r\nVCHPLAN_BLINDAJE,\r\nVCHVENDEDOR_PACKSIM,\r\nVCHPROMO_CHIPS,\r\nNUMCARGOFIJO,\r\nVCHTIPOVENTA,\r\nVCHCOMBO_CANAL,\r\nVCHPROD_VENTAREGULAR,\r\nVCHDWH_CODIGOORDEN,\r\nVCHDWH_ORDENCREADOPOR,\r\nVCHDWH_NOMBRECONSULTOR,\r\nVCHDWH_PRODUCTO,\r\nVCHMODEL_F,\r\nVCHSIM_DESBLOQUEADO,\r\nVCHPTVJAVA_PRODUCTO,\r\nVCHPTVJAVA_SKU,\r\nVCHPTVJAVA_PROMOTOR,\r\nVCHPORTA_MODOPAGOORIGEN,\r\nVCHPORTA_CEDENTE,\r\nVCHPORTA_RECEPTOR,\r\nVCHJER_PDV,\r\nVCHJER_GERENCIACANAL,\r\nVCHJER_CANALVENTA,\r\nVCHJER_KAM,\r\nVCHJER_TERRITORIO,\r\nVCHJER_DIVTERRITORIO,\r\nVCHJER_CADENADEALER,\r\nVCHJER_SOCIODENEGOCIO,\r\nVCHJER_TIPOTIENDA,\r\nVCHJER_DEPARTAMENTO,\r\nVCHJER_PROVINCIA,\r\nVCHJER_CIUDAD,\r\nVCHJER_DISTRITO,\r\nVCHJER_JEFENEGOCIO,\r\nVCHTERMINAL_GAMA,\r\nVCHTERMINAL_MARCA,\r\nVCHTERMINAL_MODELO,\r\nVCHTERMINAL_TIPOEQUIPO,\r\nVCHTERMINAL_NOMBREEQUIPO,\r\nVCHTERMINAL_TECNOLOGIA,\r\nVCHVEN_PACKSIM_DEP,\r\nVCHVEN_PACKSIM_PROV,\r\nVCHVEN_PACKSIM_DIST,\r\nVCHTIPOCONTRIBUYENTE,\r\nNUMDWH_PRECIOLISTA,\r\nNUMDWH_PRECIOPAGADO,\r\nNUMDWH_SUBTOTAL,\r\nVCHCORP_CANALVENTA,\r\nVCHCORP_VISTACLIENTE,\r\nVCHCORP_VISTANEGOCIO,\r\nNUMCANTIDADRECARGAS,\r\nNUMMONTORECARGAS,\r\nNUMTOTAL_MES_ANT,\r\nNUMTOTAL_MES_ACT,\r\nNUMTOTAL_FINAL,\r\nNUMDESCUENTO_2DALINEA,\r\nVCHTIPO_DESC_2DALINEA,\r\nVCHJER_NIVEL_TC,\r\nVCHCELDA_DEP,\r\nVCHCELDA_PROV,\r\nVCHCELDA_DIST,\r\nVCHCELDA_TIPOACTIVACION,\r\nVCHCELDA_GNT,\r\nVCHCELDA_USER_ID,\r\nVCHCELDA_TIPO_VENTA,\r\nVCHCELDA_SSNN,\r\nVCHCELDA_KAM,\r\nVCHCELDA_TOP_PDV,\r\nVCHSUSCRIPCIONESPLAN,\r\nVCHPREPAGO_TIPOCHIP,\r\nVCHPREPAGO_REGION_DEPART,\r\nNUMPRIMERA_REC,\r\nVCHTIPO_NEGOCIO,\r\nNUMSEMANA_ANHIO,\r\nVCHSINCRITERIOBASE,\r\nVCHJER_CLUSTER_GLOBAL,\r\nVCHJER_CLUSTER,\r\nNUMREC_QREC7,\r\nNUMREC_MONTO7,\r\nNUMREC_QREC15,\r\nNUMREC_MONTO15,\r\nVCHRECARGA_7,\r\nVCHRECARGA_15,\r\nVCHEMPRESAS_JER_DEPART,\r\nVCHEMPRESAS_JER_PROVIN,\r\nVCHEMPRESAS_JER_DISTRI,\r\nVCHCORP_DEPARTAMENTO,\r\nVCHCORP_PROVINCIA,\r\nVCHCORP_DISTRITO,\r\nFECFECH_CREACION_PORTA,\r\nNUMVEP_CUOTA,\r\nNUMVEP_PAGOTOTAL,\r\nNUMVEP_CUOTA_INICIAL,\r\nVCHVEP_FLAG_NUEVO,\r\nVCHVEP_FLAG_TOTAL,\r\nVCHCELDA_IMSI_SELLER,\r\nVCHCELDA_IMEI_SELLER,\r\nVCHSISTEMA_FUENTE,\r\nNUMPORTIN,\r\nVCHBONO,\r\nVCHCHNL_TDE,\r\nVCHACTIVATION_TYPE,\r\nVCHUSERNAME,\r\nVCHFLAG_APAGON,\r\nFECFECHA_APAGON,\r\nVCHEST_APAGON,\r\nVCHFLAG_REC_FECACT,\r\nVCHDESC_ORDER,\r\nVCHTIPORUC,\r\nVCHDOCUMENTO_CIERRE,\r\nVCHVENDEDOR_PACKSIM_OR,\r\nVCHENTEL_PRO,\r\nVCHPLAN_CHIP,\r\nVCHCICLOFACTURACION,\r\nVCHCODIGOCOMPANIA,\r\nVCHC_CONTRATOFS,\r\nVCHCODIGO,\r\nVCHNUMERO_RUC,\r\nVCHFLAG_ALTO_VALOR,\r\nVCHALMACEN,\r\nVCHPDV_PICKUP,\r\nVCHDIVTR_PICKUP,\r\nNUMFLAG_APERTURA,\r\nVCHFLAG_PRODUCTO,\r\nVCHNIVEL_GRUPO,\r\nVCHCLUSTER,\r\nFECFECHACREACION,\r\nVCHLLAA_BASE_CAPTURA,\r\nVCHVENDEDORDNI,\r\nVCHJER_CANALVENTATLV,\r\nNUMFLAGT0,\r\nVCHTYPEOFSALE,\r\nNUMFLAGRFT,\r\nVCHDETALLECAMPANA,\r\nNUMFLAG_OFERTA,\r\nVCHCLUSTER2,\r\nVCHCANAL2,\r\nVCHIMSISELLER,\r\nVCHAA_WEB,\r\nVCHJER_PDVRETAIL,\r\nVCHFLAG_LLA_RENO,\r\nVCHJER_CAMPANACANAL,\r\nVCHJER_CANALCHILE,\r\nVCHJER_CANALVENTA2,\r\nNUMMONTO_ORDEN,\r\nVCHFLAG_RECARGA7DIAS,\r\nVCHFLAG_LM,\r\nVCHJER_CAMPANAAGRUPADA,\r\nVCHFLAG_VALIDACIONBIO,\r\nVCHFLAG_UPSELLING,\r\nVCHFLAG_VEP2,\r\nNUMRENTAIGV_NETO,\r\nNUMRENTAIGV_NESTRUCTURAL,\r\nVCHDNI_LIDER,\r\nVCHN_PLANPILOTO,\r\nNUMRENTAIGV_NETOT,\r\nVCHDEP_ENTREGA_OL,\r\nVCHPROV_ENTREGA_OL,\r\nVCHDIST_ENTREGA_OL,\r\nVCHCLUSTER_DELIVERY,\r\nVCHFLAG_RENODIARIO,\r\nVCHFLAG_SINTERES,\r\nVCHCONCEPTO_DESC,\r\nNUMVALOR_DESC from PLC_INF.HOM_INAR where PERIODO='202306' ").Rows)
            {



            }

        }

        static void fn_PA_STOCKVALORIZADO_LISTARXFECHA(string sFecha)
        {
            try
            {

                DBCGeneric oDbGeneric = new DBCGeneric();
                int xCount = 0;
                string sSQL = "";
                string ITEM = "0";
                DataTable oOBJ = oDbGeneric.fn_ObtenerResultado("PA_DWH_STOCK_OBK_FECHA_LISTAR_EXCEL", sFecha);
                Console.WriteLine("oOBJ "+ oOBJ.Rows.Count);
                foreach (DataRow row in oOBJ.Rows)
                {

                    ITEM = row["ITEM"].ToString();


                    sSQL = "INSERT INTO " + sEsquema + "EH_CONSOLIDADO_STOCK (ATTRIBUTE1 , NPORGANIZATIONID , NPSECONDARYCODE , LLAVE , ES_KIT , NPSEGMENT1 , NPQUANTITY , VALORIZADO , FECHAREGISTRO , ARCHIVOPROCESADO , CODIGO , DESCRIPCION , ATRIBUTO1 , ORGANIZACION , SUBINVENTARIO ,  LLAVE1  , PUNTOVENTA , CANAL , MERCADO , ACTIVOREPOSICION , CADENA , SOCIO , KAM , FECHAREGISTRO1 , FECHA_CARGA_SQL)" +
    " VALUES( '" + row[0]
    + "', '" + row[1]
    + "', '" + row[2]
    + "', '" + row[3]
    + "', '" + row[4]
    + "', '" + row[5]
    + "', '" + row[6]

    + "', '" + row[7]
    + "', '" + row[8]
    + "', '" + row[9]
    + "', '" + row[10]
    + "', '" + row[11]
    + "', '" + row[12]

    + "', '" + row[13]
    + "', '" + row[14]
    + "', '" + row[15]
    + "', '" + row[16]
    + "', '" + row[17]
    + "', '" + row[18]

    + "', '" + row[19]
    + "', '" + row[20]
    + "', '" + row[21]
    + "', '" + row[22]
    + "', '" + row[23]

    + "', '" + DateTime.Now.ToShortDateString()
    + "' ) ";

                    Console.WriteLine(sSQL);

                    fn_Registrar(sSQL);


                    Console.WriteLine(xCount);

                    oDbGeneric = new DBCGeneric();
                    oDbGeneric.fn_AdicionarObjeto("PA_INSERTADO_ADICIONAR", ITEM,sFecha);

                    xCount++;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }


        }


        static void fn_GeneraPedidoExcel()
        {
            try
            {
                DataTable oEECC = fn_ObtenerResultado("SELECT * FROM VIEW_PEDIDO_LISTAR_CORREO where NVL(CORREOENVIADO,'0') = '0' ORDER BY PKID DESC ");

                using (XLWorkbook wb = new XLWorkbook())
                {

                    wb.Worksheets.Add(oEECC, "OPWEB_Excel");

                    //string Ruta = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\EntelDrive Canal Indirecto y Fraudes - EECC_WEB_MAY\\";
                    string Ruta = @"C:\Users\RPA_Entel-PE11\Entel Peru S.A\EntelDrive Canal Indirecto y Fraudes - Documentos\MACROS_CI\MONITOR SOCIOS\EECC_WEB_MAY\\";
                    string fileName = DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + (DateTime.Now.Year - 2000); //Ruta + ".xlsx";

                    fileName = Ruta + "OPWEB_"+ fileName + ".xlsx";

                    Console.WriteLine(fileName);

                    if (!System.IO.File.Exists(fileName))
                    {
                        wb.SaveAs(fileName);
                    }
                    else
                    {
                        System.IO.File.Delete(fileName);
                        wb.SaveAs(fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //Console.ReadLine();
            }

        }



        static void fn_GenerarGarantiaExcel()
        {
            try
            {
                string sSQL = "SELECT * FROM " + sEsquema + "VIEW_GARANTIA_WEB_LISTAR ";
                DataTable oEECC = fn_ObtenerResultado(sSQL);

                using (XLWorkbook wb = new XLWorkbook())
                {

                    wb.Worksheets.Add(oEECC, "EECCExcel");

                    //string Ruta = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\EntelDrive Canal Indirecto y Fraudes - EECC_WEB_MAY\\";
                    string Ruta = sRutaDestino;
                    string fileName ="GARANTIA_"+ DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + (DateTime.Now.Year - 2000); //Ruta + ".xlsx";

                    fileName = Ruta + fileName + ".xlsx";

                    Console.WriteLine(fileName);

                    if (!System.IO.File.Exists(fileName))
                    {
                        wb.SaveAs(fileName);
                    }
                    else
                    {
                        System.IO.File.Delete(fileName);
                        wb.SaveAs(fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //   Console.ReadLine();
            }

        }


        static void fn_GenerarEECCExcel()
        {
            try
            {
                DataTable oEECC = fn_ObtenerResultado("SELECT * FROM VIEW_ESTADOCUENTA_LISTAR ");

                using (XLWorkbook wb = new XLWorkbook())
                {

                    wb.Worksheets.Add(oEECC, "EECCExcel");

                    //string Ruta = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\EntelDrive Canal Indirecto y Fraudes - EECC_WEB_MAY\\";
                    string Ruta = @"C:\Users\RPA_Entel-PE11\Entel Peru S.A\EntelDrive Canal Indirecto y Fraudes - Documentos\MACROS_CI\MONITOR SOCIOS\EECC_WEB_MAY\\";
                    string fileName = DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + (DateTime.Now.Year - 2000); //Ruta + ".xlsx";

                    fileName = Ruta + fileName+".xlsx";

                    Console.WriteLine(fileName);

                    if (!System.IO.File.Exists(fileName))
                    {
                        wb.SaveAs(fileName);                         
                    }
                    else
                    {
                        System.IO.File.Delete(fileName);
                        wb.SaveAs(fileName);
                    }
                 }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
             //   Console.ReadLine();
            }

        }

        [Obsolete]
        static void fn_SFTP_ConstruirExcelNativo(string pTipo, string pDestinationTableName, string pStoreProcedure, int pColumnaAdicional, int pTipoEjecucion)
        {
            MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();
             
            string host = ConfigurationSettings.AppSettings["host"].ToString();
            string username = ConfigurationSettings.AppSettings["username"].ToString();
            string password = ConfigurationSettings.AppSettings["password"].ToString();
            string workingdirectory = ConfigurationSettings.AppSettings["workingdirectory"].ToString();
            string uploadfile = ConfigurationSettings.AppSettings["uploadfile"].ToString();
            int port = Convert.ToInt32(ConfigurationSettings.AppSettings["port"]);
            string sUltimaFecha = "";
            Console.WriteLine("Creating client and connecting");
            string sIDG = "";
            string sExcepcion = "";
            using (var client = new SftpClient(host, port, username, password))
            {
                client.Connect();
                Console.WriteLine("Connected to {0}", host);

                client.ChangeDirectory(workingdirectory);
                Console.WriteLine("Changed directory to {0}", workingdirectory);

                var listDirectory = client.ListDirectory(workingdirectory);
                Console.WriteLine("Listing directory:");

                var oLinq = (from oObj in listDirectory
                             where oObj.FullName.ToString().Contains(pTipo)
                             //where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("17/04/2023").ToShortDateString())
                              //&& Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("17/04/2023").ToShortDateString())
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                foreach (var fi in oLinq)
                {

                    Console.WriteLine(" - " + fi.Name);
                    try
                    {
                        sUltimaFecha = fi.LastWriteTime.ToShortDateString();

                        if (pTipoEjecucion == 1)
                        {
                            if (sUltimaFecha != "")
                            {
                                try
                                {
                                    //oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    //oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALDOS", sUltimaFecha);

                                    //   fn_GenerarExcel(sUltimaFecha, "1");
                                    fn_GenerarExcel_WorkBook(sUltimaFecha, "1");
                                }
                                catch (Exception ex)
                                {

                                    sExcepcion = ex.Message;
                                    oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, ex.Message);
                                    //Console.WriteLine(es.Message);
                                }

                            }
                        }
                        if (pTipoEjecucion == 2)
                        {
                            // oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            //  oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_INGRESOS");
                        }
                        if (pTipoEjecucion == 3) //SERIE
                        {
                            //  oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            // oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                        if (pTipoEjecucion == 4) //LOTE
                        {
                            // oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            //  oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                    }
                    catch (Exception es)
                    {
                        if (sIDG == "")
                            sIDG = "0";

                        sExcepcion = es.Message;
                        oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                        oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, es.Message);
                        Console.WriteLine(es.Message);
                        //Console.ReadLine();
                    }
                }

            }
        }

        [Obsolete]
        static void fn_SFTP_ConstruirExcel(string pTipo, string pDestinationTableName, string pStoreProcedure, int pColumnaAdicional, int pTipoEjecucion)
        {
            MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();

            string host = ConfigurationSettings.AppSettings["host"].ToString();
            string username = ConfigurationSettings.AppSettings["username"].ToString();
            string password = ConfigurationSettings.AppSettings["password"].ToString();
            string workingdirectory = ConfigurationSettings.AppSettings["workingdirectory"].ToString();
            string uploadfile = ConfigurationSettings.AppSettings["uploadfile"].ToString();
            int port = Convert.ToInt32(ConfigurationSettings.AppSettings["port"]);
            string sUltimaFecha = "";
            Console.WriteLine("Creating client and connecting");
            string sIDG = "";
            string sExcepcion = "";
            using (var client = new SftpClient(host, port, username, password))
            {
                client.Connect();
                Console.WriteLine("Connected to {0}", host);

                client.ChangeDirectory(workingdirectory);
                Console.WriteLine("Changed directory to {0}", workingdirectory);

                var listDirectory = client.ListDirectory(workingdirectory);
                Console.WriteLine("Listing directory:");

                var oLinq = (from oObj in listDirectory
                             where oObj.FullName.ToString().Contains(pTipo)
                             //where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("18/02/2023").ToShortDateString())
                             && Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("20/03/2023").ToShortDateString())
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                foreach (var fi in oLinq)
                {

                    Console.WriteLine(" - " + fi.Name);
                    try
                    {
                        sUltimaFecha = fi.LastWriteTime.ToShortDateString();
                         
                        if (pTipoEjecucion == 1)
                        {
                            if (sUltimaFecha != "")
                            {
                                try
                                {
                                    //oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    //oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALDOS", sUltimaFecha);

                                    fn_GenerarExcel_Fast(sUltimaFecha, "1");
                                }
                                catch (Exception ex)
                                {

                                    sExcepcion = ex.Message;
                                    oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, ex.Message);
                                    //Console.WriteLine(es.Message);
                                }

                            }
                        }
                        if (pTipoEjecucion == 2)
                        {
                            // oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            //  oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_INGRESOS");
                        }
                        if (pTipoEjecucion == 3) //SERIE
                        {
                            //  oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            // oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                        if (pTipoEjecucion == 4) //LOTE
                        {
                            // oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            //  oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                    }
                    catch (Exception es)
                    {
                        if (sIDG == "")
                            sIDG = "0";

                        sExcepcion = es.Message;
                        oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                        oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, es.Message);
                        Console.WriteLine(es.Message);
                        //Console.ReadLine();
                    }
                }

            }
        }

        [Obsolete]
        static void fn_SFTP_ConstruirExcel_Manual(string pTipo, string pDestinationTableName, string pStoreProcedure, int pColumnaAdicional, int pTipoEjecucion)
        {
            MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();

            string host = ConfigurationSettings.AppSettings["host"].ToString();
            string username = ConfigurationSettings.AppSettings["username"].ToString();
            string password = ConfigurationSettings.AppSettings["password"].ToString();
            string workingdirectory = ConfigurationSettings.AppSettings["workingdirectory"].ToString();
            string uploadfile = ConfigurationSettings.AppSettings["uploadfile"].ToString();
            int port = Convert.ToInt32(ConfigurationSettings.AppSettings["port"]);
            string sUltimaFecha = "";
            Console.WriteLine("Creating client and connecting");
            string sIDG = "";
            string sExcepcion = "";
            using (var client = new SftpClient(host, port, username, password))
            {
                client.Connect();
                Console.WriteLine("Connected to {0}", host);

                client.ChangeDirectory(workingdirectory);
                Console.WriteLine("Changed directory to {0}", workingdirectory);

                var listDirectory = client.ListDirectory(workingdirectory);
                Console.WriteLine("Listing directory:");

                var oLinq = (from oObj in listDirectory
                             where oObj.FullName.ToString().Contains(pTipo)
                             //where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("18/02/2023").ToShortDateString())
                             && Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("20/03/2023").ToShortDateString())
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                foreach (var fi in oLinq)
                {

                    Console.WriteLine(" - " + fi.Name);
                    try
                    {
                        sUltimaFecha = fi.LastWriteTime.ToShortDateString();

                        if (pTipoEjecucion == 1)
                        {
                            if (sUltimaFecha != "")
                            {
                                try
                                {
                                    //oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    //oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALDOS", sUltimaFecha);

                                    fn_GenerarExcel_Fast(sUltimaFecha, "1");
                                }
                                catch (Exception ex)
                                {

                                    sExcepcion = ex.Message;
                                    oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, ex.Message);
                                    //Console.WriteLine(es.Message);
                                }

                            }
                        }
                        if (pTipoEjecucion == 2)
                        {
                            // oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            //  oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_INGRESOS");
                        }
                        if (pTipoEjecucion == 3) //SERIE
                        {
                            //  oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            // oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                        if (pTipoEjecucion == 4) //LOTE
                        {
                            // oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            //  oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                    }
                    catch (Exception es)
                    {
                        if (sIDG == "")
                            sIDG = "0";

                        sExcepcion = es.Message;
                        oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                        oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, es.Message);
                        Console.WriteLine(es.Message);
                        //Console.ReadLine();
                    }
                }

            }
        }

        static DataTable fn_ObtenerResultado(string pQuery)
        {
            Console.WriteLine(sConexion);

            DataTable oObj = new DataTable();
            using (OracleDataAdapter adp = new OracleDataAdapter(pQuery, sConexion))
            {
                adp.Fill(oObj);//all the data in OracleAdapter will be filled into Datatable 

            }
            return oObj;
        }

        [Obsolete]
        static void fn_Obtener_CentrosSQL()
        {

            //OBTENER EL ARCHIVO

            string sQueryExcel = "";

            try
            {
                // C:\Users\RPA_Entel-PE11\Entel Peru S.A\EntelDrive Canal Indirecto y Fraudes - VALORSTOCK_CSV
                // VALOR_STOCK_WEB_20231227.csv
                //string sRuta = @"\\pelma3w12pap12v\compartido\WOPEREP055\Reporte";
                //string sRuta = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\Bernales Oliden, Felipe - Bases Maestras";
                //string sRuta = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\Bernales Oliden, Felipe - Bases Maestras";
                string sRuta = @"C:\Users\RPA_Entel-PE11\Entel Peru S.A\Bocanegra Blas, Jamir Alfredo - BASES MAESTRAS";
                DirectoryInfo listDirectory = new DirectoryInfo(sRuta);
                FileInfo[] files = listDirectory.GetFiles("*");
                string str = "";


                var oLinq = (from oObj in files
                                 //where oObj.LastWriteTime.ToString().Contains(DateTime.Now.ToShortDateString())
                                 //where oObj.FullName.ToString().Contains(pTipo)
                                 //  where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("14/06/2023").ToShortDateString())
                                 //&& Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("17/04/2023").ToShortDateString())
                             where oObj.FullName.ToString().Contains("centro")
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                string sRutaCompleta = "";
                string NuevaRuta = "C:\\ArchivosMigrados\\";
                //string sRutaArchivoCentro = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\Bernales Oliden, Felipe - Bases Maestras\\20230719_centro.xlsx";
                //string sRutaArchivoCentro = "";

                foreach (FileInfo file in oLinq)
                {
                    sRutaCompleta = sRuta + "\\" + file.Name;
                    Console.WriteLine(sRutaCompleta);
                    NuevaRuta = NuevaRuta + file.Name;

                    /*
                    if (File.Exists(NuevaRuta))
                        File.Delete(NuevaRuta);

                    File.Copy(sRutaCompleta, NuevaRuta);
                    */
                }
              

                string sLLAVEConsulta = "";

                //string sQueryExcelCentrol = "select * from [Hoja1$] where [ACTIVO PARA REPOSICION]= 1  ";
                string sQueryExcelCentrol = "select * from [Hoja1$] ";
                Console.WriteLine("sQueryExcelCentrol='" + sQueryExcelCentrol + "'");
                //DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", "select * from [Hoja1$] where COD='" + sLLAVEConsulta + "'");
                DataTable oObjCentro = fn_LeerExcel(sRutaCompleta, "Hoja1", sQueryExcelCentrol);

                int PKID = 1;

                Console.WriteLine("sRutaArchivoCentro='" + sRutaCompleta + "'");
               // oDbGeneric.fn_AdicionarObjeto("PA_CENTROS_VALOR_STOCK_FECHA_Eliminar");
                foreach (DataRow row in oObjCentro.Rows)
                { 
                    DBCGeneric oDbGeneric=new   DBCGeneric();
                    oDbGeneric.fn_AdicionarObjeto("PA_CENTROS_VALOR_STOCK_FECHA_Adicionar",
                      fn_Decodificar (row[0].ToString()), // CODIGO
                      fn_Decodificar(row[1].ToString()), // DESCRIPCION
                        fn_Decodificar(row[2].ToString()), // ATRIBUTO1 
                        fn_Decodificar(row[3].ToString()), //ORGANIZACION 
                        fn_Decodificar(row[4].ToString()), //SUBINVENTARIO 
                        fn_Decodificar( row[5] .ToString()), //@LLAVE  
                        fn_Decodificar(row[6].ToString()), //@PUNTOVENTA 
                        fn_Decodificar(row[7].ToString()), // @CANAL 
                        fn_Decodificar(row[8].ToString()), //@MERCADO  
                        fn_Decodificar(row[9].ToString()), //@ACTIVOREPOSICION  
                        fn_Decodificar(row[14].ToString()), //@CADENA 
                        fn_Decodificar(row[15].ToString()), //@SOCIO 
                        fn_Decodificar(row[26].ToString()), //@KAM 
                        DateTime.Now.ToString());
                        //DateTime.Now.AddDays(-1));

                    Console.WriteLine(row[0]);

                    //fn_Registrar(sQueryCentros);
                }

                //Console.WriteLine("bError='" + bError.ToString() + "'");

                //string sQUERY_1 = "SELECT NVL(MAX(PKID),'0') PKID FROM " + sEsquema + "DWH_STOCK ";

                //DataTable oMIGRACIONTRANSFERENCIA_U = fn_ObtenerResultado(sQUERY_1);

                //Console.WriteLine("COUNT IDENTI " + oMIGRACIONTRANSFERENCIA_U.Rows.Count);

                /*foreach (DataRow oRows2 in oMIGRACIONTRANSFERENCIA_U.Rows)
                {
                    PKID = (Convert.ToInt32(oRows2["PKID"]));
                    Console.WriteLine("IDENTI " + PKID);
                }*/

                //Console.WriteLine("IDENTI " + PKID);

                //PKID = PKID + 1;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //Console.ReadLine();
            }
        }


        static string fn_Decodificar(string sValor)
        {
            string sReturnValue = "";
                         
             sReturnValue  =System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetBytes(sValor));

            //sReturnValue Encoding.UTF8.GetString(sva);


            return sReturnValue;
        }

        static string fn_Decodificar_CSV(string sValor)
        {
            string sReturnValue = "";

            //if (sValor.Contains("682"))
            //{
            sReturnValue = System.Net.WebUtility.HtmlDecode(sValor);
            //   sReturnValue = Regex.Unescape(sValor);

            //sReturnValue = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetBytes(sValor));
            //}

            //sReturnValue  =System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetBytes(sValor));

            //sReturnValue Encoding.UTF8.GetString(sva);


            return sReturnValue;
        }

        public static void fn_ConvertirDataTableToCSV()
        {


            try
            {

            DBCGeneric dBCGeneric = new DBCGeneric();
            DataTable oObj = dBCGeneric.fn_ObtenerResultado("PA_STOCK_VALORIZADO_GENERAR_CSV", DateTime.Now.ToShortDateString());// DateTime.Now.ToShortTimeString());

            ToCSV(oObj);

            }
            catch (Exception EX)
            {
                Console.WriteLine(EX.Message);
                Console.ReadLine();
            }
        }

        public static void ToCSV(this DataTable dt)
        {
            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join(",", fields));
            }
            //C:\Users\RPA_Entel-PE11\Entel Peru S.A\EntelDrive Canal Indirecto y Fraudes - VALORSTOCK_CSV
            string sNombreCSC = "C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\EntelDrive Canal Indirecto y Fraudes - VALORSTOCK_CSV\\VALOR_STOCK_WEB_" +
                +DateTime.Now.Year + "" + DateTime.Now.Month.ToString("00") + "" + (DateTime.Now.Day).ToString("00") + ".csv";
            //string sNombreCSC = "C:\\CSV_GENERADO\\VALOR_STOCK_WEB_" + DateTime.Now.Year+""+ DateTime.Now.Month+""+ DateTime.Now.Day + ".csv";


            File.WriteAllText(sNombreCSC, sb.ToString());
        }

        [Obsolete]
        static void fn_Obtener_DWH_BulkCopy()
        {

            //OBTENER EL ARCHIVO

            string sQueryExcel = "";

            try
            {
                Console.WriteLine("ENTRO");

                string sRuta = @"\\pelma3w12pap12v\compartido\WOPEREP055\Reporte";
                //string sRuta = @"C:\Users\RPA_Entel-PE11\Entel Peru S.A\EntelDrive Canal Indirecto y Fraudes - VALORSTOCK_CSV";
                DirectoryInfo listDirectory = new DirectoryInfo(sRuta);
                FileInfo[] files = listDirectory.GetFiles("*");
                string str = "";


                var oLinq = (from oObj in files
                               //  where oObj.LastWriteTime==DateTime.Now.AddDays(-3)
                           //where oObj.Name == "DWH_Stocks_20230730.csv"
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                string sRutaCompleta = "";

                string NuevaRuta = "C:\\ArchivosMigrados\\";

                //string sRutaArchivoCentro = "";

                Console.WriteLine("ENTRO 2");

                

                Console.WriteLine("oLinq.Count " + oLinq.ToList().Count());

                
                foreach (FileInfo file in oLinq)
                {
                    Console.WriteLine(file.Name);
                    sRutaCompleta = sRuta + "\\" + file.Name;
                    Console.WriteLine(sRutaCompleta);
                    NuevaRuta = NuevaRuta + file.Name;

                    Console.WriteLine("NuevaRuta  " + NuevaRuta);

                    if (File.Exists(NuevaRuta))
                        File.Delete(NuevaRuta);

                    File.Copy(sRutaCompleta, NuevaRuta);

                }
                DataTable oObjDWH = null;
                bool bError = false;

                oObjDWH = ConvertCSVtoDataTable_2(NuevaRuta, 0);

                string consStringINTERNET = ConfigurationManager.ConnectionStrings["Conexion_BDInternet"].ConnectionString;
                int iColum = 0;

                DBCGeneric oDBGeneric = new DBCGeneric();
                oDBGeneric.fn_AdicionarObjeto("PA_DWH_STOCK_OBK_ELIMINAR");


                using (SqlConnection con = new SqlConnection(EncriptacionMartin.MetodoEncriptacion.Desencriptar(consStringINTERNET)))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        sqlBulkCopy.BulkCopyTimeout = 10800;

                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo." + "DWH_STOCK_OBK";

                        // oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                        foreach (DataColumn oColumnasSOLICITUD_IMP in oObjDWH.Columns)
                        {
                            sqlBulkCopy.ColumnMappings.Add(iColum, iColum);
                            iColum++;
                        }
                        con.Open();
                        sqlBulkCopy.WriteToServer(oObjDWH);
                        con.Close();
                    }
                }

                oDBGeneric = new DBCGeneric();
                oDBGeneric.fn_AdicionarObjeto("PA_DWH_STOCK_OBK_FECHA_ELIMINAR", NuevaRuta);
      

            }
            catch (Exception ex)
            {
                Console.WriteLine("err "+ ex.Message);
                Console.ReadLine();
            }
        }


        public static void SaveUsingOracleBulkCopy(string destTableName, DataTable dt)
        {
            try
            {
                using (var connection = new  OracleConnection(sConexion))
                {
                    connection.Open();
                    using (var bulkCopy = new  OracleBulkCopy(connection,  OracleBulkCopyOptions.UseInternalTransaction))
                    {
                        bulkCopy.DestinationTableName = destTableName;
                        bulkCopy.BulkCopyTimeout = 600;
                        bulkCopy.WriteToServer(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                    Console.ReadLine();
                
            }
        }
        [Obsolete]
        static void fn_Obtener_DWH()
        {

            //OBTENER EL ARCHIVO

            string sQueryExcel = "";

            try
            {

                string sRuta = @"\\pelma3w12pap12v\compartido\WOPEREP055\Reporte";

                DirectoryInfo listDirectory = new DirectoryInfo(sRuta);
                FileInfo[] files = listDirectory.GetFiles("*");
                string str = "";


                var oLinq = (from oObj in files
                                 //where oObj.LastWriteTime.ToString().Contains(DateTime.Now.ToShortDateString())
                                 //where oObj.FullName.ToString().Contains(pTipo)
                                 //  where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("14/06/2023").ToShortDateString())
                                 //&& Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("17/04/2023").ToShortDateString())
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                string sRutaCompleta = "";
                string NuevaRuta = "C:\\ArchivosMigrados\\";
                string sRutaArchivoCentro = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\Bernales Oliden, Felipe - Bases Maestras\\20230617_centro.xlsx";
                //string sRutaArchivoCentro = "";

                foreach (FileInfo file in oLinq)
                {
                    sRutaCompleta = sRuta + "\\" + file.Name;
                    Console.WriteLine(sRutaCompleta);
                    NuevaRuta = NuevaRuta + file;

                    if (File.Exists(NuevaRuta))
                        File.Delete(NuevaRuta);

                    File.Copy(sRutaCompleta, NuevaRuta);

                }
                DataTable oObjDWH = null;
                bool bError = false;
                try
                {
                    oObjDWH = ConvertCSVtoDataTable_2(NuevaRuta, 0);

                    if (oObjDWH.Rows.Count == 0)
                    {
                        bError = true;
                    }
                }
                catch (Exception ex)
                {
                    bError = true;
                }

                if (bError)
                {
                    string sNombreHoja = "DWH_Stocks_" + DateTime.Now.Year + "" + DateTime.Now.Month.ToString("00") + "" + DateTime.Now.Day.ToString("00"); //DWH_Stocks_20230612
                    //DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", "select * from [Hoja1$] where COD='" + sLLAVEConsulta + "'");
                    oObjDWH = fn_LeerExcel(sRutaArchivoCentro, sNombreHoja, "select * from [" + sNombreHoja + "] ");
                }

                string sLLAVEConsulta = "";

                //string sQueryExcelCentrol = "select * from [Hoja1$] where [ACTIVO PARA REPOSICION]= 1  ";
                string sQueryExcelCentrol = "select * from [Hoja1$] ";
                Console.WriteLine("sQueryExcelCentrol='" + sQueryExcelCentrol + "'");
                //DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", "select * from [Hoja1$] where COD='" + sLLAVEConsulta + "'");
                DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", sQueryExcelCentrol);

                int PKID = 1;

                Console.WriteLine("sRutaArchivoCentro='" + sRutaArchivoCentro + "'");
                Console.WriteLine("bError='" + bError.ToString() + "'");

                string sQUERY_1 = "SELECT NVL(MAX(PKID),'0') PKID FROM " + sEsquema + "DWH_STOCK ";

                DataTable oMIGRACIONTRANSFERENCIA_U = fn_ObtenerResultado(sQUERY_1);

                Console.WriteLine("COUNT IDENTI " + oMIGRACIONTRANSFERENCIA_U.Rows.Count);

                foreach (DataRow oRows2 in oMIGRACIONTRANSFERENCIA_U.Rows)
                {
                    PKID = (Convert.ToInt32(oRows2["PKID"]));
                    Console.WriteLine("IDENTI " + PKID);
                }

                Console.WriteLine("IDENTI " + PKID);

                PKID = PKID + 1;

                foreach (DataRow row in oObjDWH.Rows)
                {
                    string sScriptQuery = "INSERT INTO " + sEsquema + "DWH_VALOR_STOCK (ATTRIBUTE1,NPORGANIZATIONID,NPSECONDARYCODE,ES_KIT,NPSEGMENT1,NPQUANTITY,VALORIZADO,FECHAREGISTRO)" +
                            " VALUES( '"
                            + row[0]
                            + "', '" + row[1]
                            + "', '" + row[2]
                            + "', '" + row[3]
                            + "', '" + row[4]
                            + "', '" + row[5]
                            + "', '" + row[6]
                            + "', '" + DateTime.Now.ToString()
                            + "' ) ";

                    Console.WriteLine("sScriptQuery " + sScriptQuery);

                    fn_Registrar(sScriptQuery);                      
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }


        [Obsolete]
        static void fn_ObtenerValorStock()
        {

            //OBTENER EL ARCHIVO

            string sQueryExcel = "";

            try
            {

                string sRuta=  @"\\pelma3w12pap12v\compartido\WOPEREP055\Reporte";

                DirectoryInfo listDirectory = new DirectoryInfo(sRuta);
                FileInfo[] files = listDirectory.GetFiles("*");
                string str = "";


                var oLinq = (from oObj in files
                               //where oObj.LastWriteTime.ToString().Contains(DateTime.Now.ToShortDateString())
                                 //where oObj.FullName.ToString().Contains(pTipo)
                               //  where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("14/06/2023").ToShortDateString())
                                 //&& Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("17/04/2023").ToShortDateString())
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                string sRutaCompleta = "";
                string NuevaRuta = "C:\\ArchivosMigrados\\";
                string sRutaArchivoCentro  = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\Bernales Oliden, Felipe - Bases Maestras\\20230617_centro.xlsx";
                //string sRutaArchivoCentro = "";

                foreach (FileInfo file in oLinq)
                {
                    sRutaCompleta = sRuta + "\\" + file.Name;
                    Console.WriteLine(sRutaCompleta);
                    NuevaRuta = NuevaRuta + file;

                    if (File.Exists(NuevaRuta))
                        File.Delete(NuevaRuta);

                    File.Copy(sRutaCompleta, NuevaRuta);

                }
                DataTable oObj2 = null;
                bool bError = false;
                try
                {
                    oObj2 = ConvertCSVtoDataTable_2(NuevaRuta, 0);

                    if (oObj2.Rows.Count == 0)
                    {
                        bError = true;
                    }
                }
                catch (Exception ex)
                { 
                    bError = true;   
                }

                if (bError)
                {
                    string sNombreHoja = "DWH_Stocks_"+ DateTime.Now.Year + "" + DateTime.Now.Month.ToString("00")+""+ DateTime.Now.Day.ToString("00"); //DWH_Stocks_20230612
                    //DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", "select * from [Hoja1$] where COD='" + sLLAVEConsulta + "'");
                    oObj2 = fn_LeerExcel(sRutaArchivoCentro,sNombreHoja, "select * from ["+ sNombreHoja + "] ");
                }

                string sLLAVEConsulta = "";

                //string sQueryExcelCentrol = "select * from [Hoja1$] where [ACTIVO PARA REPOSICION]= 1  ";
                string sQueryExcelCentrol = "select * from [Hoja1$] ";
                Console.WriteLine("sQueryExcelCentrol='" + sQueryExcelCentrol + "'");
                //DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", "select * from [Hoja1$] where COD='" + sLLAVEConsulta + "'");
                DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", sQueryExcelCentrol);

                int PKID = 1;

                Console.WriteLine("sRutaArchivoCentro='" + sRutaArchivoCentro + "'");
                Console.WriteLine("bError='" + bError.ToString() + "'");

                string sQUERY_1 = "SELECT NVL(MAX(PKID),'0') PKID FROM " + sEsquema + "DWH_STOCK ";

                DataTable oMIGRACIONTRANSFERENCIA_U = fn_ObtenerResultado(sQUERY_1);

                Console.WriteLine("COUNT IDENTI " + oMIGRACIONTRANSFERENCIA_U.Rows.Count);

                foreach (DataRow oRows2 in oMIGRACIONTRANSFERENCIA_U.Rows)
                {
                    PKID = (Convert.ToInt32(oRows2["PKID"]));
                    Console.WriteLine("IDENTI " + PKID);
                }

                Console.WriteLine("IDENTI " + PKID);

                PKID = PKID + 1;

                foreach (DataRow row in oObj2.Rows)
                {

                    sLLAVEConsulta = row[1] + "" + row[2];
                    
                        //"and " + oObj2.Columns[4].ColumnName + " ='" + row[2] + "' ";
                    DataView oDataView = oObjCentro.DefaultView;

                    foreach (var xColumns in oDataView.ToTable().Columns)
                    {
                        Console.WriteLine("ColumnName='" + xColumns.ToString() + "'");
                    }

                    string sFiltro = "["+oDataView.ToTable().Columns[3].ColumnName + "]='" + row[1]+ "' and " +
                    "[" + oDataView.ToTable().Columns[4].ColumnName + "]='" + row[2] + "'" ;
                    //+ "' and '" 
                    //+ oDataView.ToTable().Columns[4].ColumnName +"'='" + row[2] + "'";
                    Console.WriteLine("sFiltro=" + sFiltro );
                    oDataView.RowFilter = sFiltro;

                    //+ "' and  \"[MTL_SECONDARY_INVENTORIES-SECONDARY_INVENTORY_NAME] Subinventory name\" ='" + row[2]+"' ";

                    //Console.WriteLine("COD='" + sLLAVEConsulta + "'" );
                    //Console.ReadLine();


                    foreach (DataRow row2 in oDataView.ToTable().Rows)
                    {
                        Console.WriteLine("ENCONTRADO");
                        Console.WriteLine(row2["PUNTO DE VENTA"].ToString() + " | "+ row2["CANAL COMERCIAL"].ToString()+ " | " + row2["SOCIO"].ToString());

                        fn_Registrar("INSERT INTO " + sEsquema + "DWH_STOCK (PKID,ATTRIBUTE1,NPORGANIZATIONID,NPSECONDARYCODE,ES_KIT,NPSEGMENT1,NPQUANTITY,VALORIZADO,CANAL,SOCIO,PUNTOVENTA,DIRECCION,CODIGOCENTRO,CENTRO,MERCADO,KAM,FECHAREGISTRO)"+
                            " VALUES( " + PKID + " ,'" + row[0] 
                            + "', '" + row[1]
                            + "', '" + row[2]
                            + "', '" + row[3]
                            + "', '" + row[4]
                            + "', '" + row[5]
                            + "', '" + row[6]

                            + "', '" + row2["CANAL COMERCIAL"]
                            + "', '" + row2["SOCIO"]
                            + "', '" + row2["PUNTO DE VENTA"]
                            + "', '" + row2["DIRECCIÓN"]
                            + "', '" + row2["CÓDIGO CENTRO"]
                            + "', '" + row2["NOMBRE CENTRO"]
                            + "', '" + row2["MERCADO"]
                            + "', '" + row2["KAM"]
                            + "', '"  + DateTime.Now.ToString()
                            + "' ) ");

                        PKID++;
                    }                    
                } 
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        [Obsolete]
        static void fn_ObtenerValorCentrosBulkCopy()
        {

            //OBTENER EL ARCHIVO

            string sQueryExcel = "";

            try
            {

                string sRuta = @"\\pelma3w12pap12v\compartido\WOPEREP055\Reporte";

                DirectoryInfo listDirectory = new DirectoryInfo(sRuta);
                FileInfo[] files = listDirectory.GetFiles("*");
                string str = "";


                var oLinq = (from oObj in files
                                 //where oObj.LastWriteTime.ToString().Contains(DateTime.Now.ToShortDateString())
                                 //where oObj.FullName.ToString().Contains(pTipo)
                                 //  where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("14/06/2023").ToShortDateString())
                                 //&& Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("17/04/2023").ToShortDateString())
                             orderby oObj.LastWriteTime descending
                             select oObj).Take(1);

                string sRutaCompleta = "";
                string NuevaRuta = "C:\\ArchivosMigrados\\";
                string sRutaArchivoCentro = @"C:\\Users\\RPA_Entel-PE11\\Entel Peru S.A\\Bernales Oliden, Felipe - Bases Maestras\\20230719_centro.xlsx";
                //string sRutaArchivoCentro = "";

                foreach (FileInfo file in oLinq)
                {
                    sRutaCompleta = sRuta + "\\" + file.Name;
                    Console.WriteLine(sRutaCompleta);
                    NuevaRuta = NuevaRuta + file;

                    if (File.Exists(NuevaRuta))
                        File.Delete(NuevaRuta);

                    File.Copy(sRutaCompleta, NuevaRuta);

                }
                

                

                string sLLAVEConsulta = "";

                //string sQueryExcelCentrol = "select * from [Hoja1$] where [ACTIVO PARA REPOSICION]= 1  ";
                string sQueryExcelCentrol = "select * from [Hoja1$] ";
                Console.WriteLine("sQueryExcelCentrol='" + sQueryExcelCentrol + "'");
                //DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", "select * from [Hoja1$] where COD='" + sLLAVEConsulta + "'");
                DataTable oObjCentro = fn_LeerExcel(sRutaArchivoCentro, "Hoja1", sQueryExcelCentrol);

                int PKID = 1;

                Console.WriteLine("sRutaArchivoCentro='" + sRutaArchivoCentro + "'");
                //Console.WriteLine("bError='" + bError.ToString() + "'");

                string sQUERY_1 = "SELECT NVL(MAX(PKID),'0') PKID FROM " + sEsquema + "DWH_STOCK ";

                DataTable oMIGRACIONTRANSFERENCIA_U = fn_ObtenerResultado(sQUERY_1);

                Console.WriteLine("COUNT IDENTI " + oMIGRACIONTRANSFERENCIA_U.Rows.Count);

                foreach (DataRow oRows2 in oMIGRACIONTRANSFERENCIA_U.Rows)
                {
                    PKID = (Convert.ToInt32(oRows2["PKID"]));
                    Console.WriteLine("IDENTI " + PKID);
                }

                Console.WriteLine("IDENTI " + PKID);

                foreach(DataColumn oRows in oObjCentro.Columns)
                {
                    Console.WriteLine( oRows.ColumnName);
                    Console.ReadLine();

                }

                PKID = PKID + 1;

                
                    try
                    {
                        OracleConnection oracleConnection = new OracleConnection(sConexion);

                        oracleConnection.Open();
                        using (OracleBulkCopy bulkCopy = new OracleBulkCopy(oracleConnection))
                        {
                            bulkCopy.DestinationTableName = "CENTROS_VALOR_STOCK";
                        foreach (DataColumn dtColumn in oObjCentro.Columns)
                        {
                            bulkCopy.ColumnMappings.Add(dtColumn.ColumnName.ToUpper(), dtColumn.ColumnName.ToUpper());
                        }

                        bulkCopy.WriteToServer(oObjCentro);
                        }
                        oracleConnection.Close();
                        oracleConnection.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);                        
                    }
                 
                 
                foreach (DataRow row in oObjCentro.Rows)
                {
                    string sQueryCentros = "INSERT INTO " + sEsquema + "CENTROS_VALOR_STOCK (CODIGO,DESCRIPCION,ATRIBUTO1,ORGANIZACION,SUBINVENTARIO,LLAVE,PUNTOVENTA,CANAL,MERCADO,ACTIVOREPOSICION,CADENA,SOCIO,KAM,FECHAREGISTRO)" +
                            " VALUES('" + row[0]
                            + "', '" + row[1]
                            + "', '" + row[2]
                            + "', '" + row[3]
                            + "', '" + row[4]
                            + "', '" + row[5]
                            + "', '" + row[6]

                            + "', '" + row[7]
                            + "', '" + row[8]
                            + "', '" + row[9]
                            + "', '" + row[10]
                            + "', '" + row[11]
                            + "', '" + row[12]

                            + "', '" + DateTime.Now.ToString()
                            + "' ) ";

                    Console.WriteLine(sQueryCentros);

                        fn_Registrar(sQueryCentros);
                     
                    }
                  
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }


        static string fn_Registrar(string pQUERY)
        {
            using (var con = new Oracle.ManagedDataAccess.Client.OracleConnection(sConexion))
            {
                con.Open();

                OracleParameter id = new OracleParameter();
                id.OracleDbType = OracleDbType.Varchar2;
                id.Value = DateTime.Now.ToLongDateString();

                // create command and set properties
                OracleCommand cmd = con.CreateCommand();
                cmd.CommandText = pQUERY;  //"INSERT INTO BULKINSERTTEST (ID, NAME, ADDRESS) VALUES (:1, :2, :3)";
                                           //cmd.ArrayBindCount = ids.Length;
                                           //cmd.Parameters.Add(id);
                                           //cmd.Parameters.Add(name);
                                           //cmd.Parameters.Add(address);
                cmd.ExecuteNonQuery();

            }
            return "1";
        }


        public static DataTable fn_LeerExcel(string pRutaBaseArchivo, string pHoja, string sQueryExcel)
        {
            DataTable oObj = new DataTable();
            string sBase = pRutaBaseArchivo; //System.Configuration.ConfigurationManager.AppSettings["param1"] ;
            string sHoja = pHoja;//System.Configuration.ConfigurationManager.AppSettings["param2"] ;

            //string sBase =  System.Configuration.ConfigurationManager.AppSettings["param1"] ;
            //string sHoja = System.Configuration.ConfigurationManager.AppSettings["param2"] ;

            int xValor = 1;
            //public static string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=excel 12.0;";

            //Fuente: https://www.iteramos.com/pregunta/9358/excel-quotla-tabla-externa-no-tiene-el-formato-esperadoquot
            System.Data.DataSet DtSet;
            DtSet = new System.Data.DataSet();

            Console.WriteLine("********Conectandose a archivo Excel************************" + pRutaBaseArchivo);

            Console.WriteLine("********Inicio ********" + DateTime.Now);
            //Console.Read();
            string sConex = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + sBase + "'; Extended Properties = 'Excel 8.0;HDR={1}'";

            //using (System.Data.OleDb.OleDbConnection oOleDbConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + sBase + "';Extended Properties=excel 12.0;"))
            using (System.Data.OleDb.OleDbConnection oOleDbConnection = new System.Data.OleDb.OleDbConnection(sConex))

            //using (System.Data.OleDb.OleDbConnection oOleDbConnection = new System.Data.OleDb.OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + sBase + "';Extended Properties = \"Excel 12.0 xlsx;  HDR=YES; IMEX=1\";"))
            {
                oOleDbConnection.Open();
                using (System.Data.OleDb.OleDbDataAdapter oOleDbDataAdapterTotal =
        new System.Data.OleDb.OleDbDataAdapter(sQueryExcel, oOleDbConnection))
                {

                    Console.WriteLine("Obteniendo registros de la hoja  " + sHoja + "............");
                    //  Console.Read();

                    oOleDbDataAdapterTotal.TableMappings.Add("Table", "TestTable");
                    DtSet = new System.Data.DataSet();
                    oOleDbDataAdapterTotal.Fill(DtSet);

                    oObj = DtSet.Tables[0];


                }

            }


            return oObj;
        }

        [Obsolete]
        static void fn_SFTP(string pTipo,string pDestinationTableName,string pStoreProcedure,int pColumnaAdicional,int pTipoEjecucion)
        {
            MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();

              string host = ConfigurationSettings.AppSettings["host"].ToString();  
              string username = ConfigurationSettings.AppSettings["username"].ToString(); 
              string password = ConfigurationSettings.AppSettings["password"].ToString();
              string workingdirectory = ConfigurationSettings.AppSettings["workingdirectory"].ToString();  
              string uploadfile = ConfigurationSettings.AppSettings["uploadfile"].ToString();  
            int port  =Convert.ToInt32( ConfigurationSettings.AppSettings["port"]);
            string sUltimaFecha = "";
            Console.WriteLine("Creating client and connecting");
            string sIDG ="";
            string sExcepcion = "";
            using (var client = new SftpClient(host, port, username, password))
            {
                client.Connect();
                Console.WriteLine("Connected to {0}", host);

                client.ChangeDirectory(workingdirectory);
                Console.WriteLine("Changed directory to {0}", workingdirectory);

                var listDirectory = client.ListDirectory(workingdirectory);
                Console.WriteLine("Listing directory:");
                 
                var oLinq = (from oObj in listDirectory
                             where oObj.FullName.ToString().Contains(pTipo)
                             //where Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("18/05/2023").ToShortDateString())
                                 && Convert.ToDateTime(oObj.LastWriteTime.ToShortDateString()) == Convert.ToDateTime(Convert.ToDateTime("11/06/2023").ToShortDateString())
                             orderby oObj.LastWriteTime descending
                            select   oObj).Take(1) ;
                
                foreach (var fi in oLinq)
                { 

                    Console.WriteLine(" - " + fi.Name);
                    try
                    {
                        sUltimaFecha = fi.LastWriteTime.ToShortDateString();
                     
                       using (Stream file1 = File.Create(uploadfile + @"\" + fi.Name))
                        {
                            client.DownloadFile(workingdirectory + @"/" + fi.Name, file1);

                            oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();
                            oDBGeneric.fn_AdicionarObjeto(sBD + "PA_ATLAS_TRANSFERENCIA_ARCHIVO_Adicionar", fi.LastWriteTime, fi.Name, fi.FullName, fi.Attributes.Size , pTipoEjecucion);
                        }

                        oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();
                        int xContadorFile = 0;
                        foreach (DataRow oRows in oDBGeneric.fn_ObtenerResultado(sBD + "PA_ATLAS_TRANSFERENCIA_ARCHIVO_PENDIENTE").Rows)
                        {
                            sIDG = oRows["ID"].ToString();
                            DataTable oObj = ConvertCSVtoDataTable_2(uploadfile + @"\" + oRows["NombreArchivo"].ToString(), pColumnaAdicional);

                            string consStringINTERNET = ConfigurationManager.ConnectionStrings["Conexion_BDInternet"].ConnectionString;
                            int iColum = 0;

                            using (SqlConnection con = new SqlConnection(EncriptacionMartin.MetodoEncriptacion.Desencriptar(consStringINTERNET)))
                            {
                                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                {
                                    sqlBulkCopy.BulkCopyTimeout = 10800;

                                    //Set the database table name
                                    sqlBulkCopy.DestinationTableName = "dbo." + pDestinationTableName;

                                    oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    foreach (DataRow oColumnasSOLICITUD_IMP in oDBGeneric.fn_ObtenerResultado(pStoreProcedure).Rows)
                                    {
                                        sqlBulkCopy.ColumnMappings.Add(iColum, iColum);
                                        iColum++;
                                    }
                                    con.Open();
                                    sqlBulkCopy.WriteToServer(oObj);
                                    con.Close();
                                }
                            }

                            oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_MIGRADO", oRows["ID"].ToString(),pTipoEjecucion);

                        }

                        oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                        oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_LIMPIAR_CAB");

                        if (pTipoEjecucion == 1)
                        {
                            if (sUltimaFecha != "")
                            {
                                try
                                {
                                    oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALDOS", sUltimaFecha);


                                }
                                catch (Exception ex)
                                {

                                    sExcepcion = ex.Message;
                                    oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                                    oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, ex.Message);
                                    //Console.WriteLine(es.Message);
                                }
                                
                            }
                        }
                        if (pTipoEjecucion == 2)
                        {
                            oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_INGRESOS");
                        }
                        if (pTipoEjecucion == 3) //SERIE
                        {
                            oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                        if (pTipoEjecucion == 4) //LOTE
                        {
                            oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                            oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_ENTEL_REPORTE_SALIDAS");
                        }
                    }
                    catch (Exception es)
                    {
                        if (sIDG == "")
                            sIDG = "0";

                        sExcepcion = es.Message;                        
                        oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric("Conexion_BDInternet");
                        oDBGeneric.fn_AdicionarObjeto("PA_ATLAS_TRANSFERENCIA_ARCHIVO_ERROR", sIDG, es.Message);
                        Console.WriteLine(es.Message);
                        //Console.ReadLine();
                    }
                }

            }
        }

        public static void ImportDataTable(string pFileName, string pSheetName, string pStartReference, DataTable pDataTable, bool pIncludeHeaders = true)
        {
            using (SLDocument doc = new SLDocument(pFileName, pSheetName))
            {
                doc.ImportDataTable(pStartReference, pDataTable, pIncludeHeaders);
                doc.Save();
            }
        }


        public static void fn_GenerarExcel_Fast_ultimo(string pFecha, string pID)
        {
            try
            {
                SpreadsheetDocument oSpreadsheetDocument = null;

                MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();
                Workbook book = new Workbook();

                DataTable oObj = oDBGeneric.fn_ObtenerResultado(sBD + "spAtlasReporte_001_StockRedes", pFecha);

                Microsoft.Office.Interop.Excel.Application xlAppToExport = new Microsoft.Office.Interop.Excel.Application();
                xlAppToExport.Workbooks.Add("");


                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport = default(Microsoft.Office.Interop.Excel.Worksheet);
                xlWorkSheetToExport = (Microsoft.Office.Interop.Excel.Worksheet)xlAppToExport.Sheets["Лист1"];


                int iRowCnt = 4;


                xlWorkSheetToExport.Cells[1, 1] = "Inform about mutilation";

                Microsoft.Office.Interop.Excel.Range range = xlWorkSheetToExport.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                range.EntireRow.Font.Name = "Palatino Linotype";
                range.EntireRow.Font.Bold = true;
                range.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A1:E1"].MergeCells = true;

                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NID";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "LastName";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Name";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Birthday";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Address";

                int i;
                for (i = 0; i <= oObj.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = oObj.Rows[i].Field<Int32>("NID");
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = oObj.Rows[i].Field<string>("LastName");
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = oObj.Rows[i].Field<string>("Name");
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = oObj.Rows[i].Field<string>("Birthday");
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = oObj.Rows[i].Field<string>("Address");
                    iRowCnt = iRowCnt + 1;
                }


                Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[11, 1] as Microsoft.Office.Interop.Excel.Range;
                range1.AutoFormat();
                string path = "c:\\";
                xlWorkSheetToExport.SaveAs(path + "MutilationSheet.xlsx");


                xlAppToExport.Workbooks.Close();
                xlAppToExport.Quit();
                xlAppToExport = null;
                xlWorkSheetToExport = null;

 
                /*
                string sRutaGrabar = @"\\196.10.10.26\fclaser\ArchivosElectronicos\";

                string pNombreArchivo = sRutaGrabar + "STOCK_REDES_" + pFecha.Replace("/", "") + ".xls";

                string pNombreArchivo2 = sRutaGrabar + "STOCK_REDES_" + pFecha.Replace("/", "") + ".zip";

                //File.WriteAllBytes(sRutaGrabar + ".xls", stream.GetBuffer());

                if (File.Exists(pNombreArchivo))
                {
                    File.Delete(pNombreArchivo);
                }

                //workbook.Save(pNombreArchivo);

                using (ZipFile zipFile = new ZipFile())
                {
                    if (File.Exists(pNombreArchivo2))
                    {
                        File.Delete(pNombreArchivo2);
                    }
                    //zipFile.AddEntry(sRutaGrabar + ".zip", stream);
                    zipFile.AddFile(pNombreArchivo, string.Empty);
                    //zipFile.Save(Response.OutputStream);
                    zipFile.Save(pNombreArchivo2);
                }
                */

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }


        public static void fn_GenerarExcel_Fast(string pFecha, string pID)
        {
            try
            { 
                MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();

                DataTable oObj = oDBGeneric.fn_ObtenerResultado(sBD + "spAtlasReporte_001_StockRedes", pFecha);

                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                // Continue to use the component in a Trial mode when free limit is reached.
                SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

                var workbook = new ExcelFile();

                GemBox.Spreadsheet.ExcelWorksheet worksheet = workbook.Worksheets.Add("Atlas_StockRedes");
                 
                // Insert DataTable to an Excel worksheet.
                worksheet.InsertDataTable(oObj,
                    new InsertDataTableOptions()
                    {
                        ColumnHeaders = true
                    });

                string sRutaGrabar = @"\\196.10.10.26\fclaser\ArchivosElectronicos\";

                string pNombreArchivo = sRutaGrabar+ "STOCK_REDES_" + pFecha.Replace("/", "") + ".xls";

                string pNombreArchivo2 = sRutaGrabar + "STOCK_REDES_" + pFecha.Replace("/", "") + ".zip";

                //File.WriteAllBytes(sRutaGrabar + ".xls", stream.GetBuffer());

                if (File.Exists(pNombreArchivo))
                {
                    File.Delete(pNombreArchivo);
                }

                workbook.Save(pNombreArchivo);

                using (ZipFile zipFile = new ZipFile())
                {
                    if (File.Exists(pNombreArchivo2))
                    {
                        File.Delete(pNombreArchivo2);
                    }
                    //zipFile.AddEntry(sRutaGrabar + ".zip", stream);
                    zipFile.AddFile(pNombreArchivo, string.Empty);
                    //zipFile.Save(Response.OutputStream);
                    zipFile.Save(pNombreArchivo2);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }
        public static void fn_GenerarExcel_WorkBook(string pFecha, string pID)
        {

            using (XLWorkbook wb = new XLWorkbook())
        {
                MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();

                DataTable oObj = oDBGeneric.fn_ObtenerResultado(sBD + "spAtlasReporte_001_StockRedes", pFecha); 
                    wb.Worksheets.Add(oObj, "StockRedes");

                //wb.Worksheet("StockRedes").Cells("m3").Style.NumberFormat.Format = "0.00000000";
                //wb.Worksheet("StockRedes").Cell("m3").DataType = XLCellValues.Number;
                /*
                wb.Worksheet("StockRedes").Cells("serie_lote").Style.NumberFormat.Format="@";
                wb.Worksheet("StockRedes").Cells("m3").Style.NumberFormat.Format="0.00000000";
                wb.Worksheet("StockRedes").Cells("total_m3").Style.NumberFormat.Format="0.0000000";
                wb.Worksheet("StockRedes").Cells("m2").Style.NumberFormat.Format = "0.0000000";
                wb.Worksheet("StockRedes").Cells("total_m2").Style.NumberFormat.Format = "0.0000000";
                wb.Worksheet("StockRedes").Cells("m3").Style.NumberFormat.Format = "0.0000000";
                */

                //  wb.Cell ("serie_lote").Style.NumberFormat.SetFormat("@");
                //wb.Cell("cantidad").Style.Format.SetFormat("0");


                string Ruta = @"\\196.10.10.26\fclaser\ArchivosElectronicos\";
                //HOJA 2
                //string fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                //                        + "\\ReporteDetallado-" + NumeroComprobante + ".xlsx";

                string pFecha2 = pFecha.Replace("/", "");

                string fileName = "STOCK_REDES_" + pFecha2 + ".xlsx";

                if (!System.IO.File.Exists(Ruta + fileName))
                {
                    wb.SaveAs(Ruta + fileName);
                    //wb.SaveAs(folderPath + "Test_1.xlsx");

                }
                else
                {
                    System.IO.File.Delete(Ruta + fileName);
                    wb.SaveAs(Ruta + fileName);
                }


                if (File.Exists(Ruta + fileName.Replace(".xlsx", ".zip")))
                {
                    File.Delete(Ruta + fileName.Replace(".xlsx", ".zip"));
                }

                //workbook.Save(pNombreArchivo);

                using (ZipFile zipFile = new ZipFile())
                {
                    if (File.Exists(Ruta + fileName.Replace(".xlsx", ".zip")))
                    {
                        File.Delete(Ruta + fileName.Replace(".xlsx", ".zip"));
                    }
                    //zipFile.AddEntry(sRutaGrabar + ".zip", stream);
                    zipFile.AddFile(Ruta + fileName, string.Empty);
                    //zipFile.Save(Response.OutputStream);
                    zipFile.Save(Ruta + fileName.Replace(".xlsx", ".zip"));
                }

            }
        }
        public static void fn_GenerarExcel(string pFecha,string pID)
        {
            try
            {

                
                
                


                // string pFecha = (DateTime.Now.Day - 1).ToString();

                object missing = Type.Missing;
                int Fila = 6;
                int Fila2 = 6;

                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook oWB = oXL.Workbooks.Add(missing);
            Microsoft.Office.Interop.Excel.Worksheet oSheet = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            oSheet.Name = "STOCK_REDES";

            Microsoft.Office.Interop.Excel.Worksheet oSheet2 = oWB.Sheets.Add(missing, missing, 1, missing)
                          as Microsoft.Office.Interop.Excel.Worksheet;
            
            fn_AddColumns(oSheet2, Fila, "estado", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "tipostock", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "sku", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "lpn_fch_creacion", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "oc_nro", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "tipo_ingreso", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "descripcion", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "serie_lote", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "cantidad", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "almacen", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "m3", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "total_m3", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "m2", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "total_m2", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "projectoscharff", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "tipocontrol", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "lpn", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "fechadescarga", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "subinventario", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "proyecto", 1, false, true, false);
            fn_AddColumns(oSheet2, Fila, "tarea", 1, false, true, false);


            MDDBCDataAccess.Maestros.DBCGeneric oDBGeneric = new MDDBCDataAccess.Maestros.DBCGeneric();

            foreach (DataRow oRows in
             oDBGeneric.fn_ObtenerResultado(sBD+ "spAtlasReporte_001_StockRedes",
             pFecha).Rows)
            {
                Fila++;
                 
                fn_AddColumns(oSheet2, Fila, oRows["estado"].ToString(), 1, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["tipostock"].ToString(), 2, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["sku"].ToString(), 3, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["lpn_fch_creacion"].ToString(), 4, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["oc_nro"].ToString(), 5, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["tipo_ingreso"].ToString(), 6, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["descripcion"].ToString(), 7, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["serie_lote"].ToString(), 8, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["cantidad"].ToString(), 9, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["almacen"].ToString(), 10, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["m3"].ToString(), 11, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["total_m3"].ToString(), 12, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["m2"].ToString(), 13, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["total_m2"].ToString(), 14, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["projectoscharff"].ToString(), 15, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["tipocontrol"].ToString(), 16, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["lpn"].ToString(), 17, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["fechadescarga"].ToString(), 18, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["subinventario"].ToString(), 19, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["proyecto"].ToString(), 20, true, false);
                fn_AddColumns(oSheet2, Fila, oRows["tarea"].ToString(), 21, true, false);
                //
            }

            string Ruta = @"\\196.10.10.26\fclaser\ArchivosElectronicos\";
                //HOJA 2
                //string fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                //                        + "\\ReporteDetallado-" + NumeroComprobante + ".xlsx";

                string pFecha2 = pFecha.Replace("/", "");

            string fileName = "StockRedes-" + pFecha2 + ".xlsx";

            if (!System.IO.File.Exists(Ruta + fileName))
            {
                oWB.SaveAs(Ruta + fileName);
            }
            else
            {
                System.IO.File.Delete(Ruta + fileName);
                oWB.SaveAs(Ruta + fileName);
            }


                if (File.Exists(Ruta + fileName.Replace(".xlsx", ".zip")))
                {
                    File.Delete(Ruta + fileName.Replace(".xlsx", ".zip"));
                }

                //workbook.Save(pNombreArchivo);

                using (ZipFile zipFile = new ZipFile())
                {
                    if (File.Exists(Ruta + fileName))
                    {
                        File.Delete(Ruta + fileName);
                    }
                    //zipFile.AddEntry(sRutaGrabar + ".zip", stream);
                    zipFile.AddFile(Ruta + fileName, string.Empty);
                    //zipFile.Save(Response.OutputStream);
                    zipFile.Save(Ruta + fileName);
                }


            }
            catch (Exception EX)
            {
                Console.WriteLine(EX.Message);
            }
            //oDBGeneric = new DBCGeneric();
            //oDBGeneric.fn_AdicionarObjeto("SustentoFacturacion.Pa_SustentoComprobanteLiquidacion_ActualizarEstado", pID); ;


        }

        public static void fn_AddColumns(Microsoft.Office.Interop.Excel.Worksheet oSheet2, int Fila, string pNombre, int pColumna, bool AplicaFormato = false, bool AplicaColor = false, bool Derecha = false)
        {
            if (AplicaFormato)
                oSheet2.Cells[Fila - 2, pColumna].NumberFormat = "@";

            oSheet2.Cells[Fila - 2, pColumna] = pNombre;

            if (AplicaColor)
            {
                oSheet2.Cells[Fila - 2, pColumna].Font.Color = System.Drawing.Color.White;
                oSheet2.Cells[Fila - 2, pColumna].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#FF7483");
            }

            if (Derecha)
            {
                //oSheet2.Columns.ColumnWidth = 50;
                //oSheet2.Cells[Fila - 2, pColumna].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                //    oSheet2.Cells[Fila - 2, pColumna].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            }
            else
            {

                //  oSheet2.Cells[Fila - 2, pColumna].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            }

            oSheet2.Cells[Fila - 2, pColumna].EntireColumn.AutoFit();
        }

        public static DataTable ConvertCSVtoDataTable_2(string sCsvFilePath,int pColumnaAdicional)
        {
            DataTable dtTable = new DataTable();
            Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
            int xContador = 0;
            int iHeader = 0;

             
            if (!System.IO.File.Exists(sCsvFilePath)
                ) { System.IO.FileStream f = System.IO.File.Create(sCsvFilePath); f.Close(); }
            using (System.IO.StreamWriter sw = System.IO.File.AppendText(sCsvFilePath) )
            { //write my text
              //
              
            }

            //Encoding.UTF8 enc = Encoding::GetEncoding("utf-8");

            using (StreamReader sr = new StreamReader(sCsvFilePath, Encoding.Default ))
            { 
                string[] headers = sr.ReadLine().Split(',');

                iHeader = headers.Length;

                iHeader = iHeader + pColumnaAdicional;

                //foreach (string header in headers)
                for (int i = 0; i < iHeader; i++)
                //for (int i = 0; i < headers.Length; i++)S
                {
                    dtTable.Columns.Add(fn_ObtenerName(xContador));
                    xContador++;
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = CSVParser.Split(sr.ReadLine());
                    DataRow dr = dtTable.NewRow();
                    //for (int i = 0; i < iHeader; i++)
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = fn_Decodificar_CSV( rows[i].Replace("\"", string.Empty));
                    }
                    dtTable.Rows.Add(dr);
                }
            }

            return dtTable;
        }

        //HttpUtility.HtmlDecode(myEncodedString, myWriter);


        public static string fn_ObtenerName(int value)
        {
            string num2Text = "";
             
            if (value == 0) num2Text = "ATTRIBUTE1";
            else if (value == 1) num2Text = "NPORGANIZATIONID";
            else if (value == 2) num2Text = "NPSECONDARYCODE";
            else if (value == 3) num2Text = "ES_KIT";
            else if (value == 4) num2Text = "NPSEGMENT1";
            else if (value == 5) num2Text = "NPQUANTITY";
            else if (value == 6) num2Text = "VALORIZADO";
            else if (value == 7) num2Text = "SIETE";
            else if (value == 8) num2Text = "OCHO";
            else if (value == 9) num2Text = "NUEVE";
            else if (value == 10) num2Text = "DIEZ";
            else if (value == 11) num2Text = "ONCE_";
            else if (value == 12) num2Text = "DOCE";
            else if (value == 13) num2Text = "TRECE";
            else if (value == 14) num2Text = "CATORCE";
            else if (value == 15) num2Text = "QUINCE";

            else if (value == 16) num2Text = "DIECISEIS";
            else if (value == 17) num2Text = "DIECISIETE";
            else if (value == 18) num2Text = "DIECIOCHO";
            else if (value == 19) num2Text = "DIECINUEVE";
            else if (value == 20) num2Text = "VEINTE";
            else if (value == 21) num2Text = "VEINTEUNO";
            else if (value == 22) num2Text = "VEINTEDOS";
            else if (value == 23) num2Text = "VEINTETRES";
            else if (value == 24) num2Text = "VEINTECUATRO";
            else if (value == 25) num2Text = "VEINTECINCO";
            else if (value == 26) num2Text = "VEINTESEIS";
            else if (value == 27) num2Text = "VEINTESIETE";
            else if (value == 28) num2Text = "VEINTEOCHO";
            else if (value == 29) num2Text = "VEINTENUEVE";
            else if (value == 30) num2Text = "TREINTA";

            return num2Text;
        }

        public static string NumeroALetras(this decimal numberAsString)
        {
            string dec;

            var entero = Convert.ToInt64(Math.Truncate(numberAsString));
            var decimales = Convert.ToInt32(Math.Round((numberAsString - entero) * 100, 2));
            if (decimales > 0)
            {
                //dec = " PESOS CON " + decimales.ToString() + "/100";
                dec = $" PESOS {decimales:0,0} /100";
            }
            //Código agregado por mí
            else
            {
                //dec = " PESOS CON " + decimales.ToString() + "/100";
                dec = $" PESOS {decimales:0,0} /100";
            }
            var res = NumeroALetras(Convert.ToDecimal(entero)) + dec;
            return res;
        }
    }
}

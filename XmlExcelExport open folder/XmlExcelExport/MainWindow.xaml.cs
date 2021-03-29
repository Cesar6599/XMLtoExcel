using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.MessageBox;

namespace XmlExcelExport
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        public string[] ColumName = {"Comprobante_xmlns:xsi", "Comprobante_xmlns:tfd", "Comprobante_xsi:schemaLocation", "Comprobante_Version", "Comprobante_Folio", "Comprobante_Fecha", "Comprobante_Sello", "Comprobante_FormaPago", "Comprobante_NoCertificado", "Comprobante_Certificado", "Comprobante_SubTotal", "Comprobante_Moneda", "Comprobante_Total", "Comprobante_TipoDeComprobante", "Comprobante_MetodoPago", "Comprobante_LugarExpedicion", "Comprobante_xmlns:cfdi", "Emisor_Rfc", "Emisor_Nombre", "Emisor_RegimenFiscal", "Receptor_Rfc", "Receptor_Nombre", "Receptor_UsoCFDI", "Conceptos_Concepto_ClaveProdServ", "Conceptos_Concepto_Cantidad", "Conceptos_Concepto_ClaveUnidad", "Conceptos_Concepto_Unidad", "Conceptos_Concepto_Descripcion", "Conceptos_Concepto_ValorUnitario", "Conceptos_Concepto_Importe", "Conceptos_Traslado_Base", "Conceptos_Traslado_Impuesto", "Conceptos_Traslado_TipoFactor", "Conceptos_Traslado_TasaOCuota", "Conceptos_Traslado_Importe", "Impuestos_TotalImpuestosTrasladados", "Impuestos_Traslado_Impuesto", "Impuestos_Traslado_TipoFactor", "Impuestos_Traslado_TasaOCuota", "Impuestos_Traslado_Importe", "TimbreFiscalDigital_xsi:schemaLocation", "TimbreFiscalDigital_Version", "TimbreFiscalDigital_UUID", "TimbreFiscalDigital_FechaTimbrado", "TimbreFiscalDigital_RfcProvCertif", "TimbreFiscalDigital_SelloCFD", "TimbreFiscalDigital_NoCertificadoSAT", "TimbreFiscalDigital_SelloSAT" };
        public string[] FilesName = new string[0];
        public string[] XMlInfo = new string[0];
        public string DtPath = "", SavePath = "";
        public int NumFiles = 0;
        public DataTable DTData = new DataTable();
        private void btnSelectFiles_Click(object sender, RoutedEventArgs e)
        {
            Array.Resize(ref FilesName,0);
            Array.Resize(ref XMlInfo, 0);
            DtPath = "";
            SavePath = "";
            NumFiles = 0;
            DTData.Clear();
            try
            {
                FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
                fbd.Description = "Selecciona una carpeta";
                fbd.SelectedPath = System.IO.Path.GetDirectoryName(@"C:\");
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)   return;
                DtPath = fbd.SelectedPath + "\\"; 
                DirectoryInfo di = new DirectoryInfo(DtPath);
                foreach (var fi in di.GetFiles("*", SearchOption.AllDirectories))
                {
                    if (fi.FullName.Contains(".xml") || fi.FullName.Contains(".XML"))
                    {
                        Array.Resize(ref FilesName, FilesName.Length + 1);
                        FilesName[FilesName.Length - 1] = fi.FullName;
                        NumFiles++;
                    }
                }
                txtEditor.Text += NumFiles + " Archivos Encontrados" + "\n";
            }
            catch (Exception) { throw; }
        }
        private void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            DTData.Reset();
            #region Adding colums
            DTData.Columns.Add("Comprobante_xmlns:xsi", typeof(String));
            DTData.Columns.Add("Comprobante_xmlns:tfd", typeof(String));
            DTData.Columns.Add("Comprobante_xsi:schemaLocation", typeof(String));
            DTData.Columns.Add("Comprobante_Version", typeof(String));
            DTData.Columns.Add("Comprobante_Folio", typeof(String));
            DTData.Columns.Add("Comprobante_Fecha", typeof(String));
            DTData.Columns.Add("Comprobante_Sello", typeof(String));
            DTData.Columns.Add("Comprobante_FormaPago", typeof(String));
            DTData.Columns.Add("Comprobante_NoCertificado", typeof(String));
            DTData.Columns.Add("Comprobante_Certificado", typeof(String));
            DTData.Columns.Add("Comprobante_SubTotal", typeof(String));
            DTData.Columns.Add("Comprobante_Moneda", typeof(String));
            DTData.Columns.Add("Comprobante_Total", typeof(String));
            DTData.Columns.Add("Comprobante_TipoDeComprobante", typeof(String));
            DTData.Columns.Add("Comprobante_MetodoPago", typeof(String));
            DTData.Columns.Add("Comprobante_LugarExpedicion", typeof(String));
            DTData.Columns.Add("Comprobante_xmlns:cfdi", typeof(String));
            DTData.Columns.Add("Emisor_Rfc", typeof(String));
            DTData.Columns.Add("Emisor_Nombre", typeof(String));
            DTData.Columns.Add("Emisor_RegimenFiscal", typeof(String));
            DTData.Columns.Add("Receptor_Rfc", typeof(String));
            DTData.Columns.Add("Receptor_Nombre", typeof(String));
            DTData.Columns.Add("Receptor_UsoCFDI", typeof(String));
            DTData.Columns.Add("Conceptos_Concepto_ClaveProdServ", typeof(String));
            DTData.Columns.Add("Conceptos_Concepto_Cantidad", typeof(String));
            DTData.Columns.Add("Conceptos_Concepto_ClaveUnidad", typeof(String));
            DTData.Columns.Add("Conceptos_Concepto_Unidad", typeof(String));
            DTData.Columns.Add("Conceptos_Concepto_Descripcion", typeof(String));
            DTData.Columns.Add("Conceptos_Concepto_ValorUnitario", typeof(String));
            DTData.Columns.Add("Conceptos_Concepto_Importe", typeof(String));
            DTData.Columns.Add("Conceptos_Traslado_Base", typeof(String));
            DTData.Columns.Add("Conceptos_Traslado_Impuesto", typeof(String));
            DTData.Columns.Add("Conceptos_Traslado_TipoFactor", typeof(String));
            DTData.Columns.Add("Conceptos_Traslado_TasaOCuota", typeof(String));
            DTData.Columns.Add("Conceptos_Traslado_Importe", typeof(String));
            DTData.Columns.Add("Impuestos_TotalImpuestosTrasladados", typeof(String));
            DTData.Columns.Add("Impuestos_Traslado_Impuesto", typeof(String));
            DTData.Columns.Add("Impuestos_Traslado_TipoFactor", typeof(String));
            DTData.Columns.Add("Impuestos_Traslado_TasaOCuota", typeof(String));
            DTData.Columns.Add("Impuestos_Traslado_Importe", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_xsi:schemaLocation", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_Version", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_UUID", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_FechaTimbrado", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_RfcProvCertif", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_SelloCFD", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_NoCertificadoSAT", typeof(String));
            DTData.Columns.Add("TimbreFiscalDigital_SelloSAT", typeof(String));
            #endregion
            for (int i = 0; i < FilesName.Length; i++)
            {
                if (FilesName[i].Contains(".xml") || FilesName[i].Contains(".XML"))
                {
                    Array.Resize(ref XMlInfo, 0);
                    //MessageBox.Show(FilesName[i]);
                    XMlInfo = XXmlReader(FilesName[i]);
                    DataRow row = DTData.NewRow();
                    for (int j = 0; j < XMlInfo.Length; j++)
                    {
                        row[j] = XMlInfo[j];
                    }
                    DTData.Rows.Add(row);
                }
            }
            ExcelWriter(DTData);
        }
        public string[] XXmlReader(string path)
        {
            String[] ColNames = new string[48];
            XmlReader reader = XmlReader.Create(path);
            while (reader.Read())
            {
                if (reader.IsStartElement())
                {
                    //return only when you have START tag  
                    switch (reader.Name.ToString())
                    {
                        case "cfdi:Comprobante":
                            ColNames[0] = reader.GetAttribute("xmlns:xsi").ToString();
                            ColNames[1] = reader.GetAttribute("xmlns:tfd").ToString();
                            ColNames[2] = reader.GetAttribute("xsi:schemaLocation").ToString();
                            ColNames[3] = reader.GetAttribute("Version").ToString();
                            ColNames[4] = reader.GetAttribute("Folio").ToString();
                            ColNames[5] = reader.GetAttribute("Fecha").ToString();
                            ColNames[6] = reader.GetAttribute("Sello").ToString();
                            ColNames[7] = reader.GetAttribute("FormaPago").ToString();
                            ColNames[8] = reader.GetAttribute("NoCertificado").ToString();
                            ColNames[9] = reader.GetAttribute("Certificado").ToString();
                            ColNames[10] = reader.GetAttribute("SubTotal").ToString();
                            ColNames[11] = reader.GetAttribute("Moneda").ToString();
                            ColNames[12] = reader.GetAttribute("Total").ToString();
                            ColNames[13] = reader.GetAttribute("TipoDeComprobante").ToString();
                            ColNames[14] = reader.GetAttribute("MetodoPago").ToString();
                            ColNames[15] = reader.GetAttribute("LugarExpedicion").ToString();
                            ColNames[16] = reader.GetAttribute("xmlns:cfdi").ToString();
                            break;
                        case "cfdi:Emisor":
                            ColNames[17] = reader.GetAttribute("Rfc").ToString();
                            ColNames[18] = reader.GetAttribute("Nombre").ToString();
                            ColNames[19] = reader.GetAttribute("RegimenFiscal").ToString();
                            break;
                        case "cfdi:Receptor":
                            ColNames[20] = reader.GetAttribute("Rfc").ToString();
                            ColNames[21] = reader.GetAttribute("Nombre").ToString();
                            ColNames[22] = reader.GetAttribute("UsoCFDI").ToString();
                            break;
                        case "cfdi:Conceptos":
                            XmlReader rr = reader.ReadSubtree();
                            while (rr.Read())
                            {
                                if (rr.IsStartElement())
                                {
                                    switch (rr.Name.ToString())
                                    {
                                        case "cfdi:Concepto":
                                            ColNames[23] = reader.GetAttribute("ClaveProdServ").ToString();
                                            ColNames[24] = reader.GetAttribute("Cantidad").ToString();
                                            ColNames[25] = reader.GetAttribute("ClaveUnidad").ToString();
                                            ColNames[26] = reader.GetAttribute("Unidad").ToString();
                                            ColNames[27] = reader.GetAttribute("Descripcion").ToString();
                                            ColNames[28] = reader.GetAttribute("ValorUnitario").ToString();
                                            ColNames[29] = reader.GetAttribute("Importe").ToString();
                                            break;
                                        case "cfdi:Traslado":
                                            ColNames[30] = reader.GetAttribute("Base").ToString();
                                            ColNames[31] = reader.GetAttribute("Impuesto").ToString();
                                            ColNames[32] = reader.GetAttribute("TipoFactor").ToString();
                                            ColNames[33] = reader.GetAttribute("TasaOCuota").ToString();
                                            ColNames[34] = reader.GetAttribute("Importe").ToString();
                                            break;
                                    }
                                }
                            }
                            break;
                        case "cfdi:Impuestos":
                            ColNames[35] = reader.GetAttribute("TotalImpuestosTrasladados").ToString();
                            break;
                        case "cfdi:Traslado":
                            ColNames[36] = reader.GetAttribute("Impuesto").ToString();
                            ColNames[37] = reader.GetAttribute("TipoFactor").ToString();
                            ColNames[38] = reader.GetAttribute("TasaOCuota").ToString();
                            ColNames[39] = reader.GetAttribute("Importe").ToString();
                            break;
                        case "tfd:TimbreFiscalDigital":
                            ColNames[40] = reader.GetAttribute("xsi:schemaLocation").ToString();
                            ColNames[41] = reader.GetAttribute("Version").ToString();
                            ColNames[42] = reader.GetAttribute("UUID").ToString();
                            ColNames[43] = reader.GetAttribute("FechaTimbrado").ToString();
                            ColNames[44] = reader.GetAttribute("RfcProvCertif").ToString();
                            ColNames[45] = reader.GetAttribute("SelloCFD").ToString();
                            ColNames[46] = reader.GetAttribute("NoCertificadoSAT").ToString();
                            ColNames[47] = reader.GetAttribute("SelloSAT").ToString();
                            break;
                    }
                }
            }
            reader.Close();
            return ColNames;
        }
        public void ExcelWriter(DataTable DTData)
        {
            SavePath = "";
            try
            {
                //definir ruta de guardado
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.Title = "Save report File";
                saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";
                saveFileDialog.RestoreDirectory = true;
                if (saveFileDialog.ShowDialog() == true)
                {
                    SavePath = saveFileDialog.FileName;
                }
                Excel.Application excel = new Excel.Application();
                Excel._Workbook libro = null;
                Excel._Worksheet hoja = null;
                libro = (Excel._Workbook)excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                hoja = (Excel._Worksheet)libro.Worksheets.Add();
                hoja.Name = "Facturas";//agregar hoja
                for (int i = 0; i < DTData.Columns.Count; i++)//agregar nombre de cabeceras
                {
                    hoja.Cells[1, i + 1] = DTData.Columns[i].ColumnName;
                }
                //agregar la info
                for (int i = 0; i < DTData.Rows.Count; i++)
                {
                    for (int j = 0; j < DTData.Columns.Count; j++)
                    {
                        hoja.Cells[i + 2, j + 1] = DTData.Rows[i][j];
                    }
                }
                ((Excel.Worksheet)excel.ActiveWorkbook.Sheets["Hoja1"]).Delete();   //Borramos la hoja que crea en el libro por defecto
                libro.Saved = true;
                libro.SaveAs(SavePath);
                libro.Close();
                releaseObject(libro);
                excel.UserControl = false;
                excel.Quit();
                releaseObject(excel);
                txtEditor.Text += SavePath+ "\n";
            }
            catch (Exception)
            {throw;}
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("\nAn error occurred while releasing memory" + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
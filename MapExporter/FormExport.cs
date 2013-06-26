using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Output;
using ESRI.ArcGIS.Display;
using System.IO;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesRaster;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.DataManagementTools;

namespace MapExporter
{
  public partial class FormExport : Form
  {
    private int iDPI, tile;
    private string filename;

    public FormExport()
    {
      InitializeComponent();
    }

    private void button2_Click(object sender, EventArgs e)
    {
      var result = saveFileDialog1.ShowDialog();
      if (result == DialogResult.OK)
      {
        textBox2.Text = saveFileDialog1.FileName;
      }
    }

    private void button3_Click(object sender, EventArgs e)
    {
      if (!int.TryParse(textBox3.Text, out iDPI)) return;
      if (iDPI <= 0) return;
      if (!int.TryParse(comboBox1.Text, out tile)) return;
      if (tile < 1) return;

      filename = textBox2.Text;
      label5.Text = "正在出图...";

      IActiveView docActiveView = ArcMap.Document.ActivatedView;

      ExportPages(docActiveView, iDPI, filename, enumExportType.byBlockCount, 0, 0, tile);

      label5.Text = "出图完成";
      label1.Text = "";
    }
   
    enum enumExportType
    {
        byPaperSize,
        byBlockCount
    }
    private string ExportPages(IActiveView docActiveView, int iDPI, string filename, enumExportType exportType, double paperWidth, double paperHeight, int blockCount)
    {
        string filepath = filename.Substring(0, filename.LastIndexOf("\\"));
        string NameExt = System.IO.Path.GetFileName(filename);
        if (docActiveView is IPageLayout)
        {
            double width = 0, height = 0;   //实际长宽
            //计算分页地图大小 paperWidth,paperHeight
            //分块数量 blockCount
            //计算分页图片大小 iWidth，iHeight
            #region 计算出图参数
            var pPageLayout = docActiveView as IPageLayout;
            pPageLayout.Page.QuerySize(out width, out height);
            IUnitConverter pUnitCon = new UnitConverterClass();
            switch (exportType)
            {
                case enumExportType.byPaperSize:
                    if (paperWidth <= 0 || paperHeight <= 0)
                    { paperWidth = 210; paperHeight = 297; }
                    paperWidth = pUnitCon.ConvertUnits(paperWidth, esriUnits.esriMillimeters, pPageLayout.Page.Units);
                    paperHeight = pUnitCon.ConvertUnits(paperHeight, esriUnits.esriMillimeters, pPageLayout.Page.Units);
                    bool bW = (width > paperWidth + 0.001) ? true : false;
                    bool bH = (height > paperHeight + 0.001) ? true : false;
                    blockCount = (bW && bH) ? 4 : 1;
                    break;
                case enumExportType.byBlockCount:
                default:
                    if (blockCount < 1) blockCount = 1;
                    paperWidth = width / blockCount;
                    paperHeight = height / blockCount;
                    break;
            }
            int iWidth = (int)(pUnitCon.ConvertUnits(paperWidth, pPageLayout.Page.Units, esriUnits.esriInches) * iDPI);
            int iHeight = (int)(pUnitCon.ConvertUnits(paperHeight, pPageLayout.Page.Units, esriUnits.esriInches) * iDPI);
            #endregion

            if (System.IO.File.Exists(filename))
            {
                var pWS = OpenWorkspace(filename, enumWsFactoryType.Raster) as IRasterWorkspace;
                var pRDs = pWS.OpenRasterDataset(NameExt);
                var pDS = pRDs as IDataset;
                pDS.Delete();
            }

            if (blockCount > 1)
            {
                #region 创建子目录，获得扩展名
                string NameNoExt = System.IO.Path.GetFileNameWithoutExtension(filename);
                string sExt = System.IO.Path.GetExtension(filename).ToLower();
                //创建子目录
                string subPath = filename.Substring(0, filename.LastIndexOf("."));
                if (System.IO.Directory.Exists(subPath))
                {
                    System.IO.Directory.Delete(subPath, true);
                }
                try
                {
                    System.IO.Directory.CreateDirectory(subPath);
                    if (!System.IO.Directory.Exists(subPath))
                    {
                        subPath = subPath + "_1";
                        System.IO.Directory.CreateDirectory(subPath);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    subPath = subPath + "_1";
                    System.IO.Directory.CreateDirectory(subPath);
                }

                //worldfile扩展名
                string worldfileExt = "." + sExt[1].ToString() + sExt[sExt.Length - 1].ToString() + "w";
                #endregion

                #region 分页输出
                int index = 0;
                int minX = 0, maxY = iHeight;
                double w, h = 0;
                string t_name = subPath + @"\" + NameNoExt + "_";
                IExport docExport = CreateExport(filename, 1);
                IEnvelope pEnv1 = new EnvelopeClass();
                while (h < height - 0.0001)
                {
                    w = 0;
                    minX = 0;
                    while (w < width - 0.0001)
                    {
                        pEnv1.XMin = w;
                        pEnv1.YMin = h;
                        pEnv1.XMax = w + paperWidth;
                        pEnv1.YMax = h + paperHeight;
                        index++;

                        label1.Text += ".";
                        Application.DoEvents();

                        //output输出
                        ActiveViewOutput(docActiveView, iDPI, iWidth, iHeight, pEnv1, t_name + index.ToString() + sExt, docExport);
                        //写入worldfile
                        WriteWorldfile(t_name + index.ToString() + worldfileExt, 1, 0, 0, -1, minX, maxY);
                        w += paperWidth;
                        minX += iWidth;
                    }
                    h += paperHeight;
                    maxY += iHeight;
                }
                #endregion

                #region 合并栅格
                //设置坐标参考
                var pRasterWS = OpenWorkspace(subPath, enumWsFactoryType.Raster);
                ISpatialReferenceFactory2 pSrF = new SpatialReferenceEnvironmentClass();
                var pSR = pSrF.CreateSpatialReference(3857);
                var pEnumDS = pRasterWS.get_Datasets(esriDatasetType.esriDTRasterDataset);
                var pDS = pEnumDS.Next();
                while (pDS != null)
                {
                    var GeoSchEdit = pDS as IGeoDatasetSchemaEdit;
                    if (GeoSchEdit.CanAlterSpatialReference)
                        GeoSchEdit.AlterSpatialReference(pSR);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pDS);
                    pDS = pEnumDS.Next();
                }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pRasterWS);


                //saveas时占用很大内存且不释放，使用GP工具
                //怎么不创建金字塔和头文件？？
                Geoprocessor geoprocessor = new Geoprocessor();
                try
                {
                    CreateRasterDataset createRD = new CreateRasterDataset();
                    createRD.cellsize = 1;
                    createRD.number_of_bands = 3;
                    createRD.out_path = filepath;
                    createRD.out_name = NameExt;
                    createRD.pyramids = "NONE";
                    createRD.compression = "NONE";
                    geoprocessor.Execute(createRD, null);

                    WorkspaceToRasterDataset MosaicToRaster = new WorkspaceToRasterDataset();
                    MosaicToRaster.in_workspace = subPath;
                    MosaicToRaster.in_raster_dataset = filename;
                    geoprocessor.Execute(MosaicToRaster, null);
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.Message);
                    for (int i = 0; i < geoprocessor.MessageCount; i++)
                    {
                        string abc = geoprocessor.GetMessage(i);
                        Console.WriteLine(abc);
                    }

                }
                #endregion

                return subPath;
            }
            else
            {
                Export10Plus(docActiveView, filename, iDPI, 0, 0, null);
                return "";
            }
        }
        else      //map
        {
            return "";
        }
    }

    private IExport CreateExport(string PathName, int iResampleRatio)
    {
        IExport docExport;
        string sExtionsName = PathName.Substring(PathName.LastIndexOf(".") + 1);
        sExtionsName = sExtionsName.ToUpper();
        switch (sExtionsName)
        {
            case "PDF":
                docExport = new ExportPDFClass();
                break;
            case "EPS":
                docExport = new ExportPSClass();
                break;
            case "AI":
                docExport = new ExportAIClass();
                break;
            case "BMP":
                docExport = new ExportBMPClass();
                break;
            case "TIFF":
                docExport = new ExportTIFFClass();
                break;
            case "TIF":
                docExport = new ExportTIFFClass();
                break;
            case "SVG":
                docExport = new ExportSVGClass();
                break;
            case "PNG":
                docExport = new ExportPNGClass();
                break;
            case "GIF":
                docExport = new ExportGIFClass();
                break;
            case "EMF":
                docExport = new ExportEMFClass();
                break;
            case "JPEG":
                docExport = new ExportJPEGClass();
                break;
            case "JPG":
                docExport = new ExportJPEGClass();
                break;
            default:
                docExport = new ExportJPEGClass();
                break;
        }
        if (docExport is IOutputRasterSettings)
        {
            if (iResampleRatio < 1) iResampleRatio = 1;
            if (iResampleRatio > 5) iResampleRatio = 5;
            IOutputRasterSettings RasterSettings = (IOutputRasterSettings)docExport;
            RasterSettings.ResampleRatio = iResampleRatio;
        }
        docExport.ExportFileName = PathName;
        return docExport;
    }

    //使用IActiveView.Output方法输出
    //Width和Hight用于设置输出的图片尺寸，默认为0则使用ActiveView.ExportFrame
    //VisibleBounds用于设置待输出的地图范围，默认为null则使用ActiveView.Extent
    private void ActiveViewOutput(IActiveView docActiveView, int iOutputResolution, int Width, int Height, IEnvelope VisibleBounds, string sExportFileName, IExport docExport)
    {
        docExport.ExportFileName = sExportFileName;
        if (iOutputResolution <= 0) iOutputResolution = 96;
        if (Width <= 0 || Height <= 0)
        {
            Width = docActiveView.ExportFrame.right;
            Height = docActiveView.ExportFrame.bottom;
            if (Width == 0)
                Width = 1024;
            if (Height == 0)
                Height = 768;
            Width = (Width * iOutputResolution) / 96;
            Height = (Height * iOutputResolution) / 96;
        }
        docExport.Resolution = iOutputResolution;
        ESRI.ArcGIS.esriSystem.tagRECT exportRECT = new ESRI.ArcGIS.esriSystem.tagRECT();
        exportRECT.left = 0;
        exportRECT.top = 0;
        exportRECT.right = Width;
        exportRECT.bottom = Height;

        IEnvelope envelope = new EnvelopeClass();
        envelope.PutCoords(exportRECT.left, exportRECT.top, exportRECT.right, exportRECT.bottom);
        docExport.PixelBounds = envelope;

        System.Int32 hDC = docExport.StartExporting();
        docActiveView.Output(hDC, iOutputResolution, ref exportRECT, VisibleBounds, null);
        docExport.FinishExporting();
        docExport.Cleanup();
    }

    //pagelayout不能设置图片大小，不能指定输出范围，由Page范围决定
    //map必须设置图片大小，结合dpi决定输出的比例尺
    private void Export10Plus(IActiveView docActiveView, string filename, int iDPI, int iWidth, int iHeight, IEnvelope pOutExtent)
    {
        IPrintAndExport docPrintExport = new PrintAndExportClass();
        IExport docExport = CreateExport(filename, 1);
        if (docActiveView is IMap)
        {
            int iWidth_96 = (int)(iWidth * (96.0 / (double)iDPI));
            int iHeight_96 = (int)(iHeight * (96.0 / (double)iDPI));
            tagRECT pDeviceFrame = new tagRECT();
            pDeviceFrame.top = 0; pDeviceFrame.left = 0; pDeviceFrame.right = iWidth_96; pDeviceFrame.bottom = iHeight_96;
            docActiveView.ScreenDisplay.DisplayTransformation.set_DeviceFrame(ref pDeviceFrame);
        }
        IEnvelope o_Env = null;
        if (pOutExtent != null)
        {
            o_Env = docActiveView.Extent.Envelope;
            docActiveView.Extent = pOutExtent;
        }
        docPrintExport.Export(docActiveView, docExport, iDPI, false, null);
        if (o_Env != null)
            docActiveView.Extent = o_Env;
        docExport.Cleanup();
    }

    private void WriteWorldfile(string pathname, double XCellsize, double XRotation, double YRotation, double negativeYCellsize, double XOffset, double YOffset)
    {
        System.IO.FileStream fs;
        StreamWriter sw;
        //写入worldfile
        fs = new System.IO.FileStream(pathname, FileMode.Create);
        sw = new StreamWriter(fs);
        sw.WriteLine(XCellsize);  //x cell size
        sw.WriteLine(XRotation);  //x旋转
        sw.WriteLine(YRotation);  //y旋转
        sw.WriteLine(negativeYCellsize);  //negative y cell size
        sw.WriteLine(XOffset);  //x平移,左上角x坐标
        sw.WriteLine(YOffset);//y平移，左上角y坐标
        //清空缓冲区    关闭流
        sw.Flush(); sw.Close(); fs.Close();
    }

    //自动判断文件名返回相应的workspace
    enum enumWsFactoryType
    {
        None,
        Shapefile,
        Raster,
        Mem,
        TextFile,
        Toolbox,
        Vpf,
        Cad
    }
    static IWorkspace OpenWorkspace(string connStr, enumWsFactoryType FactoryType)
    {
        string WsProgID = "";
        //先判断类型
        if (FactoryType == enumWsFactoryType.None)
        {
            #region 通过扩展名判断
            switch (System.IO.Path.GetExtension(connStr).ToLower())
            {
                case ".gdb":
                    WsProgID = "esriDataSourcesGDB.FileGDBWorkspaceFactory";
                    break;
                case ".mdb":
                    WsProgID = "esriDataSourcesGDB.AccessWorkspaceFactory";
                    break;
                case ".sde":
                    WsProgID = "esriDataSourcesGDB.SdeWorkspaceFactory";
                    break;
                case ".shp":
                    WsProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory";
                    connStr = connStr.Substring(0, connStr.LastIndexOf("\\"));
                    break;
                case ".tbx":
                    WsProgID = "esriGeoprocessing.ToolboxWorkspaceFactory";
                    connStr = connStr.Substring(0, connStr.LastIndexOf("\\"));
                    break;
                case ".txt":
                case ".csv":
                    WsProgID = "esriDataSourcesOleDB.TextFileWorkspaceFactory";
                    connStr = connStr.Substring(0, connStr.LastIndexOf("\\"));
                    break;
                case ".dbf":
                    WsProgID = "esriDataSourcesFile.VpfWorkspaceFactory";
                    connStr = connStr.Substring(0, connStr.LastIndexOf("\\"));
                    break;
                case ".nc":
                    WsProgID = "esriDataSourcesNetCDF.NetCDFWorkspaceFactory";
                    break;
                case ".dwg":
                case ".dxf":
                    WsProgID = "esriDataSourcesFile.CadWorkspaceFactory";
                    connStr = connStr.Substring(0, connStr.LastIndexOf("\\"));
                    break;
                case ".xls":
                case ".xlsx":
                    WsProgID = "esriDataSourcesOleDB.ExcelWorkspaceFactory";
                    break;
                case ".ai":
                case ".bmp":
                case ".emf":
                case ".exif":
                case ".eps":
                case ".gif":
                case ".jpg":
                case ".jpeg":
                case ".png":
                case ".pcx":
                case ".pdf":
                case ".tif":
                case ".tiff":
                case ".svg":
                case ".wmf":
                    WsProgID = "esriDataSourcesRaster.RasterWorkspaceFactory";
                    connStr = connStr.Substring(0, connStr.LastIndexOf("\\"));
                    break;
                default:
                    return null;
            }
            #endregion
        }
        else
        {
            switch (FactoryType)
            {
                case enumWsFactoryType.Mem:
                    WsProgID = "esriDataSourcesGDB.InMemoryWorkspaceFactory";
                    IWorkspaceName workspaceName = GetWorkspaceFactory(WsProgID).Create(null, "MyWorkspace", null, 0);
                    IName name = (IName)workspaceName;
                    IWorkspace workspace = (IWorkspace)name.Open();
                    return workspace;
                    break;
                case enumWsFactoryType.Cad:
                    WsProgID = "esriDataSourcesFile.CadWorkspaceFactory";
                    break;
                case enumWsFactoryType.Raster:
                    WsProgID = "esriDataSourcesRaster.RasterWorkspaceFactory";
                    break;
                case enumWsFactoryType.Shapefile:
                    WsProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory";
                    break;
                case enumWsFactoryType.TextFile:
                    WsProgID = "esriDataSourcesOleDB.TextFileWorkspaceFactory";
                    break;
                case enumWsFactoryType.Toolbox:
                    WsProgID = "esriGeoprocessing.ToolboxWorkspaceFactory";
                    break;
                case enumWsFactoryType.Vpf:
                    WsProgID = "esriDataSourcesFile.VpfWorkspaceFactory";
                    break;
                default:
                    return null;
                    break;
            }
            if (System.IO.Path.GetExtension(connStr) != "")
                connStr = connStr.Substring(0, connStr.LastIndexOf("\\"));
        }
        var workspaceFactory = GetWorkspaceFactory(WsProgID);
        return workspaceFactory.OpenFromFile(connStr, 0);
    }
    static IWorkspaceFactory2 GetWorkspaceFactory(string WsProgID)
    {
        Type t = Type.GetTypeFromProgID(WsProgID);
        System.Object obj = Activator.CreateInstance(t);
        return obj as IWorkspaceFactory2;
    }

  }
}



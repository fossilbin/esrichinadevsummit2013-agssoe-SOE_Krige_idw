using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Collections.Specialized;

using System.Runtime.InteropServices;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Server;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.SOESupport;
using ESRI.ArcGIS.GeoAnalyst;
using ESRI.ArcGIS.DataSourcesRaster;


//TODO: sign the project (project properties > signing tab > sign the assembly)
//      this is strongly suggested if the dll will be registered using regasm.exe <your>.dll /codebase


namespace Krige_idw
{
    [ComVisible(true)]
    [Guid("f38c90c8-66da-4b1e-90f1-68072325f866")]
    [ClassInterface(ClassInterfaceType.None)]
    [ServerObjectExtension("MapServer",
        AllCapabilities = "",
        DefaultCapabilities = "",
        Description = "This SOE was used to interpolation",
        DisplayName = "Krige_idw",
        //设置SOE的属性
        Properties = "Field_Name=spot;Layer_Name=point elevation",
        SupportsREST = true,
        SupportsSOAP = false)]
    public class Krige_idw : IServerObjectExtension, IObjectConstruct, IRESTRequestHandler
    {
        private string soe_name;

        private IPropertySet configProps;
        private IServerObjectHelper serverObjectHelper;
        private ServerLogger logger;
        private IRESTRequestHandler reqHandler;
        private IFeatureClass _featureClass = null;
        private string m_mapLayerNameToQuery = ""; // 查询图层名
        private string m_mapFieldToQuery = "";//查询字段
        public Krige_idw()
        {
            soe_name = this.GetType().Name;
            logger = new ServerLogger();
            reqHandler = new SoeRestImpl(soe_name, CreateRestSchema()) as IRESTRequestHandler;
        }

        #region IServerObjectExtension Members

        public void Init(IServerObjectHelper pSOH)
        {
            //生命周期开始时调试
            System.Diagnostics.Debugger.Launch();
            serverObjectHelper = pSOH;

        }

        public void Shutdown()
        {

            logger.LogMessage(ServerLogger.msgType.infoStandard, "Shutdown", 8000, "Custom message: Shutting down the SOE");
            soe_name = null;
            _featureClass = null;
            m_mapFieldToQuery = null;
            serverObjectHelper = null;
            logger = null;
        }

        #endregion

        #region IObjectConstruct Members

        public void Construct(IPropertySet props)
        {
            configProps = props;


            if (props.GetProperty("Field_Name") != null)
            {
                m_mapFieldToQuery = props.GetProperty("Field_Name") as string;
            }
            else
            {
                throw new ArgumentNullException();
            }
            if (props.GetProperty("Layer_Name") != null)
            {
                m_mapLayerNameToQuery = props.GetProperty("Layer_Name") as string;
            }
            else
            {
                throw new ArgumentNullException();
            }
            try
            {
                //获取数据              
                IMapServer3 mapServer = (IMapServer3)serverObjectHelper.ServerObject;
                string mapName = mapServer.DefaultMapName;
                IMapLayerInfo layerInfo;
                IMapLayerInfos layerInfos = mapServer.GetServerInfo(mapName).MapLayerInfos;
                // 获取查询图层id
                int layercount = layerInfos.Count;
                int layerIndex = 0;
                for (int i = 0; i < layercount; i++)
                {
                    layerInfo = layerInfos.get_Element(i);
                    if (layerInfo.Name == m_mapLayerNameToQuery)
                    {
                        layerIndex = i;
                        break;
                    }
                }

                IMapServerDataAccess dataAccess = (IMapServerDataAccess)mapServer;
                //获取要素          

                _featureClass = (IFeatureClass)dataAccess.GetDataSource(mapName, layerIndex);
                //确保获取到要素
                if (_featureClass == null)
                {
                    logger.LogMessage(ServerLogger.msgType.error, "Construct", 8000, "SOE custom error: Layer name not found.");
                    return;
                }
                // soe插值字段
                if (_featureClass.FindField(m_mapFieldToQuery) == -1)
                {
                    logger.LogMessage(ServerLogger.msgType.error, "Construct", 8000, "SOE custom error: Field not found in layer.");
                }
            }
            catch
            {

            }
        } 

        #endregion

        #region IRESTRequestHandler Members

        public string GetSchema()
        {
            return reqHandler.GetSchema();
        }

        public byte[] HandleRESTRequest(string Capabilities, string resourceName, string operationName, string operationInput, string outputFormat, string requestProperties, out string responseProperties)
        {
            return reqHandler.HandleRESTRequest(Capabilities, resourceName, operationName, operationInput, outputFormat, requestProperties, out responseProperties);
        }

        #endregion

        private RestResource CreateRestSchema()
        {
            RestResource rootRes = new RestResource(soe_name, false, RootResHandler);

            RestOperation sampleOper = new RestOperation("Interpolation",
                                                      new string[] { "method", "cellsize" },
                                                      new string[] { "json" },
                                                      DoInteroplatHandler);
            rootRes.operations.Add(sampleOper);
            RestResource propertiesResource = new RestResource("properties", false, PropertiesResHandler);
            rootRes.resources.Add(propertiesResource);
          

            return rootRes;
        }

        private byte[] RootResHandler(NameValueCollection boundVariables, string outputFormat, string requestProperties, out string responseProperties)
        {
            responseProperties = null;
            JSONObject result = new JSONObject();
            result.AddString("名称", "高程插值SOE");
            result.AddString("描述", "通过改程序可以对制定图层中的制定字段采用Krige或者IDW方法进行插值，并将插值结果返回");
            result.AddString("方法", "通过输入方法名Krige或者IDW方法和栅格单元大小进行插值");

            return Encoding.UTF8.GetBytes(result.ToJSONString(null));
        }

        private byte[] PropertiesResHandler(NameValueCollection boundVariables, string outputFormat, string requestProperties, out string responseProperties)
        {
            responseProperties = "{\"Content-Type\" : \"application/json\"}";

            JsonObject result = new JsonObject();
            result.AddString("Field_Name", this.m_mapFieldToQuery);
            result.AddString("Layer_Name", this.m_mapLayerNameToQuery);
           

            return Encoding.UTF8.GetBytes(result.ToJson());

        }
        private byte[] DoInteroplatHandler(NameValueCollection boundVariables,
                                                  JsonObject operationInput,
                                                      string outputFormat,
                                                      string requestProperties,
                                                  out string responseProperties)
        {
            responseProperties = null;
            // 序列化插值方法
            string method;
            if (!operationInput.TryGetString("method", out method))
                throw new ArgumentNullException("invalid interpolation method", "method");
            //string inputfield;
            //if(!operationInput.TryGetString("插值字段",out inputfield))
            //    throw new ArgumentNullException("无效的插值字段","inputfield");
            // 序列化像元大小
            double? cellsize;
            if (!operationInput.TryGetAsDouble("cellsize", out cellsize) || !cellsize.HasValue)
                throw new ArgumentException("invalid cell sieze", "cellsize");
          

            IRasterLayer pRasterLayer = DoInteroplate(_featureClass, m_mapFieldToQuery, cellsize, method);
            //构造转换矢量的文件名
            string name = "result_" + System.DateTime.Now.ToString().Replace("/","").Replace(":","").Replace(" ","")+".shp";
            IFeatureClass pFeatureClass = Raster2Polygon(pRasterLayer.Raster, name);
            byte[] result = GetResultJson(pFeatureClass);
            return result;
        }

        private byte[] GetResultJson(IFeatureClass inputFeaClass)
        {
            //查询出所有要素
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = null;
            int count = inputFeaClass.FeatureCount(pQueryFilter);

            //将每一个要素序列化成json数据
            IFeature pFeature = null;
            List<JsonObject> jsonGeometries = new List<JsonObject>();
            for (int i = 1; i < count; i++)//OBJECTID从1开始
            {
                pFeature = inputFeaClass.GetFeature(i);
                IGeometry pGeometry = pFeature.Shape;
                JsonObject featureJson = new JsonObject();
                //id号   
                featureJson.AddLong("id", i);
                //分类等级
                double grid_Code = (double)pFeature.get_Value(pFeature.Fields.FindField("GRIDCODE"));
                featureJson.AddDouble("gridCode", grid_Code);//栅格分类的等级
                JsonObject feaGeoJson = null;//几何对象   
                if (pGeometry != null)
                {
                    feaGeoJson = Conversion.ToJsonObject(pGeometry);
                    featureJson.AddJsonObject("geometry", feaGeoJson);//加入几何对象
                }




                jsonGeometries.Add(featureJson);
            }

            JsonObject resultJson = new JsonObject();
            resultJson.AddArray("geometries", jsonGeometries.ToArray());
            byte[] result = Encoding.UTF8.GetBytes(resultJson.ToJson());
            return result;
        }


        private IRasterLayer DoInteroplate(IFeatureClass inputFclass, string inputField, object cellSize, string method)
        {
            if (inputFclass == null)
            {
                logger.LogMessage(ServerLogger.msgType.error, "SOE构建错误", 8000, "无效的的FeatureClass输入.");

            }
            if (string.IsNullOrEmpty(inputField))
            {
                logger.LogMessage(ServerLogger.msgType.error, "SOE构建错误", 8000, "无效的字段输入.");
            }
            if (cellSize == null)
            {

                logger.LogMessage(ServerLogger.msgType.error, "SOE构建错误", 8000, "无效的栅格单元大小.");

            }
            if (string.IsNullOrEmpty(method))
            {

                logger.LogMessage(ServerLogger.msgType.error, "SOE构建错误", 8000, "无效的插值方法.");

            }


            //生成IFeatureClassDescriptor
            IFeatureClassDescriptor pFcDescriptor = new FeatureClassDescriptorClass();
            pFcDescriptor.Create(inputFclass, null, inputField);
            //设置分析环境
            IInterpolationOp pInterpolationOp = new RasterInterpolationOpClass();
            IRasterAnalysisEnvironment pEnv = pInterpolationOp as IRasterAnalysisEnvironment;
            pEnv.Reset();
            //栅格单元大小
            pEnv.SetCellSize(esriRasterEnvSettingEnum.esriRasterEnvValue, ref cellSize);
            //定义搜索半径,可变搜索半径
            IRasterRadius pRadius = new RasterRadius();
            object missing = Type.Missing;
            pRadius.SetVariable(12, ref missing);
            IRaster pRaster = null;

            //执行插值
            method = method.ToLower();
            switch (method)
            {
                case "krige":
                    IGeoDataset pGeoDataset = pInterpolationOp.Krige((IGeoDataset)pFcDescriptor, esriGeoAnalysisSemiVariogramEnum.esriGeoAnalysisSphericalSemiVariogram,
                                    pRadius, false, ref missing);
                    pRaster = pGeoDataset as IRaster;
                    break;
                case "idw":

                    pRaster = pInterpolationOp.IDW(pFcDescriptor as IGeoDataset, 2, pRadius, ref missing) as IRaster;
                    break;
            }

            IRasterLayer pRasterLayer = new RasterLayerClass();
            //对插值结果进行重分类，采用等间距分为4类
            pRaster = DoReClassify(pRaster, 4);
            pRasterLayer.CreateFromRaster(pRaster);
            pRasterLayer.Name = "插值结果";
            return pRasterLayer;

        }

        private IRaster DoReClassify(IRaster inputRaster, int pClassNo)
        {

            //获取栅格分类数组和频度数组
            object dataValues = null, dataCounts = null;
            GetRasterClass(inputRaster, out dataValues, out dataCounts);

            //获取栅格分类间隔数组
            IClassifyGEN pEqualIntervalClass = new EqualIntervalClass();
            pEqualIntervalClass.Classify(dataValues, dataCounts, ref pClassNo);
            double[] breaks = pEqualIntervalClass.ClassBreaks as double[];

            //设置新分类值
            INumberRemap pNemRemap = new NumberRemapClass();
            for (int i = 0; i < breaks.Length - 1; i++)
            {
                pNemRemap.MapRange(breaks[i], breaks[i + 1], i + 1);
            }
            IRemap pRemap = pNemRemap as IRemap;

            //设置环境
            IReclassOp pReclassOp = new RasterReclassOpClass();
            IGeoDataset pGeodataset = inputRaster as IGeoDataset;
            IRasterAnalysisEnvironment pEnv = pReclassOp as IRasterAnalysisEnvironment;
            object obj = Type.Missing; IEnvelope pRasterExt = new EnvelopeClass();
            //重分类      
            IRaster pRaster = pReclassOp.ReclassByRemap(pGeodataset, pRemap, false) as IRaster;

            dataValues = null;
            dataCounts = null;
            pEqualIntervalClass = null;
            breaks = null;
            pNemRemap = null;
            pRemap = null;
            pReclassOp = null;
            pGeodataset = null;
            pRasterExt = null;

            return pRaster;
        }
        //计算栅格重分类值
        private void GetRasterClass(IRaster inputRaster, out object dataValues, out object dataCounts)
        {
            IRasterBandCollection pRasBandCol = inputRaster as IRasterBandCollection;
            IRasterBand pRsBand = pRasBandCol.Item(0);
            pRsBand.ComputeStatsAndHist();
            //IRasterBand中本无统计直方图，必须先进行ComputeStatsAndHist()
            IRasterStatistics pRasterStatistic = pRsBand.Statistics;

            double mMean = pRasterStatistic.Mean;
            double mStandsrdDeviation = pRasterStatistic.StandardDeviation;

            IRasterHistogram pRasterHistogram = pRsBand.Histogram;
            double[] dblValues;
            dblValues = pRasterHistogram.Counts as double[];
            int intValueCount = dblValues.GetUpperBound(0) + 1;
            double[] vValues = new double[intValueCount];

            double dMaxValue = pRasterStatistic.Maximum;
            double dMinValue = pRasterStatistic.Minimum;
            double BinInterval = Convert.ToDouble((dMaxValue - dMinValue) / intValueCount);
            for (int i = 0; i < intValueCount; i++)
            {
                vValues[i] = i * BinInterval + pRasterStatistic.Minimum;
            }

            dataValues = vValues as object;
            dataCounts = dblValues as object;
        }

        //栅格转换成矢量
        private IFeatureClass Raster2Polygon(IRaster inputRaster, string filename)
        {
            //获得加载数据的空间

            IRasterBandCollection pRasBandCol = inputRaster as IRasterBandCollection;
            IRasterBand pRsBand = pRasBandCol.Item(0);
            IRasterDataset pRasterDataset = pRsBand as IRasterDataset;
            IGeoDataset pRasterGeoDataset = pRasterDataset as IGeoDataset;

            //栅格转面
            IConversionOp pConversionOp = new RasterConversionOpClass();
            ISpatialReference pSpatialReference = pRasterGeoDataset.SpatialReference;

            //转换
            IWorkspace _pWorkspace = (_featureClass as IDataset).Workspace;
            //判断文件夹目录下是否存在文件，如果就删除文件
            string filepath = _pWorkspace.PathName;
           
            string filefullpath = System.IO.Path.Combine(filepath, filename);
            if (System.IO.File.Exists(filefullpath))
            {
                string[] tmp = System.IO.Directory.GetFiles(filepath, System.IO.Path.GetFileNameWithoutExtension(filefullpath) + ".*");
                foreach (string item in tmp)
                {

                    System.IO.File.Delete(item);

                }
            }
            IGeoDataset pGeoDataset = pConversionOp.RasterDataToPolygonFeatureData(pRasterGeoDataset, _pWorkspace, filename, false);
            return pGeoDataset as IFeatureClass;
        }

    }
}

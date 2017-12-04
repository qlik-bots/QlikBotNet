using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using NinjaNye.SearchExtensions;

using Qlik.Engine;
using Qlik.Sense.Client;
using Qlik.Sense.Client.Visualizations;
using Qlik.Sense.Client.Visualizations.Components;

using System.Security.Cryptography.X509Certificates;
using System.IO;


namespace QlikSenseEasy
{
    public class QSFoundObject
    {
        public string Description { get; set; }
        public string ObjectId { get; set; }
        public string ObjectURL { get; set; }
        public string ObjectType { get; set; }

        public string HRef { get; set; }
        public string ThumbURL { get; set; }
    }

    public class QSFoundFilter
    {
        public string Dimension { get; set; }
        public string Element { get; set; }
    }

    public class QSValueFilter
    {
        public string MeasureExpression { get; set; }
        public string MeasureCondition { get; set; }
    }

    public class QSDataList
    {
        public string DimValue { get; set; }
        public double MeasValue { get; set; }
        public string MeasFormattedValue { get; set; }
    }

    public class QSGeoFilter
    {
        public double FromLatitude { get; set; }
        public double FromLongitude { get; set; }
        public string LatFilter { get; set; }
        public string LonFilter { get; set; }
    }

    public class QSGeoList
    {
        public double Lat { get; set; }
        public double Lon { get; set; }
        public string Address { get; set; }
        public string TextLabel { get; set; }
        public string Text { get; set; }
        public string ValueLabel { get; set; }
        public double Value { get; set; }
        public string FormattedValue { get; set; }
    }


    public class QSGasGeoList
    {
        public double Lat { get; set; }
        public double Lon { get; set; }
        public string Text { get; set; }
        public string Address { get; set; }
        public double Price { get; set; }
    }

    public class QSAlert
    {
        public string UserName { get; set; }
        public string UserID { get; set; }
        public string SendToID { get; set; }
        public string AlertRequest { get; set; }
        public string AlertMessage { get; set; }
        public string AlertPhoto { get; set; }
        public string AlertDoc { get; set; }
        public bool AlertActive { get; set; }
        public bool AlertSent { get; set; }

        public QSAlert()
        {
            AlertActive = false;
        }
    }

    public class QSField
    {
        public string FieldNameInApp { get; set; }
        public string FieldToSearch { get; set; }
        public IEnumerable<string> Tags { get; set; }
    }

    public class QSMasterItem
    {
        public string Name { get; set; }
        public string Expression { get; set; }
        public string Id { get; set; }
        public string FormattedExpression { get; set; }
        public IEnumerable<string> Tags { get; set; }

    }

    public class QSStory
    {
        public string Name { get; set; }
        public string Id { get; set; }
        //public string FirstSlideId { get; set; }
    }

    public class QSSheet
    {
        public string Name { get; set; }
        public string Id { get; set; }
    }


    public class QSAppProperties
    {
        public DateTime modifiedDate { get; set; }
        public bool published { get; set; }
        public DateTime publishTime { get; set; }
        public IList<string> privileges { get; set; }
        public string description { get; set; }
        public int qFileSize { get; set; }
        public string dynamicColor { get; set; }
        public object create { get; set; }
        public QSStreamProperties stream { get; set; }
        public bool canCreateDataConnections { get; set; }
        public string AppID { get; set; }
        public string AppName { get; set; }
        public string AppTitle { get; set; }
        public string ThumbnailUrl { get; set; }
    }

    public class QSStreamProperties
    {
        public string id { get; set; }
        public string name { get; set; }
    }




    public class QSApp
    {
        public string qsServer { get; set; }
        private string QSHeaderAuthName;
        public string qsAppName { get; set; }
        public string qsAppId { get; set; }
        public string qsAppThumbnailUrl { get; set; }
        public string qsSingleServer { get; set; }
        public string qsSingleApp { get; set; }

        public List<QSAlert> AlertList;

        private ILocation qsLocation;
        private IApp qsApp;

        const int maxFounds = 10;
        private static Random rnd;

        private List<QSField> AppFields = new List<QSField>();

        private QSField LatitudeField;
        private QSField LongitudeField;
        private QSField AddressField;
        private bool GeoLocationActive = false;

        public List<QSMasterItem> MasterMeasures = new List<QSMasterItem>();
        public List<QSMasterItem> MasterDimensions = new List<QSMasterItem>();
        public List<QSMasterItem> MasterVisualizations = new List<QSMasterItem>();

        public List<QSStory> Stories = new List<QSStory>();

        public List<QSSheet> Sheets = new List<QSSheet>();

        public List<QSAppProperties> qsAlternativeApps = new List<QSAppProperties>();
        public string qsAlternativeStreams = null;

        public QSMasterItem LastMeasure;
        public QSMasterItem LastDimension;
        public List<QSFoundFilter> LastFilters;

        public List<QSValueFilter> ValueFilters;

        private string QSUserId;

        private bool QSIsConnected = false;
        public bool IsConnected { get { return QSIsConnected; } }


        private bool QSAppIsOpen = false;
        public bool AppIsOpen { get { return QSAppIsOpen; } }



        public bool IsGeoLocationActive()
        {
            return GeoLocationActive;
        }

        public QSFoundFilter AddFilter(string Dimension, string Element)
        {
            QSFoundFilter ff = new QSFoundFilter();
            ff.Dimension = Dimension;
            ff.Element = Element;

            QSFoundFilter Repeated = LastFilters.Find(f => f.Dimension == Dimension);
            if (Repeated != null)
            {
                Repeated.Element = Element;
                return Repeated;
            }
            else
            {
                LastFilters.Add(ff);
                return ff;
            }
        }

        public QSFoundFilter AddFilter(QSFoundFilter Filter)
        {
            return AddFilter(Filter.Dimension, Filter.Element);
        }

        public QSValueFilter AddValueFilter(string MeasureExpression, string MeasureCondition)
        {
            QSValueFilter vf = new QSValueFilter();
            vf.MeasureExpression = MeasureExpression;
            vf.MeasureCondition = MeasureCondition;

            QSValueFilter Repeated = ValueFilters.Find(f => f.MeasureExpression == MeasureExpression);
            if (Repeated != null)
            {
                Repeated.MeasureCondition = MeasureCondition;
                return Repeated;
            }
            else
            {
                ValueFilters.Add(vf);
                return vf;
            }
        }

        public QSValueFilter AddValueFilter(QSValueFilter ValFilter)
        {
            return AddValueFilter(ValFilter.MeasureExpression, ValFilter.MeasureCondition);
        }


        public QSApp()
        {
            qsServer = "https://myserver.com";
            qsAppName = "Executive Dashboard";
            qsAppId = "";
            qsSingleServer = "https://myserver.com";
            qsSingleApp = "gvgvgvgvgv";

            AlertList = new List<QSAlert>();
            LastFilters = new List<QSFoundFilter>();
            ValueFilters = new List<QSValueFilter>();
            rnd = new Random();
        }

        public void CheckConnection()
        {
            if (!qsLocation.IsAlive())
            {
                this.QSConnectServerHeader(QSUserId, QSHeaderAuthName, qsLocation.VirtualProxyPath, IsUsingSSL(), IsCheckingSDKVersion());
            }
        }


        internal static byte[] ReadFile(string fileName)
        {
            FileStream f = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            int size = (int)f.Length;
            byte[] data = new byte[size];
            size = f.Read(data, 0, size);
            f.Close();
            return data;
        }

        public void QSConnectServerHeader(string UserId, string HeaderAuthName, string VirtualProxyPath = "",
            Boolean UseSSL = false, Boolean CheckSDKVersion = true)
        {
            try
            {
                QSIsConnected = false;
                string strUri = qsServer;
                Uri uri = new Uri(strUri);

                qsLocation = Qlik.Engine.Location.FromUri(uri);
                if (VirtualProxyPath.Trim() != "") qsLocation.VirtualProxyPath = VirtualProxyPath;

                qsLocation.AsStaticHeaderUserViaProxy(UserId, HeaderAuthName, UseSSL);

                qsLocation.IsVersionCheckActive = CheckSDKVersion;
                IHub MyHub = qsLocation.Hub();

                QSUserId = UserId;
                QSHeaderAuthName = HeaderAuthName;

                QSIsConnected = true;

                Console.WriteLine("QSEasy connected to Qlik Sense version: " + MyHub.ProductVersion());
                Console.WriteLine("UserID: " + UserId + " - VirtualProxy: " + VirtualProxyPath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }


        public bool IsCheckingSDKVersion()
        {
            return qsLocation.IsVersionCheckActive;
        }

        public bool IsUsingSSL()
        {
            return (qsLocation.ServerUri.Scheme == "https");
        }

        public string VirtualProxy()
        {
            if (qsLocation.VirtualProxyPath == null || qsLocation.VirtualProxyPath.Trim() == "")
                return "";
            else
                return qsLocation.VirtualProxyPath;
        }

        public void QSOpenApp()
        {
            CheckConnection();

            IAppIdentifier MyAppId;
            if (qsAppId != "" && qsAppId != null)
            {
                MyAppId = qsLocation.AppWithIdOrDefault(qsAppId);
            }
            else
            {
                MyAppId = qsLocation.AppWithNameOrDefault(qsAppName);
            }

            qsAppId = MyAppId.AppId;
            qsAppName = MyAppId.AppName;

            qsApp = qsLocation.App(MyAppId);
            qsAppThumbnailUrl = qsApp.GetAppProperties().Thumbnail.Url;

            QSReadFields();
            QSReadMasterItems();
            QSReadSheets();
            QSReadStories();
            LastFilters.Clear();
            ValueFilters.Clear();
            GetAlternativeApps(qsAlternativeStreams);

            Console.WriteLine("QSEasy opened App " + (qsAppId != "" ? qsAppId : qsAppName) + " with handle: " + qsApp.Handle);
        }

        private void QSOpenApp(QSAppProperties App)
        {
            qsAppName = App.AppName;
            qsAppId = App.AppID;
            qsSingleApp = App.AppID;
            qsAppThumbnailUrl = App.ThumbnailUrl;

            QSOpenApp();
        }

        public void QSOpenApp(string AppId)
        {
            GetAlternativeApps(qsAlternativeStreams);
            QSAppProperties AppProp = new QSAppProperties();
            AppProp.AppID = AppId;
            QSOpenApp(AppProp);
        }

        private void GetAlternativeApps(string StreamNames = null)
        {
            List<string> Streams;
            if (StreamNames != null)
                Streams = StreamNames.Split(';').ToList();
            else
                Streams = new List<string>();

            qsAlternativeApps.Clear();

            foreach (IAppIdentifier App in qsLocation.GetAppIdentifiers())
            {
                QSAppProperties AppProp = new QSAppProperties();
                AppProp = Newtonsoft.Json.JsonConvert.DeserializeObject<QSAppProperties>(App.Meta.PrintStructure());
                AppProp.AppID = App.AppId;
                AppProp.AppName = App.AppName;
                AppProp.AppTitle = App.Title;
                AppProp.ThumbnailUrl = App.Thumbnail.Url;

                if (AppProp.published && (Streams.Count == 0 || Streams.Contains(AppProp.stream?.name)))
                    qsAlternativeApps.Add(AppProp);
                if (!AppProp.published && AppProp.AppName == QSUserId)
                    qsAlternativeApps.Add(AppProp);
            }
        }


        private void QSReadSheets()
        {
            Sheets.Clear();
            foreach (Qlik.Sense.Client.ISheet AppSheet in Qlik.Sense.Client.AppExtensions.GetSheets(qsApp))
            {
                QSSheet qss = new QSSheet();
                qss.Id = AppSheet.Id;
                qss.Name = AppSheet.Properties.MetaDef.Title;
                var m = AppSheet.Properties.MetaDef;
                Sheets.Add(qss);
            }
        }

        public void QSDoSave()
        {
            qsApp.DoSave();
        }

        public bool QSRemoveSheet(string SheetID)
        {
            if (SheetID != null && SheetID != "")
                return qsApp.RemoveSheet(SheetID);
            else
                return false;
        }

        public string QSCreateSheet(string SheetID, string Title = "", string Description = "")
        {
            QSRemoveSheet(SheetID);
            QSDoSave();

            SheetProperties sp = new SheetProperties();
            if (Title != "")
                sp.MetaDef.Title = Title;
            if (Description != "")
                sp.MetaDef.Description = Description;
            sp.Rank = 0;

            ISheet Sheet = qsApp.CreateSheet(SheetID, sp);
            QSDoSave();

            return SheetID;
        }


        public void QSCreateTextImage(string SheetID, string Id, string Title, string Text)
        {
            ISheet Sheet = qsApp.GetSheet(SheetID);
            ITextImage TextImg = Sheet.CreateTextImage(Id, new TextImageProperties { Title = Title, Markdown = Text });
            //return TextImg;
        }

        public void QSCreateKPI(string SheetID, string Id, string Measure1, string Measure2 = null,
            string ChartTitle = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            IKpi Kpi = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem mm2 = GetMasterMeasure(Measure2);

                var properties = new KpiProperties
                {
                    Title = ChartTitle,
                    Subtitle = mm1.Name,
                    ShowTitles = ChartTitle == "" ? false : true,
                    HyperCubeDef = new KpiVisualizationHyperCubeDef
                    {
                        SuppressMissing = true,
                        InterColumnSortOrder = new[] { 0, 1 },
                        InitialDataFetch = new List<NxPage>
                        {
                            new NxPage
                            {
                                Height = 500,
                                Left = 0,
                                Top = 0,
                                Width = 10
                            }
                        }.ToArray(),
                        Measures = new List<KpiHyperCubeMeasureDef>
                        {
                            new KpiHyperCubeMeasureDef
                            {
                                Def =
                                    new KpiHyperCubeMeasureqDef
                                    {
                                        Def = mm1.Expression,
                                        Label = mm1.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            },
                            new KpiHyperCubeMeasureDef
                            {
                                Def =
                                    new KpiHyperCubeMeasureqDef
                                    {
                                        Def = mm2.Expression,
                                        Label = mm2.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            }
                        }
                    }
                };

                Kpi = Sheet.CreateKpi(Id, properties);

                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var kpiCell = Sheet.CellFor(Kpi);
                    //kpi1Cell.SetBounds(1, 1, 3, 3);
                    kpiCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Kpi = null;
            }

        }

        public void QSCreateBarChart(string SheetID, string Id, string Dimension1, string Measure1,
            string ChartTitle = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            IBarchart BChart = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem md1 = GetMasterDimension(Dimension1);

                var properties = new BarchartProperties()
                {
                    Title = ChartTitle == "" ? Id : ChartTitle,
                    ShowTitles = ChartTitle == "" ? false : true,
                    HyperCubeDef = new VisualizationHyperCubeDef
                    {
                        Dimensions = new List<HyperCubeDimensionDef>
                        {
                            new HyperCubeDimensionDef
                            {
                                Def = new HyperCubeDimensionqDef
                                {
                                    CId = ClientExtension.GetCid(),
                                    FieldLabels = new[] {md1.Name},
                                    FieldDefs = new[] {md1.Expression}
                                }
                            }
                        },
                        Measures = new List<HyperCubeMeasureDef>
                        {
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm1.Expression,
                                        Label = mm1.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            }
                        },
                        //InterColumnSortOrder = new[] { 1, 0 },
                        InitialDataFetch =
                            new List<NxPage>
                            {
                                new NxPage { Height = 500, Left = 0, Top = 0, Width = 10 }
                            }.ToArray()
                    }
                };
                BChart = Sheet.CreateBarchart(Id, properties);

                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var barchartCell = Sheet.CellFor(BChart);
                    barchartCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                BChart = null;
            }
            //return BChart;
        }

        public void QSCreateLineChart(string SheetID, string Id, string Dimension1, string Measure1,
           string ChartTitle = "",
           int col = 0, int row = 0, int width = 0, int height = 0)
        {
            ILinechart LChart = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem md1 = GetMasterDimension(Dimension1);

                var properties = new LinechartProperties()
                {
                    Title = ChartTitle == "" ? Id : ChartTitle,
                    ShowTitles = ChartTitle == "" ? false : true,
                    HyperCubeDef = new VisualizationHyperCubeDef
                    {
                        Dimensions = new List<HyperCubeDimensionDef>
                        {
                            new HyperCubeDimensionDef
                            {
                                Def = new HyperCubeDimensionqDef
                                {
                                    CId = ClientExtension.GetCid(),
                                    FieldLabels = new[] {md1.Name},
                                    FieldDefs = new[] {md1.Expression}
                                }
                            }
                        },
                        Measures = new List<HyperCubeMeasureDef>
                        {
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm1.Expression,
                                        Label = mm1.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            }
                        },
                        //InterColumnSortOrder = new[] { 1, 0 },
                        InitialDataFetch =
                            new List<NxPage>
                            {
                                new NxPage { Height = 500, Left = 0, Top = 0, Width = 10 }
                            }.ToArray()
                    }
                };
                LChart = Sheet.CreateLinechart(Id, properties);


                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var linechartCell = Sheet.CellFor(LChart);
                    linechartCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                LChart = null;
            }
            //return LChart;
        }

        public void QSCreatePieChart(string SheetID, string Id, string Dimension1, string Measure1,
            string ChartTitle = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            IPiechart PChart = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem md1 = GetMasterDimension(Dimension1);

                var properties = new PiechartProperties()
                {
                    Title = ChartTitle == "" ? Id : ChartTitle,
                    ShowTitles = ChartTitle == "" ? false : true,
                    HyperCubeDef = new VisualizationHyperCubeDef
                    {
                        Dimensions = new List<HyperCubeDimensionDef>
                        {
                            new HyperCubeDimensionDef
                            {
                                Def = new HyperCubeDimensionqDef
                                {
                                    CId = ClientExtension.GetCid(),
                                    FieldLabels = new[] {md1.Name},
                                    FieldDefs = new[] {md1.Expression}
                                }
                            }
                        },
                        Measures = new List<HyperCubeMeasureDef>
                        {
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm1.Expression,
                                        Label = mm1.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            }
                        },
                        //InterColumnSortOrder = new[] { 1, 0 },
                        InitialDataFetch =
                            new List<NxPage>
                            {
                                new NxPage { Height = 500, Left = 0, Top = 0, Width = 10 }
                            }.ToArray()
                    }
                };
                PChart = Sheet.CreatePiechart(Id, properties);


                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var piechartCell = Sheet.CellFor(PChart);
                    piechartCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                PChart = null;
            }
            //return PChart;
        }

        public void QSCreateTreeChart(string SheetID, string Id, string Dimension1, string Measure1,
            string ChartTitle = "", string Dimension2 = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            ITreemap TChart = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem md1 = GetMasterDimension(Dimension1);
                QSMasterItem md2 = GetMasterDimension(Dimension2);

                if (Dimension2.Length > 0)
                {
                    var properties = new TreemapProperties()
                    {
                        Title = ChartTitle == "" ? Id : ChartTitle,
                        ShowTitles = ChartTitle == "" ? false : true,
                        HyperCubeDef = new VisualizationHyperCubeDef
                        {
                            Mode = NxHypercubeMode.DATA_MODE_PIVOT_STACK,
                            Dimensions = new List<HyperCubeDimensionDef>
                            {
                                new HyperCubeDimensionDef
                                {
                                    Def = new HyperCubeDimensionqDef
                                    {
                                        CId = ClientExtension.GetCid(),
                                        FieldLabels = new[] {md1.Name},
                                        FieldDefs = new[] {md1.Expression}
                                    }
                                },
                                new HyperCubeDimensionDef
                                {
                                    Def = new HyperCubeDimensionqDef
                                    {
                                        CId = ClientExtension.GetCid(),
                                        FieldLabels = new[] {md2.Name},
                                        FieldDefs = new[] {md2.Expression}
                                    }
                                }
                            },
                            Measures = new List<HyperCubeMeasureDef>
                            {
                                new HyperCubeMeasureDef
                                {
                                    Def =
                                        new HyperCubeMeasureqDef
                                        {
                                            Def = mm1.Expression,
                                            Label = mm1.Name,
                                            CId = ClientExtension.GetCid()
                                        }
                                }
                            },
                            //InterColumnSortOrder = new[] { 1, 0 },
                            InitialDataFetch =
                                new List<NxPage>
                                {
                                    new NxPage { Height = 500, Left = 0, Top = 0, Width = 10 }
                                }.ToArray()
                        }
                    };
                    TChart = Sheet.CreateTreemap(Id, properties);
                }
                else
                {
                    var properties = new TreemapProperties()
                    {
                        Title = ChartTitle == "" ? Id : ChartTitle,
                        ShowTitles = ChartTitle == "" ? false : true,
                        Color = new ColorMode
                        {
                            Auto = false,
                            Mode = ColorModeMode.ByDimension,
                            SingleColor = 3,
                            Persistent = false,
                            DimensionScheme = ColorModeDimensionScheme.DimensionScheme12,
                            MeasureScheme = ColorModeMeasureScheme.Sg,
                            ReverseScheme = false
                        },
                        HyperCubeDef = new VisualizationHyperCubeDef
                        {
                            Mode = NxHypercubeMode.DATA_MODE_PIVOT_STACK,
                            Dimensions = new List<HyperCubeDimensionDef>
                            {
                                new HyperCubeDimensionDef
                                {
                                    Def = new HyperCubeDimensionqDef
                                    {
                                        CId = ClientExtension.GetCid(),
                                        FieldLabels = new[] {md1.Name},
                                        FieldDefs = new[] {md1.Expression}
                                    }
                                }
                            },
                            Measures = new List<HyperCubeMeasureDef>
                            {
                                new HyperCubeMeasureDef
                                {
                                    Def =
                                        new HyperCubeMeasureqDef
                                        {
                                            Def = mm1.Expression,
                                            Label = mm1.Name,
                                            CId = ClientExtension.GetCid()
                                        }
                                }
                            },
                            //InterColumnSortOrder = new[] { 1, 0 },
                            InitialDataFetch =
                                new List<NxPage>
                                {
                                    new NxPage { Height = 500, Left = 0, Top = 0, Width = 10 }
                                }.ToArray()
                        }
                    };
                    TChart = Sheet.CreateTreemap(Id, properties);
                }


                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var treechartCell = Sheet.CellFor(TChart);
                    treechartCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                TChart = null;
            }
            //return TChart;
        }

        public void QSCreatePivotTableChart(string SheetID, string Id, string Dimension1, string Dimension2, string Dimension3,
            string Measure1 = null, string Measure2 = null, string Measure3 = null,
            string ChartTitle = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            IPivottable PTChart = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem mm2 = GetMasterMeasure(Measure2);
                QSMasterItem mm3 = GetMasterMeasure(Measure3);
                QSMasterItem md1 = GetMasterDimension(Dimension1);
                QSMasterItem md2 = GetMasterDimension(Dimension2);
                QSMasterItem md3 = GetMasterDimension(Dimension3);

                var properties = new PivottableProperties()
                {
                    Title = ChartTitle == "" ? Id : ChartTitle,
                    ShowTitles = ChartTitle == "" ? false : true,
                    HyperCubeDef = new VisualizationHyperCubeDef
                    {
                        Mode = NxHypercubeMode.DATA_MODE_PIVOT,
                        Dimensions = new List<HyperCubeDimensionDef>
                        {
                            new HyperCubeDimensionDef
                            {
                                Def = new HyperCubeDimensionqDef
                                {
                                    CId = ClientExtension.GetCid(),
                                    FieldLabels = new[] {md1.Name},
                                    FieldDefs = new[] {md1.Expression},
                                    Grouping = NxGrpType.GRP_NX_COLLECTION
                                }
                            },
                            new HyperCubeDimensionDef
                            {
                                Def = new HyperCubeDimensionqDef
                                {
                                    CId = ClientExtension.GetCid(),
                                    FieldLabels = new[] {md2.Name},
                                    FieldDefs = new[] {md2.Expression},
                                    Grouping = NxGrpType.GRP_NX_COLLECTION
                                }
                            },
                            new HyperCubeDimensionDef
                            {
                                Def = new HyperCubeDimensionqDef
                                {
                                    CId = ClientExtension.GetCid(),
                                    FieldLabels = new[] {md3.Name},
                                    FieldDefs = new[] {md3.Expression},
                                    Grouping = NxGrpType.GRP_NX_COLLECTION
                                }
                            }
                        },
                        Measures = new List<HyperCubeMeasureDef>
                        {
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm1.Expression,
                                        Label = mm1.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            },
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm2.Expression,
                                        Label = mm2.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            },
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm3.Expression,
                                        Label = mm3.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            }
                        },
                        //InterColumnSortOrder = new[] { 1, 0 },
                        InitialDataFetch =
                            new List<NxPage>
                            {
                                new NxPage { Height = 500, Left = 0, Top = 0, Width = 10 }
                            }.ToArray()
                    }
                };
                PTChart = Sheet.CreatePivottable(Id, properties);


                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var pivotchartCell = Sheet.CellFor(PTChart);
                    pivotchartCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                PTChart = null;
            }
            //return PTChart;
        }

        public void QSCreateScatterChart(string SheetID, string Id, string Dimension1, string Measure1,
            string Measure2, string Measure3 = null,
            string ChartTitle = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            IScatterplot SChart = null;

            try
            {
                int NumMeasures = 3;

                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem mm2 = GetMasterMeasure(Measure2);
                QSMasterItem mm3 = null;
                if (Measure3 == null)
                    NumMeasures = 2;
                else
                    mm3 = GetMasterMeasure(Measure3);
                IMeasure[] im = new IMeasure[NumMeasures];
                im[0] = qsApp.GetMeasure(mm1.Id);
                im[1] = qsApp.GetMeasure(mm2.Id);
                if (NumMeasures > 2) im[2] = qsApp.GetMeasure(mm3.Id);

                QSMasterItem md1 = GetMasterDimension(Dimension1);

                var properties = new ScatterplotProperties()
                {
                    Title = ChartTitle == "" ? Id : ChartTitle,
                    ShowTitles = ChartTitle == "" ? false : true,
                    HyperCubeDef = new VisualizationHyperCubeDef
                    {
                        Dimensions = new List<HyperCubeDimensionDef>
                        {
                            new HyperCubeDimensionDef
                            {
                                Def = new HyperCubeDimensionqDef
                                {
                                    CId = ClientExtension.GetCid(),
                                    FieldLabels = new[] {md1.Name},
                                    FieldDefs = new[] {md1.Expression}
                                }
                            }
                        },
                        Measures = new List<HyperCubeMeasureDef>
                        {
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm1.Expression,
                                        Label = mm1.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            },
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm2.Expression,
                                        Label = mm2.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            },
                            new HyperCubeMeasureDef
                            {
                                Def =
                                    new HyperCubeMeasureqDef
                                    {
                                        Def = mm3.Expression,
                                        Label = mm3.Name,
                                        CId = ClientExtension.GetCid()
                                    }
                            }
                        },
                        //InterColumnSortOrder = new[] { 1, 0 },
                        InitialDataFetch =
                            new List<NxPage>
                            {
                                new NxPage { Height = 500, Left = 0, Top = 0, Width = 10 }
                            }.ToArray()
                    }
                };
                SChart = Sheet.CreateScatterplot(Id, properties);


                ScatterplotProperties p = SChart.Properties;
                p.Title = Measure1 + " vs " + Measure2;

                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var scatterchartCell = Sheet.CellFor(SChart);
                    scatterchartCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                SChart = null;
            }
            //return PTChart;
        }

        public void QSCreateFilterPane(string SheetID, string Id, string Dimension1, string Dimension2 = null, string Dimension3 = null,
            string ChartTitle = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            IFilterpane FPane = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem md1 = GetMasterDimension(Dimension1);
                QSMasterItem md2 = GetMasterDimension(Dimension2);
                QSMasterItem md3 = GetMasterDimension(Dimension3);

                var props = new FilterpaneProperties
                {
                    Title = ChartTitle == "" ? Id : ChartTitle,
                    ShowTitles = ChartTitle == "" ? false : true,
                    ChildListDef = new FilterpaneListboxObjectViewListDef { Data = new FilterpaneListboxObjectViewDef() },
                    Visualization = "filterpane"
                };

                var filterpane = Sheet.CreateFilterpane(Id, props);
                var dim1Props = new ListboxProperties
                {
                    Title = md1.Name,
                    ListObjectDef = new ListboxListObjectDef
                    {
                        InitialDataFetch = new[] { Pager.Default },
                        Def =
                            new ListboxListObjectDimensionDef
                            {
                                FieldDefs = new List<string> { md1.Expression },
                                FieldLabels = new List<string> { md1.Name },
                                CId = ClientExtension.GetCid()
                            }
                    }
                };
                var dim2Props = new ListboxProperties
                {
                    Title = md2.Name,
                    ListObjectDef = new ListboxListObjectDef
                    {
                        InitialDataFetch = new[] { Pager.Default },
                        Def =
                            new ListboxListObjectDimensionDef
                            {
                                FieldDefs = new List<string> { md2.Expression },
                                FieldLabels = new List<string> { md2.Name },
                                CId = ClientExtension.GetCid()
                            }
                    }
                };
                var dim3Props = new ListboxProperties
                {
                    Title = md3.Name,
                    ListObjectDef = new ListboxListObjectDef
                    {
                        InitialDataFetch = new[] { Pager.Default },
                        Def =
                            new ListboxListObjectDimensionDef
                            {
                                FieldDefs = new List<string> { md3.Expression },
                                FieldLabels = new List<string> { md3.Name },
                                CId = ClientExtension.GetCid()
                            }
                    }
                };
                filterpane.CreateListbox(md1.Id, dim1Props);
                filterpane.CreateListbox(md2.Id, dim2Props);
                filterpane.CreateListbox(md3.Id, dim3Props);

                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var fpaneCell = Sheet.CellFor(FPane);
                    fpaneCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                FPane = null;
            }
            //return FPane;
        }

        public void QSCreateMapPoint(string SheetID, string Id, string Dimension1, string Measure1,
            string ChartTitle = "",
            int col = 0, int row = 0, int width = 0, int height = 0)
        {
            IMap Map = null;

            try
            {
                ISheet Sheet = qsApp.GetSheet(SheetID);
                QSMasterItem mm1 = GetMasterMeasure(Measure1);
                QSMasterItem md1 = GetMasterDimension(Dimension1);

                MapLayerDataContainer mapLayerData = new MapLayerDataContainer(Qlik.Sense.Client.Visualizations.MapComponents.LayerType.Point)
                {
                    Dimension = qsApp.GetDimension(md1.Id),
                    Measure = qsApp.GetMeasure(mm1.Id)
                };

                Map = Sheet.CreateMap(Id, new List<MapLayerDataContainer> { mapLayerData });

                if (col > 0 && row > 0 && width > 0 && height > 0)
                {
                    var mappointCell = Sheet.CellFor(Map);
                    mappointCell.SetBounds(row, col, width, height);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Map = null;
            }
            //return Map;
        }


        public string QSCreateApp(string FileName, string AppName = "Telegram", string FolderConnection = "TelegramFiles")
        {
            IAppIdentifier myAppId;

            try
            {
                try
                {
                    myAppId = qsLocation.AppWithNameOrDefault(AppName);
                    if (myAppId != null)
                        qsLocation.Delete(myAppId);
                }
                catch (Exception e) { }

                myAppId = qsLocation.CreateAppWithName(AppName);
                IApp myApp = qsLocation.App(myAppId);
                string localizationScript = myApp.GetEmptyScript();
                string newScript = "";

                if (Path.GetExtension(FileName).ToLower() == ".csv")
                {
                    newScript = localizationScript + "\n\nLOAD  * FROM [lib://" + FolderConnection + "/" + FileName
                        + "] (txt, utf8, embedded labels, delimiter is ';', msq);";
                }
                else if (Path.GetExtension(FileName) == ".xls" || Path.GetExtension(FileName) == ".xlsx")
                {
                    return null;
                }

                myApp.SetScript(newScript);
                string currentScript = myApp.GetScript();
                bool didReload = myApp.DoReload();
                if (!didReload) Console.WriteLine("The app " + AppName + " did not reload.");

                myApp.DoSave();

            }
            catch (Exception e)
            {
                if (e.HResult == -2146233074)
                    Console.WriteLine("You cannot create apps :-(");
                else
                {
                    Console.WriteLine(e);
                }
                return null;
            }

            return myAppId.AppId;
        }

        public bool QSPublishApp(string StreamId, string NewName = null)
        {
            try
            {
                qsApp.Publish(StreamId, NewName);
                return true;
            }
            catch (Exception e)
            {   // Error
                Console.WriteLine(e);
                return false;
            }
        }

        public void QSCreateMasterItemsFromFields()
        {
            if (MasterDimensions.Count + MasterMeasures.Count + MasterVisualizations.Count > 0)
                return;

            List<string> Aggregations = new List<string> { "Sum", "Sum", "Sum", "Count", "Count", "Count", "Count", "Count", "Count", "Avg", "Max", "Min" };

            foreach (QSField field in AppFields)
            {
                List<string> tags = field.Tags.ToList();

                if (tags.Contains("$hidden") || tags.Contains("$system") || tags.Contains("$key")
                    || field.FieldToSearch.Contains("fax") || field.FieldToSearch.Contains("phone")
                    || field.FieldToSearch.StartsWith("id") || field.FieldToSearch.EndsWith("id")
                    || field.FieldToSearch.StartsWith("code") || field.FieldToSearch.EndsWith("code")
                    || field.FieldToSearch.StartsWith("_")
                    )
                    continue;
                else if (tags.Contains("$text") || tags.Contains("$ascii")
                    || field.FieldToSearch.Contains("name") || field.FieldToSearch.Contains("desc") || field.FieldToSearch.Contains("address")
                    || field.FieldToSearch.Contains("title") || field.FieldToSearch.Contains("city") || field.FieldToSearch.Contains("customer")
                    || field.FieldToSearch.Contains("nombre") || field.FieldToSearch.Contains("direcc") || field.FieldToSearch.Contains("titulo")
                    || field.FieldToSearch.Contains("título") || field.FieldToSearch.Contains("ciudad") || field.FieldToSearch.Contains("cliente"))
                {
                    string dim = field.FieldNameInApp;
                    IDimension myDimension = qsApp.CreateDimension(dim,
                        new DimensionProperties
                        {
                            MetaDef = new MetaAttributesDef
                            {
                                Title = dim,
                                Description = "Dimension " + dim + " created by the Qlik Sense Bot"
                            },
                            Dim = new NxLibraryDimensionDef
                            {
                                FieldDefs = new[] { dim },
                                FieldLabels = new[] { dim }
                            }
                        });
                }
                else if (tags.Contains("$date") || tags.Contains("$timestamp")
                    || field.FieldToSearch.Contains("date") || field.FieldToSearch.Contains("time") || field.FieldToSearch.Contains("year")
                    || field.FieldToSearch.Contains("month") || field.FieldToSearch.Contains("week") || field.FieldToSearch.Contains("quarter")
                    || field.FieldToSearch.Contains("fecha") || field.FieldToSearch.Contains("hora") || field.FieldToSearch.Contains("año")
                    || field.FieldToSearch.Contains("mes") || field.FieldToSearch.Contains("semana") || field.FieldToSearch.Contains("trimestre")
                    )
                {
                    string dim = field.FieldNameInApp;
                    IDimension myDimension = qsApp.CreateDimension(dim,
                        new DimensionProperties
                        {
                            MetaDef = new MetaAttributesDef
                            {
                                Title = dim,
                                Description = "Dimension " + dim + " created by the Qlik Sense Bot"
                            },
                            Dim = new NxLibraryDimensionDef
                            {
                                FieldDefs = new[] { dim },
                                FieldLabels = new[] { dim }
                            }
                        });
                }
                else if (tags.Contains("$geoname") || tags.Contains("$geopoint") || tags.Contains("$geomultipolygon")
                    || field.FieldToSearch.Contains("latitude") || field.FieldToSearch.Contains("longitude"))
                {
                    string dim = field.FieldNameInApp;
                    IDimension myDimension = qsApp.CreateDimension(dim,
                        new DimensionProperties
                        {
                            MetaDef = new MetaAttributesDef
                            {
                                Title = dim,
                                Description = "Dimension " + dim + " created by the Qlik Sense Bot"
                            },
                            Dim = new NxLibraryDimensionDef
                            {
                                FieldDefs = new[] { dim },
                                FieldLabels = new[] { dim }
                            }
                        });
                }
                else if (tags.Contains("$numeric") || tags.Contains("$integer")
                    || field.FieldToSearch.Contains("#") || field.FieldToSearch.Contains("$") || field.FieldToSearch.Contains("€")
                    || field.FieldToSearch.Contains("sales") || field.FieldToSearch.Contains("cos") || field.FieldToSearch.Contains("discount")
                    || field.FieldToSearch.Contains("quantity") || field.FieldToSearch.Contains("units") || field.FieldToSearch.Contains("salary")
                    || field.FieldToSearch.Contains("ventas") || field.FieldToSearch.Contains("descuento")
                    || field.FieldToSearch.Contains("cantidad") || field.FieldToSearch.Contains("unidad") || field.FieldToSearch.Contains("salario")
                    )
                {
                    string meas = field.FieldNameInApp;
                    IMeasure myMeasure = qsApp.CreateMeasure(meas,
                        new MeasureProperties
                        {
                            Measure = new NxLibraryMeasureDef
                            {
                                Def = "Sum([" + meas + "])",
                                Label = meas
                            },
                            MetaDef = new MetaAttributesDef
                            {
                                Title = meas,
                                Description = "Measure " + meas + " created by the Qlik Sense Bot"
                            }
                        });
                }
                else if (!field.FieldToSearch.Contains("id") && !field.FieldToSearch.Contains("cod"))
                {
                    string meas = field.FieldNameInApp;
                    string aggr = Aggregations.PickRandom();

                    IMeasure myMeasure = qsApp.CreateMeasure(meas,
                        new MeasureProperties
                        {
                            Measure = new NxLibraryMeasureDef
                            {
                                Def = aggr + "([" + meas + "])",
                                Label = aggr + " of " + meas
                            },
                            MetaDef = new MetaAttributesDef
                            {
                                Title = aggr + " of " + meas,
                                Description = "Measure " + meas + " created by the Qlik Sense Bot"
                            }
                        });
                }
            }
            qsApp.DoSave();
            QSReadMasterItems();
        }


        private void QSReadStories()
        {
            Stories.Clear();
            foreach (Qlik.Sense.Client.Storytelling.IStory AppStory in Qlik.Sense.Client.AppExtensions.GetStories(qsApp))
            {
                QSStory qst = new QSStory();
                qst.Id = AppStory.Id;
                qst.Name = AppStory.Properties.MetaDef.Title;
                Stories.Add(qst);
            }
        }


        private void QSReadFields()
        {
            AppFields.Clear();
            Qlik.Sense.Client.IFieldList fieldList = qsApp.GetFieldList();
            LatitudeField = null;
            LongitudeField = null;
            AddressField = null;
            GeoLocationActive = false;

            foreach (NxFieldDescription field in fieldList.Items)
            {
                QSField f = new QSField();
                f.FieldNameInApp = field.Name;
                f.FieldToSearch = field.Name.ToLower().Trim();
                f.Tags = field.Tags;
                AppFields.Add(f);

                if (f.FieldToSearch == "latitude") LatitudeField = f;
                if (f.FieldToSearch.Contains("latitude") && LatitudeField == null) LatitudeField = f;

                if (f.FieldToSearch == "longitude") LongitudeField = f;
                if (f.FieldToSearch.Contains("longitude") && LongitudeField == null) LongitudeField = f;

                if (f.FieldToSearch == "address") AddressField = f;
                if (f.FieldToSearch.Contains("address") && AddressField == null) AddressField = f;
            }

            if (LatitudeField != null && LongitudeField != null) GeoLocationActive = true;
        }

        private void QSReadMasterItems()
        {
            MasterVisualizations.Clear();
            try
            {
                var allMasterObjects = GetAllMasterObjects(qsApp);

                foreach (IMasterObject mo in allMasterObjects)
                {
                    QSMasterItem mi = new QSMasterItem();
                    var properties = mo.Properties;
                    mi.Id = properties.Info.Id;
                    mi.Name = properties.MetaDef.Title;
                    mi.Tags = mo.MetaAttributes.Tags;
                    MasterVisualizations.Add(mi);
                }
                MasterVisualizations = (List<QSMasterItem>)MasterVisualizations.Shuffle();
            }
            catch (Exception e)
            {
                Console.WriteLine("QSEasy Error in QSReadMasterItems: {0} Exception caught.", e);
            }

            MasterMeasures.Clear();
            try
            {
                var allMeasures = qsApp.GetMeasureList().Items;

                foreach (IMeasureObjectViewListContainer mm in allMeasures)
                {
                    QSMasterItem mi = new QSMasterItem();
                    INxLibraryMeasureDef md = qsApp.GetMeasure(mm.Info.Id).NxLibraryMeasureDef;
                    mi.Id = mm.Info.Id;
                    mi.Name = md.Label;
                    mi.Expression = md.Def;
                    mi.FormattedExpression = GetExpressionFormattedValue(Measure: mi.Expression, Label: mi.Name);
                    MasterMeasures.Add(mi);
                }
                MasterMeasures = (List<QSMasterItem>)MasterMeasures.Shuffle();
                LastMeasure = MasterMeasures.Count > 0 ? MasterMeasures.First() : null;
            }
            catch (Exception e)
            {
                Console.WriteLine("QSEasy Error in QSReadMasterItems: {0} Exception caught.", e);
            }

            MasterDimensions.Clear();
            try
            {
                var allDimensions = qsApp.GetDimensionList().Items;

                foreach (DimensionObjectViewListContainer md in allDimensions)
                {
                    if (md.Data.Grouping == NxGrpType.GRP_NX_NONE)
                    {
                        QSMasterItem mi = new QSMasterItem();

                        INxLibraryDimensionDef dd = qsApp.GetDimension(md.Info.Id).NxLibraryDimensionDef;
                        mi.Id = md.Info.Id;
                        mi.Name = md.Data.Title;
                        mi.Expression = dd.FieldDefs.First();
                        MasterDimensions.Add(mi);
                    }
                }
                MasterDimensions = (List<QSMasterItem>)MasterDimensions.Shuffle();
                LastDimension = MasterDimensions.Count > 0 ? MasterDimensions.First() : null;
            }
            catch (Exception e)
            {
                Console.WriteLine("QSEasy Error in QSReadMasterItems: {0} Exception caught.", e);
            }

        }

        private static IEnumerable<IMasterObject> GetAllMasterObjects(IApp app)
        {
            return app.GetMasterObjectList().Items?.Select(item => app.GetObject<MasterObject>(item.Info.Id));
        }


        public string QSFindField(string FieldName)
        {
            QSField field = AppFields.Find(f => f.FieldToSearch == FieldName.ToLower().Trim());
            if (field != null)
            {
                return field.FieldNameInApp;
            }
            else
            {
                return null;
            }
        }

        public QSMasterItem GetMasterMeasure(string MeasureName)
        {
            QSMasterItem meas = MasterMeasures.Find(m => m.Name.ToLower().Trim() == MeasureName.ToLower().Trim());

            if (meas == null)
            {
                meas = MasterMeasures.Find(m => m.Name.ToLower().Trim().Contains(MeasureName.ToLower().Trim()));
            }
            if (meas == null)
            {
                var result = MasterMeasures.LevenshteinDistanceOf(m => m.Name)
                    .ComparedTo(MeasureName)
                    .OrderBy(m => m.Distance);
                if (result.Count() > 0)
                    meas = (QSMasterItem)result.First().Item;
            }

            if (meas != null)
            {
                LastMeasure = meas;
                return meas;
            }
            else
            {
                return null;
            }
        }

        public string GetDimensionExpression(string DimensionName)
        {
            string DimensionExpression;

            QSMasterItem dim = MasterDimensions.Find(d => d.Name.ToLower().Trim() == DimensionName.ToLower().Trim());

            if (dim != null)
            {
                DimensionExpression = dim.Expression;
                return DimensionExpression;
            }
            else
            {
                DimensionExpression = QSFindField(DimensionName);

                if (DimensionExpression != null)
                {
                    return DimensionExpression;
                }
                else
                {
                    dim = MasterDimensions.Find(d => d.Name.ToLower().Trim().Contains(DimensionName.ToLower().Trim()));
                    if (dim != null)
                    {
                        DimensionExpression = dim.Expression;
                        return DimensionExpression;
                    }
                    else
                    {
                        var result = MasterDimensions.LevenshteinDistanceOf(d => d.Name)
                            .ComparedTo(DimensionName)
                            .OrderBy(d => d.Distance);
                        if (result.Count() > 0)
                            dim = (QSMasterItem)result.First().Item;
                        if (dim != null)
                        {
                            DimensionExpression = dim.Expression;
                            return DimensionExpression;
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }


        public QSMasterItem GetMasterDimension(string DimensionName)
        {
            QSMasterItem dim = MasterDimensions.Find(d => d.Name.ToLower().Trim() == DimensionName.ToLower().Trim());

            if (dim == null)
            {
                dim = MasterDimensions.Find(d => d.Name.ToLower().Trim().Contains(DimensionName.ToLower().Trim()));
            }

            if (dim == null)
            {
                var result = MasterDimensions.LevenshteinDistanceOf(d => d.Name)
                    .ComparedTo(DimensionName)
                    .OrderBy(d => d.Distance);
                if (result.Count() > 0)
                    dim = (QSMasterItem)result.First().Item;
            }

            LastDimension = dim;
            return dim;
        }


        public QSFoundFilter[] QSSearch(string MyText)
        {
            QSFoundFilter[] Founds;

            SearchCombinationOptions MySearchOptions = new SearchCombinationOptions();
            MySearchOptions.Context = SearchContextType.CONTEXT_CLEARED;

            SearchPage MySearchPage = new SearchPage();
            MySearchPage.Count = 5;

            string[] MySearchTerms = { MyText };


            SearchResult MyResult = qsApp.SearchResults(MySearchOptions, MySearchTerms, MySearchPage);

            Founds = new QSFoundFilter[5];
            var i = 0;

            foreach (var group in MyResult.SearchGroupArray)
            {
                foreach (var item in group.Items)
                {
                    QSFoundFilter f = new QSFoundFilter();

                    f.Dimension = item.Identifier;
                    f.Element = item.ItemMatches.First().Text;
                    Founds[i] = f;
                    i++;
                }
            }

            return Founds;

        }


        public QSFoundFilter[] QSSearchInDimension(string MyText, MasterObject Dim)
        {
            QSFoundFilter[] Founds;

            SearchCombinationOptions MySearchOptions = new SearchCombinationOptions();
            MySearchOptions.Context = SearchContextType.CONTEXT_CLEARED;

            SearchPage MySearchPage = new SearchPage();
            MySearchPage.Count = 5;

            string[] MySearchTerms = { MyText };


            SearchResult MyResult = qsApp.SearchResults(MySearchOptions, MySearchTerms, MySearchPage);

            Founds = new QSFoundFilter[5];
            var i = 0;

            foreach (var group in MyResult.SearchGroupArray)
            {
                foreach (var item in group.Items)
                {
                    QSFoundFilter f = new QSFoundFilter();

                    f.Dimension = item.Identifier;
                    f.Element = item.ItemMatches.First().Text;
                    Founds[i] = f;
                    i++;
                }
            }

            return Founds;

        }



        public QSFoundObject[] QSSearchObjects(string MyText, bool ShowSelectionsBar = false)
        {
            List<string> Founds = new List<string>();
            List<string> Ids = new List<string>();
            List<string> OTypes = new List<string>();

            SearchCombinationOptions MySearchOptions = new SearchCombinationOptions();
            MySearchOptions.Context = SearchContextType.CONTEXT_CLEARED;

            SearchPage MySearchPage = new SearchPage();
            MySearchPage.Count = 1;
            MySearchPage.MaxNbrFieldMatches = maxFounds;

            SearchGroupOptions[] qGroupOptions = new SearchGroupOptions[1];
            SearchGroupOptions sgo = new SearchGroupOptions();
            sgo.Count = maxFounds;
            sgo.GroupType = SearchGroupType.GENERIC_OBJECTS_GROUP;
            qGroupOptions[0] = sgo;

            SearchGroupItemOptions[] qGroupItemOptions = new SearchGroupItemOptions[1];
            SearchGroupItemOptions sgio = new SearchGroupItemOptions();
            sgio.Count = 1;
            sgio.GroupItemType = SearchGroupItemType.GENERIC_OBJECT;
            qGroupItemOptions[0] = sgio;

            MySearchPage.GroupOptions = qGroupOptions;
            MySearchPage.GroupItemOptions = qGroupItemOptions;

            char[] sep = { ' ' };
            string[] MySearchTerms = MyText.Split(sep);

            SearchObjectOptions MySearchObjectOptions = new SearchObjectOptions();


            SearchResult MyObjects = qsApp.SearchObjects(MySearchObjectOptions, MySearchTerms, MySearchPage);

            foreach (var obGroup in MyObjects.SearchGroupArray)
            {
                foreach (var obItem in obGroup.Items)
                {
                    if (Founds.Count > maxFounds) break;

                    GenericObject MyObject = qsApp.GetGenericObject(obItem.Identifier);

                    Qlik.Sense.Client.Visualizations.VisualizationBaseProperties MyProp = MyObject.Properties as Qlik.Sense.Client.Visualizations.VisualizationBaseProperties;

                    string t;
                    t = MyProp.Title;
                    if (t.IndexOf("'") > -1 && t.IndexOf("&") > -1) t = qsApp.Evaluate(t);
                    Founds.Add(t);
                    Ids.Add(obItem.Identifier);
                    OTypes.Add(MyProp.Visualization);
                }
            }

            QSFoundObject[] qsFounds;

            if (Founds.Count < maxFounds)
            {
                qsFounds = new QSFoundObject[Founds.Count];
            }
            else
            {
                qsFounds = new QSFoundObject[maxFounds];
            }


            for (int i = 0; i < Founds.Count && i < maxFounds; i++)
            {
                if (Founds[i] == "") Founds[i] = "-----";
                qsFounds[i] = new QSFoundObject();
                qsFounds[i].Description = Founds[i];
                qsFounds[i].ObjectType = OTypes[i];
                qsFounds[i].ObjectId = Ids[i];
                qsFounds[i].ObjectURL = qsSingleServer + "/single?appid=" + qsSingleApp + "&obj=" + Ids[i]; //+ "&select=clearall";
                if (ShowSelectionsBar) qsFounds[i].ObjectURL += "&opt=currsel";
                qsFounds[i].HRef = "<a href=\"" + qsFounds[i].ObjectURL + "\">" + Founds[i] + "</a>";
                qsFounds[i].ThumbURL = qsSingleServer + "/resources/qsimg/" + OTypes[i] + ".png";
            }
            return qsFounds;
        }

        public string GetExpression(string Expression)
        {
            string exp;

            try
            {
                exp = qsApp.Evaluate(Expression);
            }
            catch (Exception e)
            {
                Console.WriteLine("QSEasy Error: {0} Exception caught.", e);
                exp = "";
            }

            if (Expression == "" || exp.StartsWith("Error:"))
            {
                exp = "";
            }

            return (exp);
        }


        public string ApplyGeoFilter(string LatField, string LatFilter, string LonField, string LonFilter, string IDField)
        {
            string Selection = "";

            qsApp.ClearAll();
            qsApp.GetField(LatField).Select(LatFilter);
            qsApp.GetField(LonField).Select(LonFilter);

            IAppField fID = qsApp.GetAppField(IDField);

            var p = new List<NxPage> { new NxPage { Height = 30, Width = 1 } };
            //var dataPages = fID.GetData(p);
            var dataPages = fID.GetOptional(p);

            foreach (var dataPage in dataPages)
            {
                var matrix = dataPage.Matrix;
                foreach (var cellRows in matrix)
                {
                    foreach (var cellRow in cellRows)
                    {
                        //Console.WriteLine("## " + cellRow.Text + " - " + cellRow.State);
                        Selection = Selection + ',' + cellRow.Text.Replace(',', '.');
                    }
                }
            }

            if (Selection != "")
            {
                Selection = "&select=" + IDField + Selection;
            }
            return Selection;
        }


        public QSGeoFilter GetGeoFilters(double Latitude, double Longitude, double Distance)
        {
            QSGeoFilter Filter = new QSGeoFilter();

            double MIN_LAT = -90 * (Math.PI / 180);
            double MAX_LAT = 90 * (Math.PI / 180);
            double MIN_LON = -180 * (Math.PI / 180);
            double MAX_LON = 180 * (Math.PI / 180);
            double R = 6378.1;

            double radDist = Distance / R;
            double degLat = Latitude;
            double degLon = Longitude;
            double radLat = degLat * (Math.PI / 180);
            double radLon = degLon * (Math.PI / 180);
            double minLat = radLat - radDist;
            double maxLat = radLat + radDist;
            double minLon = 0;
            double maxLon = 0;
            double deltaLon = Math.Asin(Math.Sin(radDist) / Math.Cos(radLat));
            if (minLat > MIN_LAT && maxLat < MAX_LAT)
            {
                minLon = radLon - deltaLon;
                maxLon = radLon + deltaLon;
                if (minLon < MIN_LON)
                {
                    minLon = minLon + 2 * Math.PI;
                }
                if (maxLon > MAX_LON)
                {
                    maxLon = maxLon - 2 * Math.PI;
                }
            }
            else
            {
                minLat = Math.Max(minLat, MIN_LAT);
                maxLat = Math.Min(maxLat, MAX_LAT);
                minLon = MIN_LON;
                maxLon = MAX_LON;
            }

            minLat = (180 * minLat) / Math.PI;
            maxLat = (180 * maxLat) / Math.PI;
            minLon = (180 * minLon) / Math.PI;
            maxLon = (180 * maxLon) / Math.PI;

            Filter.FromLatitude = Latitude;
            Filter.FromLongitude = Longitude;
            Filter.LatFilter = ">=" + minLat + "<=" + maxLat;
            Filter.LonFilter = ">=" + minLon + "<=" + maxLon;

            Filter.LatFilter = Filter.LatFilter.Replace(",", ".");
            Filter.LonFilter = Filter.LonFilter.Replace(",", ".");
            return Filter;
        }

        public List<QSGeoList> GetGeoList(QSGeoFilter GeoFilter, QSMasterItem Dimension, QSMasterItem Measure,
            List<QSFoundFilter> Filters = null, int NoOfRows = 5)
        {
            if (!GeoLocationActive) return null;

            List<QSGeoList> gl = new List<QSGeoList>();

            qsApp.ClearAll();
            if (Filters != null && Filters.Count > 0)
            {
                foreach (QSFoundFilter ff in Filters)
                {
                    try
                    {
                        string DimExpression = GetDimensionExpression(ff.Dimension);
                        if (DimExpression != null)
                            qsApp.GetField(DimExpression).Select(ff.Element.ToUpper() + "*");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("QSEasy Error in GetGeoList with filter \"{0} ={1}\": {2} Exception caught.", ff.Dimension, ff.Element, e);
                    }
                }
            }

            qsApp.GetField(LatitudeField.FieldNameInApp).Select(GeoFilter.LatFilter);
            qsApp.GetField(LongitudeField.FieldNameInApp).Select(GeoFilter.LonFilter);

            string DistanceExpression = "sqrt(sqr([" + LatitudeField.FieldNameInApp + "] - (" + GeoFilter.FromLatitude + ")) + sqr([" +
                LongitudeField.FieldNameInApp + "] - (" + GeoFilter.FromLongitude + ")))";

            var myDimension1 = new NxDimension
            {
                Def = new NxInlineDimensionDef()
            };
            myDimension1.Def.FieldDefs = new[] { LatitudeField.FieldNameInApp };

            SortCriteria dsc = new SortCriteria();
            dsc.SortByExpression = SortDirection.Ascending;
            dsc.Expression = DistanceExpression;
            myDimension1.Def.SortCriterias = new[] { dsc };

            var myDimension2 = new NxDimension
            {
                Def = new NxInlineDimensionDef()
            };
            myDimension2.Def.FieldDefs = new[] { LongitudeField.FieldNameInApp };

            var myDimension3 = new NxDimension
            {
                Def = new NxInlineDimensionDef()
            };
            myDimension3.Def.FieldDefs = new[] { Dimension.Expression };

            var myDimension4 = new NxDimension
            {
                Def = new NxInlineDimensionDef()
            };
            if (AddressField == null)
                myDimension4.Def.FieldDefs = new[] { Dimension.Expression };
            else
                myDimension4.Def.FieldDefs = new[] { AddressField.FieldNameInApp };


            var myMeasure = new NxMeasure
            {
                Def = new NxInlineMeasureDef()
            };
            myMeasure.Def.Def = Measure.Expression;

            var Distance = new NxMeasure
            {
                Def = new NxInlineMeasureDef()
            };
            myMeasure.Def.Def = DistanceExpression;

            HyperCubeDef MyCubeDef = new HyperCubeDef();
            MyCubeDef.Dimensions = new List<NxDimension>
            {
                new NxDimension {Def = myDimension1.Def },
                new NxDimension {Def = myDimension2.Def },
                new NxDimension {Def = myDimension3.Def },
                new NxDimension {Def = myDimension4.Def }
            };
            MyCubeDef.Measures = new List<NxMeasure>
            {
                new NxMeasure { Def = myMeasure.Def },
                new NxMeasure { Def = Distance.Def }
            };

            MyCubeDef.InitialDataFetch = new List<NxPage>
            {
                new NxPage {Height = NoOfRows, Width = 6}
            };

            GenericObjectProperties gp = new GenericObjectProperties();
            gp.Info = new NxInfo();
            gp.Info.Type = "hypercube";
            gp.Set<HyperCubeDef>("qHyperCubeDef", MyCubeDef);
            GenericObject obj = qsApp.CreateGenericSessionObject(gp);

            var p = new List<NxPage> { new NxPage { Height = NoOfRows, Width = 5 } };
            var dataPages = obj.GetHyperCubeData("/qHyperCubeDef", p);

            foreach (var dataPage in dataPages)
            {
                var matrix = dataPage.Matrix;
                foreach (var cellRows in matrix)
                {
                    if (cellRows[0].IsNull || cellRows[1].IsNull) continue;

                    gl.Add(new QSGeoList
                    {
                        Lat = cellRows[0].Num,
                        Lon = cellRows[1].Num,
                        TextLabel = Dimension.Name,
                        Text = cellRows[2].Text,
                        Address = cellRows[3].Text,
                        ValueLabel = Measure.Name,
                        Value = cellRows[4].Num,
                        FormattedValue = FormatValue(cellRows[4].Num, Measure.Name)
                    });
                }
            }

            return gl;
        }


        public QSDataList[] GetDataList(string Measure, string Dimension, List<QSFoundFilter> Filters = null, int NoOfRows = 0,
            bool Descending = true, double? MeasureThreshold = null)
        {
            List<QSDataList> gl = new List<QSDataList>();

            if (NoOfRows == 0) NoOfRows = maxFounds;

            qsApp.ClearAll();
            if (Filters != null && Filters.Count > 0)
            {
                foreach (QSFoundFilter ff in Filters)
                {
                    try
                    {
                        string DimExpression = GetDimensionExpression(ff.Dimension);
                        if (DimExpression != null)
                            qsApp.GetField(DimExpression).Select(ff.Element.ToUpper() + "*");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("QSEasy Error in GetDataList with filter \"{0} ={1}\": {2} Exception caught.", ff.Dimension, ff.Element, e);
                    }
                }
            }

            var myDimension1 = new NxDimension
            {
                Def = new NxInlineDimensionDef()
            };
            myDimension1.Def.FieldDefs = new[] { Dimension };

            SortCriteria dsc = new SortCriteria();
            if (Descending)
                dsc.SortByExpression = SortDirection.Descending;
            else
                dsc.SortByExpression = SortDirection.Ascending;

            dsc.Expression = Measure;
            myDimension1.Def.SortCriterias = new[] { dsc };

            var myMeasure = new NxMeasure
            {
                Def = new NxInlineMeasureDef()
            };
            if (Measure.First() != '=') Measure = "=" + Measure;
            myMeasure.Def.Def = Measure;

            if (MeasureThreshold != null)
            {
                if (Measure.First() == '=') Measure = Measure.Remove(0, 1);
                Measure = string.Format("=if({0} {2} {1}, {0}, Null())", Measure, MeasureThreshold.ToString(),
                    Descending ? ">" : "<");
            }
            myMeasure.Def.Def = Measure;

            HyperCubeDef MyCubeDef = new HyperCubeDef();
            MyCubeDef.Dimensions = new List<NxDimension>
            {
                new NxDimension {Def = myDimension1.Def },
            };
            MyCubeDef.Measures = new List<NxMeasure>
            {
                new NxMeasure { Def = myMeasure.Def }
            };

            MyCubeDef.InitialDataFetch = new List<NxPage>
            {
                new NxPage {Height = NoOfRows, Width = 5}
            };
            MyCubeDef.SuppressMissing = true;
            MyCubeDef.SuppressZero = true;

            GenericObjectProperties gp = new GenericObjectProperties();
            gp.Info = new NxInfo();
            gp.Info.Type = "hypercube";
            gp.Set<HyperCubeDef>("qHyperCubeDef", MyCubeDef);
            GenericObject obj = qsApp.CreateGenericSessionObject(gp);

            var p = new List<NxPage> { new NxPage { Height = NoOfRows, Width = 5 } };
            var dataPages = obj.GetHyperCubeData("/qHyperCubeDef", p);

            foreach (var dataPage in dataPages)
            {
                var matrix = dataPage.Matrix;
                foreach (var cellRows in matrix)
                {
                    if (cellRows.Count > 1)
                    {
                        gl.Add(new QSDataList
                        {
                            DimValue = cellRows[0].Text,
                            MeasValue = cellRows[1].Num,
                            MeasFormattedValue = FormatValue(cellRows[1].Num, Measure)
                        });
                    }
                }
            }

            return gl.ToArray();
        }

        public string GetExpressionFormattedValue(string Measure, List<QSFoundFilter> Filters = null, string Label = "")
        {
            return FormatValue(GetExpressionValue(Measure, Filters), Label);
        }

        public double GetExpressionValue(string Measure, List<QSFoundFilter> Filters = null)
        {
            double val = 0;

            qsApp.ClearAll();
            if (Filters != null && Filters.Count > 0)
            {
                foreach (QSFoundFilter ff in Filters)
                {
                    try
                    {
                        string DimExpression = GetDimensionExpression(ff.Dimension);
                        if (DimExpression != null)
                            qsApp.GetField(DimExpression).Select(ff.Element.ToUpper() + "*");

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("QSEasy Error in GetExpressionValue with filter \"{0} ={1}\": {2} Exception caught.", ff.Dimension, ff.Element, e);
                    }
                }
            }

            var myMeasure = new NxMeasure
            {
                Def = new NxInlineMeasureDef()
            };
            if (Measure.First() != '=') Measure = "=" + Measure;
            myMeasure.Def.Def = Measure;

            HyperCubeDef MyCubeDef = new HyperCubeDef();
            MyCubeDef.Measures = new List<NxMeasure>
            {
                new NxMeasure { Def = myMeasure.Def }
            };

            MyCubeDef.InitialDataFetch = new List<NxPage>
            {
                new NxPage {Height = 1, Width = 1}
            };

            GenericObjectProperties gp = new GenericObjectProperties();
            gp.Info = new NxInfo();
            gp.Info.Type = "hypercube";
            gp.Set<HyperCubeDef>("qHyperCubeDef", MyCubeDef);
            GenericObject obj = qsApp.CreateGenericSessionObject(gp);

            var p = new List<NxPage> { new NxPage { Height = 1, Width = 1 } };
            var dataPages = obj.GetHyperCubeData("/qHyperCubeDef", p);
            if (dataPages.Count() > 0)
            {
                var cr = dataPages.First().Matrix.First();
                val = cr[0].Num;
            }

            return val;
        }

        private string FormatValue(double Value, string Label = "")
        {
            string strValue = null;

            if (Value != 0)
            {
                if (Label.Contains("%")) strValue = Value.ToString("P1");
                else if (Value > 0 && Value < 1) strValue = Value.ToString("P1");
                else if (Label.Contains("€") || Label.Contains("$")) strValue = Value.ToString("C2");
                else strValue = Value.ToString("N2");
            }
            else
            {
                strValue = "0";
            }
            return strValue;
        }


        public void QSCheckAlerts(string ReportName = "ReportSales.pdf")
        {
            foreach (var a in AlertList)
            {
                a.AlertMessage = a.AlertRequest;
                int i = rnd.Next(1, 9);
                a.AlertPhoto = string.Format("chart{0}.jpg", i.ToString());
                a.AlertDoc = ReportName;
                a.AlertActive = true;
            }
        }


    }
}

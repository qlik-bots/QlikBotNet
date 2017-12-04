using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Threading;
using System.Configuration;
using QlikConversationService.Multilanguage;

using QlikNLP;
using QlikSenseEasy;


namespace QlikConversationService
{
    public enum ResponseAction { None = 0, OpenApp = 1, ShowKPI = 2, ShowDimension = 3, ShowMeasure = 4, ShowSheet = 5, ShowStory = 6 };

    public class ResponseOptions
    {
        public string Title;
        public string ID;
        public ResponseAction Action;
    }
    public class Response
    {
        public string TextMessage = "";
        public string VoiceMessage = "";
        public string NewsSearch = "";
        public NLP NLPrediction;
        public string WarningText = "";
        public string ErrorText = "";
        public string OtherAction = "";
        public QSFoundObject ChartFound;
        public List<ResponseOptions> Options = new List<ResponseOptions>();
    }

    public class TelegramConversation
    {
        private static NLP Pred = new NLP();

        private const int NumberOfResults = 5;

        public TelegramConversation(string ApiAiKey, string Language = "Spanish")
        {
            Pred.NLPStartApiAi(ApiAiKey, Language);
        }

        public TelegramConversation(string LuisURL, string LuisAppID, string LuisKey)
        {
            Pred.NLPStartLUIS(LuisURL, LuisAppID, LuisKey);
        }

        private static string cntSavvyNarrativeKey;
        public string SavvyNarrativeKey { set { cntSavvyNarrativeKey = value; } }

        private static string cntNarrativeScienceKey;
        public string NarrativeScienceKey { set { cntNarrativeScienceKey = value; } }


        public Response ReplySync(string inText, QSUser Usr)
        {
            Task<Response> ReplyTask = Reply(inText, Usr);
            ReplyTask.Wait();
            Response Resp = ReplyTask.Result;
            return Resp;
        }


        public async Task<Response> Reply(string inText, QSUser Usr)
        {
            Response Resp = new Response();

            string Dimension = "";
            string Dimension2 = "";
            string Measure = "";
            string Measure2 = "";
            string Element = "";
            string Percentage = "";
            double? Number = null;
            string ChartType = "";
            double DistanceKm = 0;
            string Language = "";

            string url = "";
            string UserChart = "";
            string UserSheet = "";

            if (Usr.TimeSincePreviousAccess().Minutes > 10)
            {
                Usr.TheBotIsAngry = false;
                Usr.QS.LastFilters.Clear();
            }


            if (await UnderstandSentence(Usr, inText))
            {
                Dimension = Pred.Dimension;
                Dimension2 = Pred.Dimension2;
                Measure = Pred.Measure;
                Measure2 = Pred.Measure2;
                Element = Pred.Element;
                Percentage = Pred.Percentage;
                Number = Pred.Number;
                ChartType = Pred.ChartType;
                DistanceKm = Pred.DistanceKm;
                Language = Pred.Language;

                try
                {
                    QSFoundFilter[] ff = new QSFoundFilter[1];

                    if (Usr.TheBotIsAngry && Pred.Intent != IntentType.Apologize)
                    {
                        Resp.TextMessage = StringResources.cnvApologizeOrNoAnswer;
                        return Resp;
                    }

                    switch (Pred.Intent)
                    {
                        case IntentType.DirectResponse:
                            Resp.TextMessage = Pred.Response;
                            break;

                        case IntentType.KPI:
                            if (Measure != null)
                            {
                                if (Usr.QS.LastFilters.Count() > 0)
                                {
                                    Resp.TextMessage = GetFilteredMeasure(Usr.QS, ref Measure, Usr.QS.LastFilters);
                                }
                                else
                                {
                                    Resp.TextMessage = GetKpiValue(Usr.QS, ref Measure, Usr.UserName);
                                }
                            }
                            else
                            {
                                Resp.TextMessage = StringResources.nlMeasureNotFound;
                                Resp.ErrorText = StringResources.nlMeasureNotFound;
                            }

                            break;

                        case IntentType.Chart:
                            if (Dimension != null && Measure != null)
                            {
                                Resp.TextMessage = GetChart(Usr.QS, ref Measure, ref Dimension, Usr.UserName, ref Resp.VoiceMessage, ref Resp.ChartFound);
                            }
                            else if (Measure != null)
                            {
                                Resp.TextMessage = StringResources.kpiChartWithoutDimension;
                                Resp.TextMessage += GetKpiValue(Usr.QS, ref Measure, Usr.UserName);
                            }
                            else if (Dimension != null)
                            {
                                Measure = Usr.QS.LastMeasure.Name;
                                Resp.TextMessage = GetChart(Usr.QS, ref Measure, ref Dimension, Usr.UserName, ref Resp.VoiceMessage, ref Resp.ChartFound);
                            }
                            Resp.NewsSearch = string.Format(StringResources.cnvShowChartNewsSearch, Measure, Dimension);
                            if (Usr.QS.LastFilters.Count > 0)
                            {
                                Resp.NewsSearch += " " + StringResources.nlWith + " ";
                                foreach (QSFoundFilter f in Usr.QS.LastFilters)
                                {
                                    if (f != Usr.QS.LastFilters.First()) Resp.NewsSearch += " " + StringResources.nlAnd + " ";
                                    Resp.NewsSearch += f.Element;
                                }
                            }
                            break;

                        case IntentType.Measure4Element:
                            if (Measure != null && Dimension != null && Element != null)
                            {
                                Usr.QS.AddFilter(Dimension, Element);
                                Resp.TextMessage = GetFilteredMeasure(Usr.QS, ref Measure, Usr.QS.LastFilters, Usr.UserName);
                                Resp.NewsSearch = string.Format(StringResources.cnvMeasure4ElementNewsSearch, Measure, Pred.Dimension, Pred.Element);
                                break;
                            }

                            if (Element != null)
                            {
                                ff = Usr.QS.QSSearch(Element);
                                if (ff.Length == 0 || ff.First() == null)
                                {
                                    Resp.TextMessage = StringResources.nlGetElementNotFound;
                                    break;
                                }
                            }
                            if (Measure == null) Measure = Usr.QS.LastMeasure.Name;

                            if (ff.First() != null && Measure != null)
                            {
                                Usr.QS.AddFilter(ff.First());

                                Resp.TextMessage = GetFilteredMeasure(Usr.QS, ref Measure, Usr.QS.LastFilters, Usr.UserName);
                                Resp.NewsSearch = string.Format(StringResources.cnvShowChartNewsSearch, Measure, Element);
                            }
                            else if (Element == null)
                            {
                                Resp.TextMessage = StringResources.kpiM4EWithoutElement;
                                Resp.TextMessage += GetKpiValue(Usr.QS, ref Measure, Usr.UserName);
                            }
                            break;

                        case IntentType.Filter:
                            if (Dimension != null && Element != null)
                            {
                                Resp.TextMessage = ApplyFilter(Usr.QS, Dimension, Element);
                            }
                            else if (Element != null)
                            {
                                ff = Usr.QS.QSSearch(Element);
                                if (ff.Length == 0 || ff.First() == null)
                                {
                                    Resp.TextMessage = StringResources.nlGetElementNotFound;
                                }
                                else
                                    Resp.TextMessage = ApplyFilter(Usr.QS, ff.First().Dimension, ff.First().Element);
                            }

                            break;

                        case IntentType.RankingTop:
                            if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                            if (Measure == null) Measure = Usr.QS.LastMeasure.Name;
                            if (Element != null)
                            {
                                ff = Usr.QS.QSSearch(Element);
                                if (ff.Length > 0 && ff.First() != null)
                                    ApplyFilter(Usr.QS, ff.First().Dimension, ff.First().Element);
                            }

                            Resp.TextMessage = GetRanking(Usr.QS, ref Measure, ref Dimension, Usr.UserName, true);
                            break;

                        case IntentType.RankingBottom:
                            if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                            if (Measure == null) Measure = Usr.QS.LastMeasure.Name;
                            if (Element != null)
                            {
                                ff = Usr.QS.QSSearch(Element);
                                if (ff.Length > 0 && ff.First() != null)
                                    ApplyFilter(Usr.QS, ff.First().Dimension, ff.First().Element);
                            }

                            Resp.TextMessage = GetRanking(Usr.QS, ref Measure, ref Dimension, Usr.UserName, false);

                            break;

                        case IntentType.Alert:
                            if (Percentage != null)
                            {
                                if (Measure == null) Measure = Usr.QS.LastMeasure.Name;

                                var a = new QlikSenseEasy.QSAlert();
                                a.UserName = Usr.UserName;
                                a.UserID = Usr.UserId;
                                a.SendToID = Usr.UserId;
                                a.AlertRequest = Measure + " " + StringResources.kpiHasChangedMoreThan + " " + Percentage;
                                Usr.QS.AlertList.Add(a);

                                Resp.TextMessage = string.Format(StringResources.alertAdknowledge, a.UserName, a.AlertRequest);
                            }
                            break;

                        case IntentType.GoodAnswer:
                            Resp.TextMessage = ":-)";
                            break;

                        case IntentType.BadWords:
                            Usr.TheBotIsAngry = true;
                            Resp.TextMessage = StringResources.cnvBadwords;
                            break;

                        case IntentType.Apologize:
                            Usr.TheBotIsAngry = false;
                            Resp.TextMessage = StringResources.cnvApologizeAnswer;
                            break;

                        case IntentType.Hello:
                            Resp.TextMessage = string.Format(StringResources.Welcome, Usr.UserName);
                            break;

                        case IntentType.Bye:
                            Resp.TextMessage = StringResources.Bye + " " + Usr.UserName;
                            Resp.TextMessage += "\n👋";
                            break;

                        case IntentType.Reports:
                            Resp.OtherAction = "ShowReports";
                            break;

                        case IntentType.CreateChart:
                            if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                            if (Dimension2 == null) Dimension2 = "";
                            if (Measure == null) Measure = Usr.QS.LastMeasure.Name;
                            if (ChartType == null) ChartType = "BarChart";


                            UserChart = CreateChart(Usr, ChartType, Measure, Dimension, Dimension2, Element: Element);
                            if (UserChart == "forbidden")
                                Resp.TextMessage = Resp.ErrorText = StringResources.cnvCreateChartNoRights;
                            else if (UserChart == "")
                                Resp.TextMessage = Resp.ErrorText = StringResources.cnvCreateChartError;
                            else
                            {
                                //url = Usr.QS.qsServer + "/single?appid=" + Usr.QS.qsAppId + "&obj=" + UserChart + "&opt=currsel";
                                url = Usr.QS.qsSingleServer + "/single?appid=" + Usr.QS.qsSingleApp + "&obj=" + UserChart + "&opt=currsel";
                                if (Element != null)
                                {
                                    ff = Usr.QS.QSSearch(Element);
                                    if (ff.Length > 0 && ff.First() != null)
                                    {
                                        url += "&select=" + Uri.EscapeDataString(ff.First().Dimension) + "," + Uri.EscapeDataString(ff.First().Element);
                                    }
                                }
                                else
                                {
                                    foreach (QSFoundFilter f in Usr.QS.LastFilters)
                                    {
                                        url += "&select=" + Uri.EscapeDataString(f.Dimension) + "," + Uri.EscapeDataString(f.Element);
                                    }
                                }
                                Resp.TextMessage = string.Format(StringResources.cnvCreateChartResult, Usr.UserName, url);
                                string strDim2 = Dimension2 == "" ? "" : (" " + StringResources.nlAnd + " " + Dimension2);
                                Resp.ChartFound = new QSFoundObject
                                {
                                    ObjectURL = url,
                                    Description = Measure + " " + StringResources.By + " " + Dimension + strDim2
                                };
                            }
                            break;

                        case IntentType.ShowAnalysis:
                            Usr.SummarizeHistory();
                            UserSheet = CreateSummarySheet(Usr);
                            if (UserSheet == "forbidden")
                                Resp.TextMessage = Resp.ErrorText = StringResources.cnvAnalysisNoRights;
                            else if (UserSheet == "")
                                Resp.TextMessage = Resp.ErrorText = StringResources.cnvAnalysisError;
                            else
                            {
                                url = Usr.QS.qsSingleServer;
                                url += "/sense/app/" + Usr.QS.qsSingleApp + "/sheet/" + UserSheet + "/state/analysis";

                                Resp.TextMessage = string.Format(StringResources.cnvAnalysisResult, Usr.UserName, url);
                            }

                            break;

                        case IntentType.GeoFilter:
                            Resp.OtherAction = "GeoFilter";
                            break;

                        case IntentType.ContactQlik:
                            url = "http://www.qlik.com/us/try-or-buy/buy-now";
                            Resp.TextMessage = string.Format(StringResources.cnvContactQlik, Usr.UserName, url);
                            break;

                        case IntentType.Help:
                            Resp.TextMessage = StringResources.UsageInstructions;
                            Resp.TextMessage += "\r\n\r\nQlik Sense App: " + Usr.QS.qsAppName;
                            Resp.TextMessage += "\r\n\r\nQlikSenseBot by Juan Gerardo(jcz) @ Qlik";
                            Resp.OtherAction = "Help";
                            break;

                        case IntentType.PersonalInformation:
                            url = "https://www.facebook.com/profile.php?id=xxxxxx";
                            Resp.TextMessage = string.Format(StringResources.cnvPersonalInformation, Usr.UserName, url);
                            break;

                        case IntentType.ClearAllFilters:
                            Resp = ClearFilters(Usr);
                            break;

                        case IntentType.ClearDimensionFilter:
                            if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                            Resp = ClearFilters(Usr, Dimension);
                            break;

                        case IntentType.CurrentSelections:
                            if (Usr.QS.LastFilters != null && Usr.QS.LastFilters.Count() > 0)
                            {
                                string strFilter = "";

                                foreach (QSFoundFilter qsFF in Usr.QS.LastFilters)
                                {
                                    if (qsFF != Usr.QS.LastFilters.First()) strFilter += ", " + StringResources.nlAnd + " ";
                                    strFilter += String.Format(StringResources.nlGetElementMoreFilter, qsFF.Dimension, qsFF.Element);
                                }
                                Resp.TextMessage = string.Format(StringResources.cnvCurrentSelections, strFilter);
                            }
                            else
                            {
                                Resp.TextMessage = StringResources.cnvCurrentSelectionsEmpty;
                            }
                            break;

                        case IntentType.ShowAllApps:
                            foreach (QSAppProperties qa in Usr.QS.qsAlternativeApps)
                            {
                                Resp.Options.Add(
                                    new ResponseOptions()
                                    { Title = qa.AppName, ID = qa.AppID, Action = ResponseAction.OpenApp }
                                    );
                            }
                            if (Resp.Options.Count() > 0) Resp.TextMessage = StringResources.appSelectApp;
                            break;

                        case IntentType.ShowAllDimensions:
                            foreach (QSMasterItem d in Usr.QS.MasterDimensions)
                            {
                                Resp.Options.Add(
                                    new ResponseOptions()
                                    { Title = d.Name, ID = d.Name, Action = ResponseAction.ShowDimension }
                                    );
                            }
                            if (Resp.Options.Count() > 0) Resp.TextMessage = StringResources.kpiSelectDimension;
                            break;

                        case IntentType.ShowAllMeasures:
                            foreach (QSMasterItem m in Usr.QS.MasterMeasures)
                            {
                                Resp.Options.Add(
                                    new ResponseOptions()
                                    { Title = m.Name, ID = m.Name, Action = ResponseAction.ShowMeasure }
                                    );
                            }
                            if (Resp.Options.Count() > 0) Resp.TextMessage = StringResources.kpiSelectMeasure;
                            break;

                        case IntentType.ShowAllSheets:
                            foreach (QSSheet s in Usr.QS.Sheets)
                            {
                                Resp.Options.Add(
                                    new ResponseOptions()
                                    { Title = s.Name, ID = s.Id, Action = ResponseAction.ShowSheet }
                                    );
                            }
                            if (Resp.Options.Count() > 0) Resp.TextMessage = StringResources.kpiSelectSheet;
                            break;

                        case IntentType.ShowAllStories:
                            foreach (QSStory s in Usr.QS.Stories)
                            {
                                Resp.Options.Add(
                                    new ResponseOptions()
                                    { Title = s.Name, ID = s.Id, Action = ResponseAction.ShowStory }
                                    );
                            }
                            if (Resp.Options.Count() > 0) Resp.TextMessage = StringResources.kpiSelectStory;
                            break;

                        case IntentType.ShowKPIs:
                            foreach (QSMasterItem m in Usr.QS.MasterMeasures)
                            {
                                Resp.Options.Add(
                                    new ResponseOptions()
                                    { Title = m.Name, ID = m.Name, Action = ResponseAction.ShowMeasure }
                                    );
                                if (Resp.Options.Count > 5) break;
                            }
                            Resp.Options.Add(
                                new ResponseOptions()
                                { Title = StringResources.kpiAnalysis, ID = Usr.QS.Sheets.First().Id, Action = ResponseAction.ShowSheet }
                                );
                            if (Resp.Options.Count() > 0) Resp.TextMessage = StringResources.kpiSelectMeasure;
                            break;

                        case IntentType.ShowMeasureByMeasure:
                            if (Measure != null && Measure2 != null)
                            {
                                if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                                if (Measure == null) Measure = Usr.QS.LastMeasure.Name;
                                if (Measure2 == null) Measure2 = Usr.QS.LastMeasure.Name;
                                ChartType = "ScatterChart";

                                UserChart = CreateChart(Usr, ChartType, Measure, Dimension,
                                    Measure2: Measure2, Element: Element);
                                if (UserChart == "forbidden")
                                    Resp.TextMessage = Resp.ErrorText = StringResources.cnvCreateChartNoRights;
                                else if (UserChart == "")
                                    Resp.TextMessage = Resp.ErrorText = StringResources.cnvCreateChartError;
                                else
                                {
                                    url = Usr.QS.qsSingleServer + "/single?appid=" + Usr.QS.qsSingleApp + "&obj=" + UserChart + "&opt=currsel";
                                    if (Element != null)
                                    {
                                        ff = Usr.QS.QSSearch(Element);
                                        if (ff.Length > 0 && ff.First() != null)
                                        {
                                            url += "&select=" + Uri.EscapeDataString(ff.First().Dimension) + "," + Uri.EscapeDataString(ff.First().Element);
                                        }
                                    }
                                    else
                                    {
                                        foreach (QSFoundFilter f in Usr.QS.LastFilters)
                                        {
                                            url += "&select=" + Uri.EscapeDataString(f.Dimension) + "," + Uri.EscapeDataString(f.Element);
                                        }
                                    }
                                    Resp.TextMessage = string.Format(StringResources.cnvCreateChartResult, Usr.UserName, url);
                                    Resp.ChartFound = new QSFoundObject { ObjectURL = url, Description = Measure + " vs " + Measure2 + " " + StringResources.By + " " + Dimension };
                                }
                                break;
                            }
                            else
                            {
                                Resp.TextMessage = StringResources.nlMeasureNotFound;
                                Resp.ErrorText = StringResources.nlMeasureNotFound;
                            }

                            break;

                        case IntentType.ShowListOfElements:
                            if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                            if (Element != null)
                            {
                                ff = Usr.QS.QSSearch(Element);
                                if (ff.Length > 0 && ff.First() != null)
                                {
                                    Usr.QS.AddFilter(ff.First());
                                }
                            }

                            Resp.TextMessage = GetFilteredList(Usr.QS, ref Dimension, Usr.QS.LastFilters, Usr.UserName);

                            break;

                        case IntentType.ChangeLanguage:
                            switch (Language.ToLower())
                            {
                                case "español":
                                case "spanish":
                                    Resp.OtherAction = "Language=es-ES";
                                    Resp.TextMessage = "Cambiado el idioma a Español";
                                    break;
                                case "english":
                                    Resp.OtherAction = "Language=en-US";
                                    Resp.TextMessage = "Language changed to English";
                                    break;
                                case "português":
                                case "portuguese":
                                    Resp.OtherAction = "Language=pt-PT";
                                    Resp.TextMessage = "Linguagem mudado para Português";
                                    break;
                                case "italiano":
                                case "italian":
                                    Resp.OtherAction = "Language=it-IT";
                                    Resp.TextMessage = "Cambiato la lingua Italiana per";
                                    break;
                                case "pусский":
                                case "russian":
                                    Resp.OtherAction = "Language=ru-RU";
                                    Resp.TextMessage = "Язык изменен на Pусский";
                                    break;
                                case "français":
                                case "french":
                                    Resp.OtherAction = "Language=fr-FR";
                                    Resp.TextMessage = "Langue changée en Français";
                                    break;
                                case "português-br":
                                case "brazilian":
                                    Resp.OtherAction = "Language=pt-BR";
                                    Resp.TextMessage = "Linguagem mudado para Português-BR";
                                    break;
                            }
                            break;

                        case IntentType.ShowElementsAboveValue:
                            if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                            if (Measure == null) Measure = Usr.QS.LastMeasure.Name;

                            Resp.TextMessage = GetRanking(Usr.QS, ref Measure, ref Dimension, Usr.UserName, true, Number);

                            break;

                        case IntentType.ShowElementsBelowValue:
                            if (Dimension == null) Dimension = Usr.QS.LastDimension.Name;
                            if (Measure == null) Measure = Usr.QS.LastMeasure.Name;

                            Resp.TextMessage = GetRanking(Usr.QS, ref Measure, ref Dimension, Usr.UserName, false, Number);

                            break;

                        case IntentType.CreateCollaborationGroup:
                            if (Usr.LastChart == null)
                                Resp.TextMessage = StringResources.cnvCreateGroupNoChart;
                            else
                            {
                                Resp.OtherAction = "CreateGroup";
                            }

                            break;

                    }
                }
                catch (Exception e)
                {
                    Resp.ErrorText = string.Format("Error trying to predict '{0}': {1}", inText, e);
                }

            }
            else
            {
                Resp.TextMessage = StringResources.nlMessageNotManaged;
                Resp.WarningText = string.Format("Message not managed: {0}", inText);

            }

            Pred.Dimension = Dimension;
            Pred.Measure = Measure;
            Pred.Measure2 = Measure2;
            Pred.Element = Element;
            Pred.Percentage = Percentage;
            Pred.Number = Number;
            Pred.ChartType = ChartType;
            Pred.DistanceKm = DistanceKm;

            Resp.NLPrediction = Pred;

            PredictedIntent pi = Pred.GetCopyOfPredictedIntent();
            Usr.ConversationHistory.Add(pi);
            Usr.LastResponse = Resp;
            if (Resp.ChartFound != null)
                Usr.LastChart = Resp.ChartFound;

            return Resp;
        }


        // Proccess an action
        public Response ProcessAction(QSUser Usr, ResponseAction Action, string Value)
        {
            Response Resp = new Response();

            switch (Action)
            {
                case ResponseAction.OpenApp:
                    if (Value.Trim().Length > 0)
                    {
                        string AppId = Value.Trim();
                        Usr.QS.QSOpenApp(AppId);
                        if (Usr.QS.qsAppId == AppId)
                            Resp.TextMessage = string.Format(StringResources.appOpenedApp, Usr.QS.qsAppName);
                        else
                            Resp.TextMessage = string.Format(StringResources.appOpenAppError, Usr.QS.qsAppId);
                    }
                    break;
            }

            return Resp;
        }


        public Response ShowAMeasure(string MeasureName, QSUser Usr)
        {
            Response Resp = new Response();
            try
            {
                Resp.TextMessage = GetMeasureValue(Usr.QS, MeasureName, Usr.UserName);
            }
            catch (Exception e)
            {
                Resp.ErrorText = string.Format("Error in ShowAMeasure() for measure '{0}': {1}", MeasureName, e);
            }
            finally
            {
                Usr.ConversationHistory.Add(new PredictedIntent
                {
                    Intent = IntentType.KPI,
                    Measure = MeasureName
                });
            }
            return Resp;
        }

        public Response ShowMeasureByDimension(ref string MeasureName, ref string DimensionName, QSUser Usr)
        {
            Response Resp = new Response();
            try
            {
                Resp.TextMessage = GetChart(Usr.QS, ref MeasureName, ref DimensionName, Usr.UserName, ref Resp.VoiceMessage, ref Resp.ChartFound);
            }
            catch (Exception e)
            {
                Resp.ErrorText = string.Format("Error in ShowMeasureByDimension() for measure '{0}' and dimension {1}: {2}", MeasureName, DimensionName, e);
            }
            finally
            {
                Usr.ConversationHistory.Add(new PredictedIntent
                {
                    Intent = IntentType.Chart,
                    Measure = MeasureName,
                    Dimension = DimensionName
                });
            }
            return Resp;
        }

        public Response ShowSheet(string SheetID, QSUser Usr)
        {
            Response Resp = new Response();
            try
            {
                string url = Usr.QS.qsSingleServer + "/sense/app/" + Usr.QS.qsSingleApp
                    + "/sheet/" + SheetID + "/state/play";
                Resp.TextMessage = string.Format(StringResources.kpiShowSheet, Usr.UserName, url);
            }
            catch (Exception e)
            {
                Resp.ErrorText = string.Format("Error in ShowSheet() for SheetID '{0}': {1}", SheetID, e);
            }
            return Resp;
        }

        public Response ShowStory(string StoryID, QSUser Usr)
        {
            Response Resp = new Response();
            try
            {
                string url = Usr.QS.qsSingleServer + "/sense/app/" + Usr.QS.qsSingleApp
                    + "/story/" + StoryID + "/state/play"; // Open the Story
                Resp.TextMessage = string.Format(StringResources.kpiShowStory, Usr.UserName, url);
            }
            catch (Exception e)
            {
                Resp.ErrorText = string.Format("Error in ShowStory() for StoryID '{0}': {1}", StoryID, e);
            }

            return Resp;
        }

        public Response ClearFilters(QSUser Usr, string DimensionName = "")
        {
            Response Resp = new Response();
            try
            {
                string DimensionExpression = null;
                if (DimensionName != "")
                    DimensionExpression = Usr.QS.GetDimensionExpression(DimensionName);

                if (DimensionExpression == null)
                {
                    Usr.QS.LastFilters.Clear();
                    Resp.TextMessage = StringResources.nlClearFilters;
                }
                else
                {
                    foreach (QSFoundFilter ff in Usr.QS.LastFilters)
                    {
                        if (ff.Dimension.Trim().ToLower() == DimensionName.Trim().ToLower())
                        {
                            Usr.QS.LastFilters.Remove(ff);
                            Resp.TextMessage = String.Format(StringResources.nlClearOneFilter, ff.Dimension);
                            break;
                        }
                        Resp.TextMessage = String.Format(StringResources.nlNoFilterToClear, DimensionName);
                    }
                }
            }
            catch (Exception e)
            {
                Resp.ErrorText = string.Format("Error in ClearFilters() for DimensionName '{0}': {1}", DimensionName, e);
            }

            return Resp;
        }


        public string CreateSummarySheet(QSUser Usr)
        {
            string SheetID = "BotSumary-" + Usr.UserId;
            if (Usr.QS.MasterMeasures.Count < 4 || Usr.QS.MasterDimensions.Count < 4) return "";

            try
            {
                string Meas0 = (Usr.ConversationSummary.Measure.Count > 0) ? Usr.ConversationSummary.Measure[0].Measure : Usr.QS.MasterMeasures[0].Name;
                string Meas1 = (Usr.ConversationSummary.Measure.Count > 1) ? Usr.ConversationSummary.Measure[1].Measure : Usr.QS.MasterMeasures[1].Name;
                string Meas2 = (Usr.ConversationSummary.Measure.Count > 2) ? Usr.ConversationSummary.Measure[2].Measure : Usr.QS.MasterMeasures[2].Name;
                string Meas3 = (Usr.ConversationSummary.Measure.Count > 3) ? Usr.ConversationSummary.Measure[3].Measure : Usr.QS.MasterMeasures[3].Name;

                string MxDDim0 = (Usr.ConversationSummary.MeasureByDimension.Count > 0) ? Usr.ConversationSummary.MeasureByDimension[0].Dimension : Usr.QS.MasterDimensions[0].Name;
                string MxDDim1 = (Usr.ConversationSummary.MeasureByDimension.Count > 1) ? Usr.ConversationSummary.MeasureByDimension[1].Dimension : Usr.QS.MasterDimensions[1].Name;
                string MxDDim2 = (Usr.ConversationSummary.MeasureByDimension.Count > 2) ? Usr.ConversationSummary.MeasureByDimension[2].Dimension : Usr.QS.MasterDimensions[2].Name;
                string MxDDim3 = (Usr.ConversationSummary.MeasureByDimension.Count > 3) ? Usr.ConversationSummary.MeasureByDimension[3].Dimension : Usr.QS.MasterDimensions[3].Name;

                string MxDMeas0 = (Usr.ConversationSummary.MeasureByDimension.Count > 0) ? Usr.ConversationSummary.MeasureByDimension[0].Measure : Usr.QS.MasterMeasures[0].Name;
                string MxDMeas1 = (Usr.ConversationSummary.MeasureByDimension.Count > 1) ? Usr.ConversationSummary.MeasureByDimension[1].Measure : Usr.QS.MasterMeasures[1].Name;
                string MxDMeas2 = (Usr.ConversationSummary.MeasureByDimension.Count > 2) ? Usr.ConversationSummary.MeasureByDimension[2].Measure : Usr.QS.MasterMeasures[2].Name;
                string MxDMeas3 = (Usr.ConversationSummary.MeasureByDimension.Count > 3) ? Usr.ConversationSummary.MeasureByDimension[3].Measure : Usr.QS.MasterMeasures[3].Name;

                string FilDim0 = (Usr.ConversationSummary.Filter.Count > 0) ? Usr.ConversationSummary.Filter[0].Dimension : Usr.QS.MasterDimensions[0].Name;
                string FilDim1 = (Usr.ConversationSummary.Filter.Count > 1) ? Usr.ConversationSummary.Filter[1].Dimension : Usr.QS.MasterDimensions[1].Name;
                string FilDim2 = (Usr.ConversationSummary.Filter.Count > 2) ? Usr.ConversationSummary.Filter[2].Dimension : Usr.QS.MasterDimensions[2].Name;
                string FilDim3 = (Usr.ConversationSummary.Filter.Count > 3) ? Usr.ConversationSummary.Filter[3].Dimension : Usr.QS.MasterDimensions[3].Name;

                string MxEDim0 = (Usr.ConversationSummary.MeasureForElement.Count > 0) ? Usr.ConversationSummary.MeasureForElement[0].Dimension : Usr.QS.MasterDimensions[0].Name;
                string MxEDim1 = (Usr.ConversationSummary.MeasureForElement.Count > 1) ? Usr.ConversationSummary.MeasureForElement[1].Dimension : Usr.QS.MasterDimensions[1].Name;
                string MxEDim2 = (Usr.ConversationSummary.MeasureForElement.Count > 2) ? Usr.ConversationSummary.MeasureForElement[2].Dimension : Usr.QS.MasterDimensions[2].Name;
                string MxEDim3 = (Usr.ConversationSummary.MeasureForElement.Count > 3) ? Usr.ConversationSummary.MeasureForElement[3].Dimension : Usr.QS.MasterDimensions[3].Name;

                string MxEMeas0 = (Usr.ConversationSummary.MeasureForElement.Count > 0) ? Usr.ConversationSummary.MeasureForElement[0].Measure : Usr.QS.MasterMeasures[0].Name;
                string MxEMeas1 = (Usr.ConversationSummary.MeasureForElement.Count > 1) ? Usr.ConversationSummary.MeasureForElement[1].Measure : Usr.QS.MasterMeasures[1].Name;
                string MxEMeas2 = (Usr.ConversationSummary.MeasureForElement.Count > 2) ? Usr.ConversationSummary.MeasureForElement[2].Measure : Usr.QS.MasterMeasures[2].Name;
                string MxEMeas3 = (Usr.ConversationSummary.MeasureForElement.Count > 3) ? Usr.ConversationSummary.MeasureForElement[3].Measure : Usr.QS.MasterMeasures[3].Name;

                string TopDim0 = (Usr.ConversationSummary.Top.Count > 0) ? Usr.ConversationSummary.Top[0].Dimension : Usr.QS.MasterDimensions[0].Name;
                string TopDim1 = (Usr.ConversationSummary.Top.Count > 1) ? Usr.ConversationSummary.Top[1].Dimension : Usr.QS.MasterDimensions[1].Name;
                string TopDim2 = (Usr.ConversationSummary.Top.Count > 2) ? Usr.ConversationSummary.Top[2].Dimension : Usr.QS.MasterDimensions[2].Name;
                string TopDim3 = (Usr.ConversationSummary.Top.Count > 3) ? Usr.ConversationSummary.Top[3].Dimension : Usr.QS.MasterDimensions[3].Name;

                string TopMeas0 = (Usr.ConversationSummary.Top.Count > 0) ? Usr.ConversationSummary.Top[0].Measure : Usr.QS.MasterMeasures[0].Name;
                string TopMeas1 = (Usr.ConversationSummary.Top.Count > 1) ? Usr.ConversationSummary.Top[1].Measure : Usr.QS.MasterMeasures[1].Name;
                string TopMeas2 = (Usr.ConversationSummary.Top.Count > 2) ? Usr.ConversationSummary.Top[2].Measure : Usr.QS.MasterMeasures[2].Name;
                string TopMeas3 = (Usr.ConversationSummary.Top.Count > 3) ? Usr.ConversationSummary.Top[3].Measure : Usr.QS.MasterMeasures[3].Name;

                string BottomDim0 = (Usr.ConversationSummary.Bottom.Count > 0) ? Usr.ConversationSummary.Bottom[0].Dimension : Usr.QS.MasterDimensions[0].Name;
                string BottomDim1 = (Usr.ConversationSummary.Bottom.Count > 1) ? Usr.ConversationSummary.Bottom[1].Dimension : Usr.QS.MasterDimensions[1].Name;
                string BottomDim2 = (Usr.ConversationSummary.Bottom.Count > 2) ? Usr.ConversationSummary.Bottom[2].Dimension : Usr.QS.MasterDimensions[2].Name;
                string BottomDim3 = (Usr.ConversationSummary.Bottom.Count > 3) ? Usr.ConversationSummary.Bottom[3].Dimension : Usr.QS.MasterDimensions[3].Name;

                string BottomMeas0 = (Usr.ConversationSummary.Bottom.Count > 0) ? Usr.ConversationSummary.Bottom[0].Measure : Usr.QS.MasterMeasures[0].Name;
                string BottomMeas1 = (Usr.ConversationSummary.Bottom.Count > 1) ? Usr.ConversationSummary.Bottom[1].Measure : Usr.QS.MasterMeasures[1].Name;
                string BottomMeas2 = (Usr.ConversationSummary.Bottom.Count > 2) ? Usr.ConversationSummary.Bottom[2].Measure : Usr.QS.MasterMeasures[2].Name;
                string BottomMeas3 = (Usr.ConversationSummary.Bottom.Count > 3) ? Usr.ConversationSummary.Bottom[3].Measure : Usr.QS.MasterMeasures[3].Name;


                Usr.QS.QSCreateSheet(SheetID, string.Format(StringResources.cnvAnalysisSheetTitle, Usr.UserName),
                    string.Format("Sheet for Bot user {0}", Usr.UserName));

                Usr.QS.QSCreateKPI(SheetID, SheetID + "kpi1", Meas0, Meas2);

                Usr.QS.QSCreateKPI(SheetID, SheetID + "kpi2", Meas1, Meas3);

                Usr.QS.QSCreatePivotTableChart(SheetID, SheetID + "Pivotchart1", MxEDim0, MxEDim1, MxEDim2, MxEMeas0, MxEMeas1, MxEMeas2
                    , MxEMeas0 + " " + StringResources.By + " " + MxEDim0);

                Usr.QS.QSCreateBarChart(SheetID, SheetID + "Barchart1", MxDDim0, MxDMeas0, MxDMeas0 + " " + StringResources.By + " " + MxDDim0);
                Usr.QS.QSCreateBarChart(SheetID, SheetID + "Barchart2", MxDDim1, MxDMeas1, MxDMeas1 + " " + StringResources.By + " " + MxDDim1);

                Usr.QS.QSCreateLineChart(SheetID, SheetID + "Linechart1", MxDDim2, MxDMeas2, MxDMeas2 + " " + StringResources.By + " " + MxDDim2);
                if (Usr.QS.IsGeoLocationActive())
                    Usr.QS.QSCreateMapPoint(SheetID, SheetID + "mappoint1", MxDDim0, MxDMeas0, MxDMeas0 + " " + StringResources.By + " " + MxDDim0);
                //Usr.QS.QSCreateTreeChart(SheetID, SheetID + "Treechart1", TopDim0, TopDim1, TopMeas0, TopMeas0 + " " + StringResources.By + " " + TopDim0 + " " + StringResources.By + " " + TopDim1);
                Usr.QS.QSCreateTreeChart(SheetID, SheetID + "Treechart1", TopDim0, TopMeas0, TopMeas0 + " " + StringResources.By + " " + TopDim0);
                Usr.QS.QSCreateFilterPane(SheetID, SheetID + "filter1", FilDim0, FilDim1, FilDim2, FilDim0 + " - " + FilDim1 + " - " + FilDim2);

                Usr.QS.QSCreatePieChart(SheetID, SheetID + "Piechart1", MxDDim3, MxDMeas3, MxDMeas3 + " " + StringResources.By + " " + MxDDim3);

                Usr.QS.QSCreateTextImage(SheetID, SheetID + "Text2", StringResources.cnvAnalysisSheetTextboxTitle,
                    StringResources.cnvAnalysisSheetTextboxText);

                Usr.QS.QSDoSave();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                SheetID = "";

                if (e.Message.ToLower().StartsWith("forbidden")) SheetID = "forbidden";
            }

            return SheetID;
        }


        public string CreateChart(QSUser Usr, string ChartType, string Measure, string Dimension
            , string Dimension2 = "", string Dimension3 = "", string Measure2 = "", string Measure3 = "", string Element = "")
        {
            string SheetID = "BotChart-" + Usr.UserId;
            string ChartID = SheetID + "-UserChart";

            try
            {
                Usr.QS.QSRemoveSheet(SheetID);
                Usr.QS.QSDoSave();

                Usr.QS.QSCreateSheet(SheetID, string.Format(StringResources.cnvCreateChartSheetTitle, Usr.UserName),
                    string.Format(StringResources.cnvCreateChartSheetDescription, Usr.UserName));

                ChartType = ChartType.Trim().ToLower();
                if (ChartType.Contains("barchart"))
                    Usr.QS.QSCreateBarChart(SheetID, ChartID, Dimension, Measure);

                else if (ChartType.Contains("linechart"))
                    Usr.QS.QSCreateLineChart(SheetID, ChartID, Dimension, Measure);

                else if (ChartType.Contains("piechart"))
                    Usr.QS.QSCreatePieChart(SheetID, ChartID, Dimension, Measure);

                else if (ChartType.Contains("treemap"))
                    Usr.QS.QSCreateTreeChart(SheetID, ChartID, Dimension, Measure);
                //Usr.QS.QSCreateTreeChart(SheetID, ChartID, Dimension, Dimension2, Measure);

                else if (ChartType.Contains("pivotchart"))
                    Usr.QS.QSCreatePivotTableChart(SheetID, ChartID, Dimension, Dimension2, Dimension3, Measure);

                else if (ChartType.Contains("mapchart"))
                    Usr.QS.QSCreateMapPoint(SheetID, ChartID, Dimension, Measure);

                else if (ChartType.Contains("scatterchart"))
                    Usr.QS.QSCreateScatterChart(SheetID, ChartID, Dimension, Measure, Measure2, Measure3);

                else
                    Usr.QS.QSCreateBarChart(SheetID, ChartID, Dimension, Measure);

                Usr.QS.QSDoSave();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                SheetID = "";
                ChartID = "";

                if (e.Message.ToLower().StartsWith("forbidden")) ChartID = "forbidden";
            }

            return ChartID;
        }



        private async Task<bool> UnderstandSentence(QSUser Usr, string Message)
        {
            bool Predicted;

            Predicted = DetectDirectCommand(Message);

            if (!Predicted)
                Predicted = await Pred.Predict(Message);

            if (Predicted && Pred.Intent != IntentType.None)
                return Predicted;

            string[] Words = Message.Split(' ');
            QSMasterItem dim = null;
            QSMasterItem mea = null;

            foreach (string w in Words)
            {
                if (dim == null)
                    dim = Usr.QS.MasterDimensions.Find(d => d.Name.ToLower().Trim() == w.ToLower().Trim());
                if (mea == null)
                    mea = Usr.QS.MasterMeasures.Find(m => m.Name.ToLower().Trim() == w.ToLower().Trim());
            }

            if (dim != null) Pred.Dimension = dim.Name;
            if (mea != null) Pred.Measure = mea.Name;

            if (dim != null && mea != null)
            {
                Pred.Intent = IntentType.Chart;
                Predicted = true;
            }
            else if (dim == null && mea != null)
            {
                Pred.Intent = IntentType.KPI;
                Predicted = true;
            }
            else if (dim != null && mea == null)
            {
                Pred.Intent = IntentType.Chart;
                Pred.Measure = Usr.QS.LastMeasure.Name;
                Predicted = true;
            }
            else
            {
                Pred.Intent = IntentType.None;
                Predicted = false;
            }

            return Predicted;
        }

        private static bool DetectDirectCommand(string Message)
        {
            bool Detected = false;

            if (Message.StartsWith("#") && Message.EndsWith("#"))
            {
                Detected = true;
                Message = Message.Replace("#", "");
                if (Message == IntentType.ShowAnalysis.ToString())
                {
                    Pred.Intent = IntentType.ShowAnalysis;
                }
                else
                    Pred.Intent = IntentType.None;
            }
            else
                Detected = false;

            return Detected;
        }


        private static string GetMeasureValue(QSApp qa, string Measure, string UserName = "")
        {
            string msg;
            QSMasterItem m;

            m = qa.GetMasterMeasure(Measure);

            if (Measure == null || m == null)
            {
                msg = string.Format(StringResources.kpiUnknownMessage, Measure);
            }
            else
            {
                msg = string.Format(StringResources.kpiValueMessage, m.Name, m.FormattedExpression);

                qa.MasterMeasures.Remove(m);
                qa.MasterMeasures.Insert(0, m);
            }

            return msg;
        }

        private static string GetKpiValue(QSApp qa, ref string Kpi, string UserName = "")
        {
            string msg;

            if (Kpi == null || Kpi.Trim() == "")
            {
                msg = string.Format(StringResources.kpiUnknownMessage, Kpi);
            }
            else
            {
                QSMasterItem m = qa.GetMasterMeasure(Kpi);
                if (m != null)
                {
                    Kpi = m.Name;
                    msg = GetMeasureValue(qa, Kpi, UserName);

                    qa.MasterMeasures.Remove(m);
                    qa.MasterMeasures.Insert(0, m);
                }
                else
                {
                    msg = string.Format(StringResources.kpiUnknownMessage, Kpi);
                }
            }

            return msg;
        }

        private static string GetChart(QSApp qa, ref string Measure, ref string Dimension, string UserName
            , ref string MsgToSpeak, ref QSFoundObject ChartFound)
        {
            string msg = "";
            QSFoundObject[] qsFounds;
            string qsQuery;

            string AppDimension;
            QSMasterItem d = qa.GetMasterDimension(Dimension);
            if (d != null)
            {
                AppDimension = d.Expression;
            }
            else
            {
                AppDimension = qa.QSFindField(Dimension);
            }

            string AppExpression = "";
            QSMasterItem m = qa.GetMasterMeasure(Measure);
            if (m != null)
            {
                AppExpression = m.Expression;
            }

            qsQuery = Measure + " " + StringResources.By + " " + Dimension;

            if (AppDimension.Length > 0 && AppExpression.Length > 0)
            {
                Measure = m.Name;
                Dimension = d.Name;

                try
                {
                    qsFounds = qa.QSSearchObjects(qsQuery, true);
                }
                catch (Exception e)
                {
                    //BotLog.AddBotLine(string.Format("{0} Exception caught.", e), LogFile.LogType.logError);
                    qsFounds = new QSFoundObject[0];
                }

                QSDataList[] dl = qa.GetDataList(AppExpression, AppDimension, NoOfRows: 50);
                qa.LastDimension = d;
                qa.LastMeasure = m;

                if (dl.Count() > 0)
                {
                    msg += InterpretData(qa, dl, Measure, Dimension, AppExpression, AppDimension);

                    string msgNLG = "";

                    if (msgNLG != "")
                    {
                        msg += "\n\n" + msgNLG;
                        MsgToSpeak = msgNLG;
                    }

                }

                if (qsFounds.Length > 0)
                {
                    QSFoundObject f = qsFounds.First();

                    msg += "\n\n";
                    msg += string.Format(StringResources.nlGetChart, f.Description, f.HRef);
                    ChartFound = f;
                }
            }
            else
            {
                msg = StringResources.nlGetChartNotFound;
            }

            return msg;
        }

        private static string GetTextChart(QSDataList[] Data, int NumBuckets = 5)
        {
            string TextChart = "";
            string BarText = "";
            int BarUnits = 0;
            int i = 0;

            int NofElements = Data.Length;
            double MinVal = 0;
            double MaxVal = 0;

            i = 0;
            foreach (QSDataList d in Data)
            {
                i += 1;
                if (i > NumBuckets) break;
                if (i == 1 || d.MeasValue < MinVal) MinVal = d.MeasValue;
                if (i == 1 || d.MeasValue > MaxVal) MaxVal = d.MeasValue;
            }

            i = 0;
            foreach (QSDataList d in Data)
            {
                i += 1;
                if (i > NumBuckets) break;

                BarText = "#";
                BarUnits = 0;
                if (MaxVal > MinVal && !double.IsNaN(d.MeasValue))
                    BarUnits = Convert.ToInt32(Math.Round((d.MeasValue - MinVal) / (MaxVal - MinVal) * (NumBuckets)));
                BarText += "#".PadRight(BarUnits, '#');

                if (i > 1) TextChart += "\r\n";
                //TextChart += BarText + "\t";
                TextChart += BarText.PadRight(NumBuckets + 4, ' ');
                TextChart += d.DimValue + " (" + d.MeasFormattedValue + ")";
            }

            return TextChart;
        }

        private static string GetRanking(QSApp qa, ref string Measure, ref string Dimension, string UserName,
            bool Descending = true, double? MeasureThreshold = null)
        {
            string msg = "";

            string AppDimension;
            QSMasterItem d = qa.GetMasterDimension(Dimension);
            if (d == null)
                AppDimension = qa.GetDimensionExpression(Dimension);
            else
            {
                AppDimension = d.Expression;
            }

            string AppExpression = "";
            QSMasterItem m = qa.GetMasterMeasure(Measure);
            if (m != null)
            {
                AppExpression = m.Expression;
            }

            if (AppDimension.Length > 0 && AppExpression.Length > 0)
            {
                Measure = m.Name;
                Dimension = d.Name;

                QSDataList[] dl = qa.GetDataList(AppExpression, AppDimension, qa.LastFilters, NumberOfResults, Descending, MeasureThreshold);
                if (d != null)
                    qa.LastDimension = d;
                if (m != null)
                    qa.LastMeasure = m;

                if (dl.Count() > 0)
                {
                    msg += InterpretDataRanking(qa, dl, Measure, Dimension, AppExpression, AppDimension, Descending, MeasureThreshold);
                }
            }
            else
            {
                msg = StringResources.nlGetChartNotFound;
            }

            return msg;
        }


        private static string InterpretData(QSApp qa, QSDataList[] gl, string Measure, string Dimension, string AppExpression, string AppDimension,
            bool Descending = true)
        {
            string msg = "";

            string Total = qa.GetExpressionFormattedValue(Measure: AppExpression, Label: Measure);

            if (Descending)
            {
                msg += String.Format(StringResources.nlgInterpretDataDescending, Measure, Total, Dimension, gl.First().DimValue, gl.First().MeasFormattedValue);
            }
            else
            {
                msg += String.Format(StringResources.nlgInterpretDataAscending, Measure, Total, Dimension, gl.First().DimValue, gl.First().MeasFormattedValue);
            }

            if (gl.Count() > 1)
            {
                if (!double.IsNaN(gl[1].MeasValue))
                {
                    if (Descending)
                    {
                        msg += " " + String.Format(StringResources.nlgInterpretDataNextDescending,
                            gl[1].DimValue, gl[1].MeasFormattedValue, (1 - gl[1].MeasValue / gl[0].MeasValue).ToString("P1"));
                    }
                    else
                    {
                        msg += " " + String.Format(StringResources.nlgInterpretDataNextAscending,
                            gl[1].DimValue, gl[1].MeasFormattedValue, (1 - gl[1].MeasValue / gl[0].MeasValue).ToString("P1"));
                    }
                }
            }

            return msg;
        }

        private static string InterpretDataRanking(QSApp qa, QSDataList[] gl, string Measure, string Dimension, string AppExpression,
            string AppDimension, bool Descending = true, double? MeasureThreshold = null)
        {
            string msg = "";

            if (Descending)
            {
                msg += String.Format(StringResources.nlgInterpretDataRankingDescending, Dimension, Measure);
            }
            else
            {
                msg += String.Format(StringResources.nlgInterpretDataRankingAscending, Dimension, Measure);
            }

            msg += "\r\n" + GetTextChart(gl, NumberOfResults);


            string strFilter = "";

            if (MeasureThreshold != null)
            {
                double t = MeasureThreshold.Value;

                if (Descending)
                    strFilter += String.Format(StringResources.nlGetElementAboveValue, Measure, t.ToString("N2"));
                else
                    strFilter += String.Format(StringResources.nlGetElementBelowValue, Measure, t.ToString("N2"));
            }


            foreach (QSFoundFilter ff in qa.LastFilters)
            {
                if (ff == qa.LastFilters.First())
                {
                    if (strFilter.Length > 0) strFilter += "\r\n";
                }
                else
                    strFilter += ", " + StringResources.nlAnd + " ";

                strFilter += String.Format(StringResources.nlGetElementMoreFilter, ff.Dimension, ff.Element);
            }
            if (strFilter != "")
                msg += "\r\n\r\n(" + strFilter + ")";

            return msg;
        }


        private static string GetFilteredMeasure(QSApp qa, ref string Measure, List<QSFoundFilter> Filters = null, string UserName = "")
        {
            string msg = "";

            string AppExpression = "";
            QSMasterItem m = qa.GetMasterMeasure(Measure);

            if (Measure == null || m == null)
            {
                msg = string.Format(StringResources.kpiUnknownMessage, Measure);
            }
            else if (Filters != null && Filters.Count() > 0)
            {
                AppExpression = m.Expression;
                Measure = m.Name;

                if (AppExpression.Length > 0)
                {
                    string val = qa.GetExpressionFormattedValue(AppExpression, Filters, Measure);
                    string strFilter = "";

                    foreach (QSFoundFilter ff in Filters)
                    {
                        if (ff != Filters.First()) strFilter += ", " + StringResources.nlAnd + " ";
                        strFilter += String.Format(StringResources.nlGetElementMoreFilter, ff.Dimension, ff.Element);
                    }
                    msg = string.Format(StringResources.nlGetElement, Measure, strFilter, Measure, val);
                }
                else
                {
                    msg = StringResources.nlGetElementNotFound;
                }
            }
            else
            {
                msg = string.Format(StringResources.kpiValueMessage, m.Name, m.FormattedExpression);
            }

            return msg;
        }

        private static string GetFilteredList(QSApp qa, ref string Dimension, List<QSFoundFilter> Filters = null, string UserName = "")
        {
            string msg = "";

            string AppDimension;
            QSMasterItem d = qa.GetMasterDimension(Dimension);
            if (d == null)
                AppDimension = qa.GetDimensionExpression(Dimension);
            else
            {
                AppDimension = d.Expression;
            }

            string AppExpression = "=1";

            if (AppDimension.Length > 0)
            {
                Dimension = d.Name;

                QSDataList[] dl = qa.GetDataList(AppExpression, AppDimension, qa.LastFilters, NumberOfResults * 2);
                if (d != null)
                    qa.LastDimension = d;

                if (dl.Count() > 0)
                {
                    msg += CreateListOfElements(qa, dl, Dimension);
                }
            }
            else
            {
                msg = StringResources.nlGetChartNotFound;
            }

            return msg;
        }

        private static string CreateListOfElements(QSApp qa, QSDataList[] gl, string Dimension)
        {
            string msg = "";
            msg += string.Format(StringResources.nlgShowListOfElements, Dimension, NumberOfResults * 2);

            msg += "\r\n" + GetTextList(gl, NumberOfResults * 2);

            string strFilter = "";
            foreach (QSFoundFilter ff in qa.LastFilters)
            {
                if (ff != qa.LastFilters.First()) strFilter += ", " + StringResources.nlAnd + " ";
                strFilter += String.Format(StringResources.nlGetElementMoreFilter, ff.Dimension, ff.Element);
            }
            if (strFilter != "")
                msg += "\r\n\r\n(" + strFilter + ")";

            return msg;
        }

        private static string GetTextList(QSDataList[] Data, int NumBuckets = 5)
        {
            string TextList = "";
            int i = 0;

            int NofElements = Data.Length;

            foreach (QSDataList d in Data)
            {
                i += 1;
                if (i > NumBuckets) break;

                if (i > 1) TextList += "\r\n";
                TextList += d.DimValue;
            }

            return TextList;
        }

        private static string ApplyFilter(QSApp qa, string Filter, string FilterValue)
        {
            string msg = "";

            if (Filter != "" && FilterValue != "")
            {
                qa.AddFilter(Filter, FilterValue);
                msg = string.Format(StringResources.nlFilterApplied, Filter, FilterValue);

                string strFilter = "";

                foreach (QSFoundFilter ff in qa.LastFilters)
                {
                    if (ff != qa.LastFilters.First()) strFilter += ", " + StringResources.nlAnd + " ";
                    strFilter += String.Format(StringResources.nlGetElementMoreFilter, ff.Dimension, ff.Element);
                }
                msg += "\n" + string.Format(StringResources.nlTotalFilters, strFilter);
            }
            else
            {
                msg = StringResources.nlNoFilterToApply;
            }

            return msg;
        }

    }
}

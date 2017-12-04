using System;
using System.Threading.Tasks;
using System.Net.Http;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QlikNLP
{
    class MicrosoftLuisNLP
    {
        string _LuisBaseUrl;
        string _appId;
        string _subscriptionKey;

        public MicrosoftLuisNLP(string LuisURL, string AppID, string LuisKey)
        {
            _LuisBaseUrl = LuisURL;
            _appId = AppID;
            _subscriptionKey = LuisKey;
        }

        private PredictedIntent Predicted = new PredictedIntent();

        public bool HasPrediction { get { return Predicted.HasPrediction; } }
        public string OriginalQuery { get { return Predicted.OriginalQuery; } }
        public IntentType Intent { get { return Predicted.Intent; } }
        public string Measure { get { return Predicted.Measure; } }
        public string Measure2 { get { return Predicted.Measure2; } }
        public string Element { get { return Predicted.Element; } }
        public string Dimension { get { return Predicted.Dimension; } }
        public string Dimension2 { get { return Predicted.Dimension2; } }
        public string Percentage { get { return Predicted.Percentage; } }
        public double? Number { get { return Predicted.Number; } }
        public string ChartType { get { return Predicted.ChartType; } }
        public double DistanceKm { get { return Predicted.DistanceKm; } }
        public string Language { get { return Predicted.Language; } }
        public string Response { get { return Predicted.Response; } }

        public async Task<bool> Predict(string TextToPredict)
        {
            Predicted.HasPrediction = false;
            Predicted.OriginalQuery = null;
            Predicted.Intent = IntentType.None;
            Predicted.Response = null;

            Predicted.Measure = null;
            Predicted.Measure2 = null;
            Predicted.Element = null;
            Predicted.Dimension = null;
            Predicted.Dimension2 = null;
            Predicted.Percentage = null;
            Predicted.Number = 0;
            Predicted.ChartType = null;
            Predicted.DistanceKm = 0;
            Predicted.Language = null;
            string DistanceUnit = "km";

            try
            {
                LUISAnalyticAssistant LuisResult = await QueryLuis(TextToPredict);

                foreach (Entity et in LuisResult.entities)
                {
                    if (et.type == "Measure") this.Predicted.Measure = et.entity;
                    if (et.type == "Measure2") this.Predicted.Measure2 = et.entity;
                    if (et.type == "Element") this.Predicted.Element = et.entity;
                    if (et.type == "Dimension") this.Predicted.Dimension = et.entity;
                    if (et.type == "Dimension1") this.Predicted.Dimension2 = et.entity;
                    if (et.type == "builtin.percentage" || et.type == "percentage")
                    {
                        this.Predicted.Percentage = et.entity;
                        if (this.Predicted.Number == 0) this.Predicted.Number = PercentageToDouble(et.entity);
                    }
                    if (et.type == "builtin.number") this.Predicted.Number = Convert.ToDouble(et.entity);
                    if (et.type == "ChartType") this.Predicted.ChartType = et.entity;
                    if (et.type == "Distance") this.Predicted.DistanceKm = Convert.ToDouble(et.entity);
                    if (et.type == "DistanceUnit") DistanceUnit = et.entity;
                    if (et.type == "Language") this.Predicted.Language = et.entity;

                    if (DistanceKm > 0 && DistanceUnit != "km")
                    {
                        switch (DistanceUnit)
                        {
                            case "mi": case "miles": Predicted.DistanceKm *= 1.61; break;
                            case "m": case "meters": Predicted.DistanceKm *= 0.001; break;
                        }
                    }
                }


                if (LuisResult.topScoringIntent.intent == "None")
                {
                    this.Predicted.HasPrediction = false;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.None;
                    this.Predicted.Number = 0;
                    this.Predicted.Response = "I do not understand you, could you try again with other question?";
                }
                else if (LuisResult.topScoringIntent.intent == "Show a measure")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.KPI;
                    this.Predicted.Response = "You want to know the value of " + this.Predicted.Measure;
                }
                else if (LuisResult.topScoringIntent.intent == "Show a chart")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Chart;
                    this.Predicted.Response = "You want to know the value of " + this.Predicted.Measure;
                }
                else if (LuisResult.topScoringIntent.intent == "Show a measure for element")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Measure4Element;
                    this.Predicted.Response = "You want to know the value of " + this.Predicted.Measure;
                }
                else if (LuisResult.topScoringIntent.intent == "Filter")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Filter;
                    this.Predicted.Response = "You want to filter " + Predicted.Dimension + " by " + Predicted.Element;
                }
                else if (LuisResult.topScoringIntent.intent == "Alert")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Alert;
                    this.Predicted.Response = "You want to be alerted when " + this.Predicted.Measure + " changes by " + this.Predicted.Percentage;
                }
                else if (LuisResult.topScoringIntent.intent == "Good Answer")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.GoodAnswer;
                    this.Predicted.Response = ":-)";
                }
                else if (LuisResult.topScoringIntent.intent == "Ranking Top")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.RankingTop;
                    this.Predicted.Response = "You want to know the top elements by " + this.Predicted.Dimension;
                }
                else if (LuisResult.topScoringIntent.intent == "Ranking Bottom")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.RankingBottom;
                    this.Predicted.Response = "You want to know the bottom elements by " + this.Predicted.Dimension;
                }
                else if (LuisResult.topScoringIntent.intent == "Hello")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Hello;
                    this.Predicted.Response = "Hello";
                }
                else if (LuisResult.topScoringIntent.intent == "Bye")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Bye;
                    this.Predicted.Response = "Bye";
                }
                else if (LuisResult.topScoringIntent.intent == "BadWorks")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.BadWords;
                    this.Predicted.Response = ":-(\nI prefer not to answer this type of questions, I am a robot but I am very polite.";
                }
                else if (LuisResult.topScoringIntent.intent == "Reports")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Reports;
                    this.Predicted.Response = "I will show you the available reports";
                }
                else if (LuisResult.topScoringIntent.intent == "CreateChart")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.CreateChart;
                    this.Predicted.Response = "I will create a chart for you";
                }
                else if (LuisResult.topScoringIntent.intent == "ShowAnalysis")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ShowAnalysis;
                    this.Predicted.Response = "I will create an analysis for you";
                }
                else if (LuisResult.topScoringIntent.intent == "GeoFilter")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.GeoFilter;
                    this.Predicted.Response = "I will filter the information based on your location";
                }
                else if (LuisResult.topScoringIntent.intent == "Apologize")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Apologize;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ClearAllFilters")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ClearAllFilters;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ClearDimensionFilter")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ClearDimensionFilter;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "CreateChart")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.CreateChart;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "CreateCollaborationGroup")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.CreateCollaborationGroup;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "CurrentSelections")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.CurrentSelections;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "Help")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Help;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "Reports")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.Reports;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ShowAllApps")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ShowAllApps;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ShowAllDimensions")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ShowAllDimensions;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ShowAllMeasures")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ShowAllMeasures;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ShowAllSheets")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ShowAllSheets;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ShowAllStories")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ShowAllStories;
                    this.Predicted.Response = "OK";
                }
                else if (LuisResult.topScoringIntent.intent == "ShowAnalysis")
                {
                    this.Predicted.HasPrediction = true;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.ShowAnalysis;
                    this.Predicted.Response = "OK";
                }

                else
                {
                    this.Predicted.HasPrediction = false;
                    this.Predicted.OriginalQuery = TextToPredict;
                    this.Predicted.Intent = IntentType.None;
                    this.Predicted.Response = "I do not have the logic to answer " + TextToPredict;
                }

                return this.HasPrediction;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        private double PercentageToDouble(string Percentage)
        {
            double d = 0;
            string strPct = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.PercentSymbol;

            if (!Percentage.Contains(strPct)) return 0;

            try
            {
                d = double.Parse(Percentage.Replace(strPct, "")) / 100;
            }
            catch (Exception e)
            {
                d = 0;
            }

            return d;
        }


        private async Task<LUISAnalyticAssistant> QueryLuis(string message)
        {
            using (var client = new HttpClient())
            {
                string LuisUrl = _LuisBaseUrl + _appId + "?subscription-key=" + _subscriptionKey + "&q=" + System.Uri.EscapeDataString(message);
                string LuisJson = await client.GetStringAsync(LuisUrl);

                LUISAnalyticAssistant LuisResult = new LUISAnalyticAssistant();
                LuisResult = Newtonsoft.Json.JsonConvert.DeserializeObject<LUISAnalyticAssistant>(LuisJson);

                return LuisResult;
            }
        }

    }

        public class Resolution
    {
    }

    public class Value
    {
        public string entity { get; set; }
        public string type { get; set; }
        public Resolution resolution { get; set; }
    }

    public class Parameter
    {
        public string name { get; set; }
        public string type { get; set; }
        public bool required { get; set; }
        public IList<Value> value { get; set; }
    }

    public class Action
    {
        public bool triggered { get; set; }
        public string name { get; set; }
        public IList<Parameter> parameters { get; set; }
    }

    public class Intent
    {
        public string intent { get; set; }
        public double score { get; set; }
        public IList<Action> actions { get; set; }
    }

    public class TopScoringIntent
    {
        public string intent { get; set; }
        public double score { get; set; }
        public IList<Action> actions { get; set; }
    }

    public class Entity
    {
        public string entity { get; set; }
        public string type { get; set; }
        public int startIndex { get; set; }
        public int endIndex { get; set; }
        public double score { get; set; }
        public object resolution { get; set; }
}

public class Dialog
{
    public string contextId { get; set; }
    public string status { get; set; }
}

public class LUISAnalyticAssistant
{
    public string query { get; set; }
    public TopScoringIntent topScoringIntent { get; set; }
    public IList<Intent> intents { get; set; }
    public IList<Entity> entities { get; set; }
    public Dialog dialog { get; set; }
}



//public class Entity
//    {
//        public string entity { get; set; }
//        public string type { get; set; }
//        public int startIndex { get; set; }
//        public int endIndex { get; set; }
//        public double score { get; set; }
//    }

//    public class LUISAnalyticAssistant
//    {
//        public string query { get; set; }
//        public TopScoringIntent topScoringIntent { get; set; }
//        public IList<Entity> entities { get; set; }
//    }

    public class Query
    {
        public string Measure { get; set; }
        public string Element { get; set; }
        public string Dimension { get; set; }
    }

}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net.Http;
using Newtonsoft.Json.Linq;


namespace QlikNLP
{
    public enum IntentType
    {
        None = 0, KPI = 1, Chart = 2, Measure4Element = 3, Alert = 4, GoodAnswer = 5
        , RankingTop = 6, RankingBottom = 7, Hello = 8, Bye = 9, Filter = 10, Reports = 11, BadWords = 12, CreateChart = 13
        , ShowAnalysis = 14, GeoFilter = 15
        , ContactQlik = 16, Help = 17, PersonalInformation = 18, ClearAllFilters = 19, ClearDimensionFilter = 20
        , CurrentSelections = 21, ShowAllApps = 22, ShowAllDimensions = 23, ShowAllMeasures = 24, ShowAllSheets = 25
        , ShowAllStories = 26, ShowKPIs = 27, ShowMeasureByMeasure = 28, ShowListOfElements = 29, ChangeLanguage = 30
        , Apologize = 31, DirectResponse = 32, ShowElementsAboveValue = 33, ShowElementsBelowValue = 34, CreateCollaborationGroup = 35
    };
    public enum NLPEngine { None = 0, MicrosoftLuis = 1, GoogleApiAi = 2, AmazonAlexa = 3 };

    public class PredictedIntent
    {
        public bool HasPrediction { get; set; }
        public string OriginalQuery { get; set; }
        public IntentType Intent { get; set; }
        public string Measure { get; set; }
        public string Measure2 { get; set; }
        public string Element { get; set; }
        public string Dimension { get; set; }
        public string Dimension2 { get; set; }
        public string Percentage { get; set; }
        public double? Number { get; set; }
        public string ChartType { get; set; }
        public double DistanceKm { get; set; }
        public string Language { get; set; }
        public string Response { get; set; }
    }

    public class NLP
    {
        GoogleNLP QlikApiAi;
        MicrosoftLuisNLP QlikLuis;
        NLPEngine UsedEngine = NLPEngine.None;

        //public void NLPStartLUIS(string LuisURL, string AppID, string LuisKey)
        //{
        //    QlikLuis = new NLPMicrosoftLUIS(LuisURL, AppID, LuisKey);
        //    UsedEngine = NLPEngine.MicrosoftLuis;
        //}
        public void NLPStartLUIS(string LuisURL, string AppID, string LuisKey)
        {
            QlikLuis = new MicrosoftLuisNLP(LuisURL, AppID, LuisKey);
            UsedEngine = NLPEngine.MicrosoftLuis;
        }
        public void NLPStartApiAi(string AppID, string Language = "English")
        {
            QlikApiAi = new GoogleNLP(AppID, Language);
            UsedEngine = NLPEngine.GoogleApiAi;
        }

        private PredictedIntent Predicted = new PredictedIntent();

        public bool HasPrediction { get { return Predicted.HasPrediction; } }
        public string OriginalQuery { get { return Predicted.OriginalQuery; } }
        public IntentType Intent { get { return Predicted.Intent; } set { Predicted.Intent = value; } }
        public string Measure { get { return Predicted.Measure; } set { Predicted.Measure = value; } }
        public string Measure2 { get { return Predicted.Measure2; } set { Predicted.Measure2 = value; } }
        public string Element { get { return Predicted.Element; } set { Predicted.Element = value; } }
        public string Dimension { get { return Predicted.Dimension; } set { Predicted.Dimension = value; } }
        public string Dimension2 { get { return Predicted.Dimension2; } set { Predicted.Dimension2 = value; } }
        public string Percentage { get { return Predicted.Percentage; } set { Predicted.Percentage = value; } }
        public double? Number { get { return Predicted.Number; } set { Predicted.Number = value; } }
        public string ChartType { get { return Predicted.ChartType; } set { Predicted.ChartType = value; } }
        public double DistanceKm { get { return Predicted.DistanceKm; } set { Predicted.DistanceKm = value; } }
        public string Language { get { return Predicted.Language; } set { Predicted.Language = value; } }
        public string Response { get { return Predicted.Response; } set { Predicted.Response = value; } }

        public PredictedIntent GetPredictedIntent()
        {
            return Predicted;
        }

        public PredictedIntent GetCopyOfPredictedIntent()
        {
            NLP n = new NLP();
            n.Predicted.HasPrediction = this.HasPrediction;
            n.Predicted.OriginalQuery = this.OriginalQuery;
            n.Predicted.Intent = this.Intent;
            n.Predicted.Measure = this.Measure;
            n.Predicted.Measure2 = this.Measure2;
            n.Predicted.Element = this.Element;
            n.Predicted.Dimension = this.Dimension;
            n.Predicted.Dimension2 = this.Dimension2;
            n.Predicted.Percentage = this.Percentage;
            n.Predicted.Number = this.Number;
            n.Predicted.ChartType = this.ChartType;
            n.Predicted.DistanceKm = this.DistanceKm;
            n.Predicted.Language = this.Language;
            n.Predicted.Response = this.Response;

            return n.Predicted;
        }

        public async Task<bool> Predict(string TextToPredict)
        {
            bool result = false;

            Predicted.Measure = null;
            Predicted.Measure2 = null;
            Predicted.Element = null;
            Predicted.Dimension = null;
            Predicted.Dimension2 = null;
            Predicted.Percentage = null;
            Predicted.ChartType = null;
            Predicted.DistanceKm = 0;
            Predicted.Language = null;

            //if (UsedEngine == NLPEngine.MicrosoftLuis)
            //{
            //    result = await QlikLuis.Predict(TextToPredict);

            //    Predicted.HasPrediction = QlikLuis.HasPrediction;
            //    Predicted.OriginalQuery = QlikLuis.OriginalQuery;
            //    Predicted.Intent = QlikLuis.Intent;
            //    Predicted.Measure = QlikLuis.Measure;
            //    Predicted.Measure2 = QlikLuis.Measure2;
            //    Predicted.Element = QlikLuis.Element;
            //    Predicted.Dimension = QlikLuis.Dimension;
            //    Predicted.Dimension2 = QlikLuis.Dimension2;
            //    Predicted.Percentage = QlikLuis.Percentage;
            //    Predicted.Number = QlikLuis.Number;
            //    Predicted.ChartType = QlikLuis.ChartType;
            //    Predicted.DistanceKm = QlikLuis.DistanceKm;
            //    Predicted.Language = QlikLuis.Language;
            //    Predicted.Response = QlikLuis.Response;
            //}
            if (UsedEngine == NLPEngine.GoogleApiAi)
            {
                result = QlikApiAi.Predict(TextToPredict);

                Predicted.HasPrediction = QlikApiAi.HasPrediction;
                Predicted.OriginalQuery = QlikApiAi.OriginalQuery;
                Predicted.Intent = QlikApiAi.Intent;
                Predicted.Measure = QlikApiAi.Measure;
                Predicted.Measure2 = QlikApiAi.Measure2;
                Predicted.Element = QlikApiAi.Element;
                Predicted.Dimension = QlikApiAi.Dimension;
                Predicted.Dimension2 = QlikApiAi.Dimension2;
                Predicted.Percentage = QlikApiAi.Percentage;
                Predicted.Number = QlikApiAi.Number;
                Predicted.ChartType = QlikApiAi.ChartType;
                Predicted.DistanceKm = QlikApiAi.DistanceKm;
                Predicted.Language = QlikApiAi.Language;
                Predicted.Response = QlikApiAi.Response;
            }

            return result;
        }


    }
}

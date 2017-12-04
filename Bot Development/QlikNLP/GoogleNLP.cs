using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ApiAiSDK;
using ApiAiSDK.Model;


namespace QlikNLP
{
    public class GoogleNLP
    {
        private ApiAi apiAI;
        public bool IsConnected = false;

        public GoogleNLP(string ClientToken, string Language = "Spanish")
        {
            SupportedLanguage supLang = ApiAiSDK.SupportedLanguage.FromLanguageTag(Language);
            var config = new AIConfiguration(ClientToken, supLang);
            apiAI = new ApiAi(config);
            IsConnected = true;
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

        public bool Predict(string TextToPredict)
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
                AIResponse response = apiAI.TextRequest(TextToPredict);

                if (!response.IsError && response.Status.ErrorType == "success")
                {
                    foreach (System.Collections.Generic.KeyValuePair<string, object> param in response.Result.Parameters)
                    {
                        string strParam = (string)param.Value.ToString();

                        if (strParam.StartsWith("[\r\n"))
                            strParam = (string)param.Value.ToString().Replace("[\r\n  \"", "").Replace("\"\r\n]", "");

                        if (param.Key == "Measure" && strParam != "") Predicted.Measure = strParam;
                        if (param.Key == "Measure2" && strParam != "") Predicted.Measure2 = strParam;
                        if (param.Key == "Element" && strParam != "") Predicted.Element = strParam;
                        if (param.Key == "Dimension" && strParam != "") Predicted.Dimension = strParam;
                        if (param.Key == "Dimension1" && strParam != "") Predicted.Dimension2 = strParam;
                        if (param.Key == "Percentage" && strParam != "")
                        {
                            Predicted.Percentage = strParam;
                            Predicted.Number = PercentageToDouble(Predicted.Percentage);
                        }
                        if (param.Key == "Number" && strParam != "") Predicted.Number = Convert.ToDouble(param.Value);
                        if (param.Key == "ChartType" && strParam != "") Predicted.ChartType = strParam;
                        if (param.Key == "Distance" && strParam != "") Predicted.DistanceKm = Convert.ToDouble(param.Value);
                        if (param.Key == "Language" && strParam != "") Predicted.Language = strParam;
                        if (param.Key == "DistanceUnit" && strParam != "") DistanceUnit = strParam.ToLower();
                    }
                    if (DistanceKm > 0 && DistanceUnit != "km")
                    {
                        switch (DistanceUnit)
                        {
                            case "mi": case "miles": Predicted.DistanceKm *= 1.61; break;
                            case "m": case "meters": Predicted.DistanceKm *= 0.001; break;
                        }
                    }


                    if (response.Result.Action == "input.unknown")
                    {
                        this.Predicted.HasPrediction = false;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.None;
                        this.Predicted.Number = 0;
                        this.Predicted.Response = "I do not understand you, could you try again with other question?";
                    }
                    else if (response.Result.Source == "domains")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.DirectResponse;
                        this.Predicted.Response = response.Result.Fulfillment.Speech;
                    }
                    else if (response.Result.Action == "ShowMeasure")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.KPI;
                        this.Predicted.Response = "You want to know the value of " + Predicted.Measure;
                    }
                    else if (response.Result.Action == "ShowChart")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Chart;
                        this.Predicted.Response = "You want to know the value of " + Predicted.Measure;
                    }
                    else if (response.Result.Action == "ShowMeasureForElement")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Measure4Element;
                        this.Predicted.Response = "You want to know the value of " + Predicted.Measure;
                    }
                    else if (response.Result.Action == "Alert")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Alert;
                        this.Predicted.Response = "You want to be alerted when " + Predicted.Measure + " changes by " + Predicted.Percentage;
                    }
                    else if (response.Result.Action == "GoodAnswer")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.GoodAnswer;
                        this.Predicted.Response = ":-)";
                    }
                    else if (response.Result.Action == "RankingTop")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.RankingTop;
                        this.Predicted.Response = "You want to know the top elements by " + Predicted.Dimension;
                    }
                    else if (response.Result.Action == "RankingBottom")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.RankingBottom;
                        this.Predicted.Response = "You want to know the bottom elements by " + Predicted.Dimension;
                    }
                    else if (response.Result.Action == "Filter")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Filter;
                        this.Predicted.Response = "You want to filter " + Predicted.Dimension + " by " + Predicted.Element;
                    }
                    else if (response.Result.Action == "Hello" || response.Result.Action == "input.welcome")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Hello;
                        this.Predicted.Response = "Hello";
                    }
                    else if (response.Result.Action == "Bye")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Bye;
                        this.Predicted.Response = "Bye";
                    }
                    else if (response.Result.Action == "BadWords")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.BadWords;
                        this.Predicted.Response = ":-(\nI prefer not to answer this type of questions, I am a robot but I am very polite.";
                    }
                    else if (response.Result.Action == "Reports")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Reports;
                        this.Predicted.Response = "I will show you the available reports";
                    }
                    else if (response.Result.Action == "CreateChart")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.CreateChart;
                        this.Predicted.Response = "I will create a chart for you";
                    }
                    else if (response.Result.Action == "ShowAnalysis")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowAnalysis;
                        this.Predicted.Response = "I will create an analysis for you";
                    }
                    else if (response.Result.Action == "GeoFilter")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.GeoFilter;
                        this.Predicted.Response = "I will filter the information based on your location";
                    }
                    else if (response.Result.Action == "ContactQlik")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ContactQlik;
                        this.Predicted.Response = "You can go to www.qlik.com";
                    }
                    else if (response.Result.Action == "Help")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Help;
                        this.Predicted.Response = "You can ask for any information in the current Qlik Sense app";
                    }
                    else if (response.Result.Action == "PersonalInformation")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.PersonalInformation;
                        this.Predicted.Response = "I think you better go to https://www.facebook.com/profile.php?id=100016396000544 and know everything about me";    // StringResources
                    }
                    else if (response.Result.Action == "ClearAllFilters")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ClearAllFilters;
                        this.Predicted.Response = "I will clear all filters";
                    }
                    else if (response.Result.Action == "ClearDimensionFilter")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ClearDimensionFilter;
                        this.Predicted.Response = "I will clear this dimension filter";
                    }
                    else if (response.Result.Action == "CurrentSelections")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.CurrentSelections;
                        this.Predicted.Response = "I will show you the current selections in the app";
                    }
                    else if (response.Result.Action == "ShowAllApps")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowAllApps;
                        this.Predicted.Response = "I will show all the available apps you can connect";
                    }
                    else if (response.Result.Action == "ShowAllDimensions")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowAllDimensions;
                        this.Predicted.Response = "I will show you all the master dimensions in the current app";
                    }
                    else if (response.Result.Action == "ShowAllMeasures")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowAllMeasures;
                        this.Predicted.Response = "I will show you all the master measures in the current app";
                    }
                    else if (response.Result.Action == "ShowAllSheets")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowAllSheets;
                        this.Predicted.Response = "I will show you all the sheets in the current app";
                    }
                    else if (response.Result.Action == "ShowAllStories")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowAllStories;
                        this.Predicted.Response = "I will show you all the stories in the current app";
                    }
                    else if (response.Result.Action == "ShowKPIs")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowKPIs;
                        this.Predicted.Response = "I will show you the most used metrics in the current app";
                    }
                    else if (response.Result.Action == "ShowMeasureByMeasure")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowMeasureByMeasure;
                        this.Predicted.Response = "I will show you the result of analyzing these two measures";
                    }
                    else if (response.Result.Action == "ShowListOfElements")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowListOfElements;
                        this.Predicted.Response = "I will show you a list of elements for this dimension";
                    }
                    else if (response.Result.Action == "ChangeLanguage")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ChangeLanguage;
                        this.Predicted.Response = "I will change the language";
                    }
                    else if (response.Result.Action == "Apologize")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.Apologize;
                        this.Predicted.Response = "OK";
                    }
                    else if (response.Result.Action == "ShowElementsAboveValue")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowElementsAboveValue;
                        this.Predicted.Response = "You want to know the elements that meet a measure is above this value";
                    }
                    else if (response.Result.Action == "ShowElementsBelowValue")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.ShowElementsBelowValue;
                        this.Predicted.Response = "You want to know the elements that meet a measure is below this value";
                    }
                    else if (response.Result.Action == "CreateCollaborationGroup")
                    {
                        this.Predicted.HasPrediction = true;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.CreateCollaborationGroup;
                        this.Predicted.Response = "You want to create a group to collaborate with other people";
                    }
                    else
                    {
                        this.Predicted.HasPrediction = false;
                        this.Predicted.OriginalQuery = TextToPredict;
                        this.Predicted.Intent = IntentType.None;
                        this.Predicted.Response = "I do not have the logic to answer " + TextToPredict + " yet";
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Predicted.HasPrediction = false;
            }

            return Predicted.HasPrediction;
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
    }
}

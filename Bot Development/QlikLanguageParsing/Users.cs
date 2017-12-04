using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Globalization;

using QlikSenseEasy;
using QlikNLP;


namespace QlikConversationService
{
    public class QSUsers
    {
        private static List<QSUser> Users;

        public QSUsers()
        {
            Users = new List<QSUser>();
        }

        public List<QSUser> GetUserList()
        {
            return (Users);
        }

        public void ReadFromCSV(string FileName)
        {
            List<QSUser> values;
            try
            {
                values = File.ReadAllLines(FileName, System.Text.Encoding.Default)
                                               .Skip(0)
                                               .Select(v => QSUser.FromCsv(v))
                                               .ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error in QSUsers with file \"{0}\": {1} Exception caught.", FileName, e);
                values = null;
            }

            if (values != null) Users = values;
        }

        public void WriteToCSV(string FileName)
        {
            try
            {
                File.WriteAllLines(FileName, Users.Select(u => QSUser.ToCsv(u)), System.Text.Encoding.Default);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error in QSUsers writing to file \"{0}\": {1} Exception caught.", FileName, e);
            }
        }

        public QSUser GetUser(string UserId)
        {
            QSUser result;
            try
            {
                result = Users.Find(x => x.UserId.ToLower() == UserId.ToLower().Trim());
            }
            catch (Exception e)
            {
                result = null;
            }

            if (result != null) return result;
            else return null;
        }

        public QSUser AddUser(string UserId, string UserName = "", string QSUserId = "", string QSUserDir = "", string QSUserName = ""
            , string LastAccess = "", bool Allowed = true)
        {
            QSUser Usr;

            try
            {
                Usr = GetUser(UserId);

                if (Usr == null)
                {
                    Usr = new QSUser();
                    Usr.UserId = UserId.Trim();
                    Usr.Allowed = Allowed;
                    Users.Add(Usr);
                }

                if (Usr.UserName == "" || UserName != "") Usr.UserName = UserName.Trim();
                if (Usr.QSUserId == "" || QSUserId != "") Usr.QSUserId = QSUserId.Trim();
                if (Usr.QSUserDir == "" || QSUserDir != "") Usr.QSUserDir = QSUserDir.Trim();
                if (Usr.QSUserName == "" || QSUserName != "") Usr.QSUserName = QSUserName.Trim();

                DateTime LastAccessDateTime;
                if (LastAccess == "")
                {
                    LastAccessDateTime = DateTime.Now;
                }
                else
                {
                    LastAccessDateTime = DateTime.ParseExact(LastAccess, QSUser.cntDateTimeParseFormat, CultureInfo.InvariantCulture);
                }
                Usr.LastAccess = LastAccessDateTime;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error in AddUser with UserId \"{0}\": {1} Exception caught.", UserId, e);
                return null;
            }

            return Usr;
        }
    }

    public class QSConversationSummaryItem
    {
        public IntentType Intent;
        public string Measure;
        public string Dimension;
        public string Element;
        public int Times = 0;
    }

    public class QSConversationSummary
    {
        public List<QSConversationSummaryItem> Measure = new List<QSConversationSummaryItem>();
        public List<QSConversationSummaryItem> MeasureByDimension = new List<QSConversationSummaryItem>();
        public List<QSConversationSummaryItem> MeasureForElement = new List<QSConversationSummaryItem>();
        public List<QSConversationSummaryItem> Top = new List<QSConversationSummaryItem>();
        public List<QSConversationSummaryItem> Bottom = new List<QSConversationSummaryItem>();
        public List<QSConversationSummaryItem> Filter = new List<QSConversationSummaryItem>();
    }

    public class QSUser
    {
        public string UserId;
        public string UserName;
        public string QSUserId;
        public string QSUserDir;
        public string QSUserName;
        private DateTime _LastAccess;
        private DateTime _PreviousAccess;
        public bool Allowed;
        public string Language = "en-US";
        public QSApp QS;
        public Response LastResponse;
        public QSFoundObject LastChart;
        public List<PredictedIntent> ConversationHistory = new List<PredictedIntent>();
        public QSConversationSummary ConversationSummary = new QSConversationSummary();

        public DateTime LastAccess
        {
            get { return _LastAccess; }
            set
            {
                _PreviousAccess = LastAccess;
                _LastAccess = value;
            }
        }



        private bool _TheBotIsAngry = false;
        public bool TheBotIsAngry
        {
            get { return _TheBotIsAngry; }
            set
            {
                _TheBotIsAngry = value;
                if (_TheBotIsAngry)
                    _LastArgument = DateTime.Now;
            }
        }

        private DateTime _LastArgument;


        public TimeSpan TimeSincePreviousAccess()
        {
            return LastAccess.Subtract(_PreviousAccess);
        }

        public TimeSpan TimeSinceLastArgument()
        {
            TimeSpan t = new TimeSpan(0);

            if (TheBotIsAngry)
                t = DateTime.Now.Subtract(_LastArgument);

            return t;
        }



        public const string cntDateTimeParseFormat = "yyyy-MM-dd HH:mm:ss";

        public static QSUser FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(';');
            QSUser U = new QSUser();

            U.UserId = values[0].ToLower().Trim();
            U.UserName = values[1].Trim();
            U.QSUserId = values[2].Trim();
            U.QSUserDir = values[3].Trim();
            U.QSUserName = values[4].Trim();
            U.LastAccess = DateTime.ParseExact(values[5], cntDateTimeParseFormat, CultureInfo.InvariantCulture);
            U.Allowed = values[6].Trim() == "Y" ? true : false;
            if (values.Length > 7) U.Language = values[7].Trim();

            return U;
        }

        public static string ToCsv(QSUser U)
        {
            string csvLine;

            csvLine = string.Format("{0};{1};{2};{3};{4};{5};{6};{7}",
                U.UserId, U.UserName, U.QSUserId, U.QSUserDir, U.QSUserName, U.LastAccess.ToString(cntDateTimeParseFormat), U.Allowed ? "Y" : "N", U.Language);

            return csvLine;
        }

        public void SummarizeHistory()
        {
            foreach (PredictedIntent Pred in ConversationHistory)
            {
                QSConversationSummaryItem NewItem = new QSConversationSummaryItem
                {
                    Intent = Pred.Intent,
                    Measure = Pred.Measure,
                    Dimension = Pred.Dimension,
                    Element = Pred.Element,
                    Times = 1
                };

                QSConversationSummaryItem item;

                if (Pred.Intent == IntentType.KPI || Pred.Intent == IntentType.Chart || Pred.Intent == IntentType.Measure4Element)
                {
                    item = ConversationSummary.Measure.Find(s => s.Measure == Pred.Measure);
                    if (item == null)
                    {
                        item = NewItem;
                        ConversationSummary.Measure.Add(item);
                    }
                    else
                        item.Times++;
                }

                if (Pred.Intent == IntentType.Chart)
                {
                    item = ConversationSummary.MeasureByDimension.Find(s => s.Measure == Pred.Measure && s.Dimension == Pred.Dimension);
                    if (item == null)
                    {
                        item = NewItem;
                        ConversationSummary.MeasureByDimension.Add(item);
                    }
                    else
                        item.Times++;
                }

                if (Pred.Intent == IntentType.Measure4Element)
                {
                    item = ConversationSummary.MeasureForElement.Find(s => s.Measure == Pred.Measure && s.Element == Pred.Element && s.Dimension == Pred.Dimension);
                    if (item == null)
                    {
                        item = NewItem;
                        ConversationSummary.MeasureForElement.Add(item);
                    }
                    else
                        item.Times++;
                }

                if (Pred.Intent == IntentType.Filter || Pred.Intent == IntentType.Measure4Element)
                {
                    item = ConversationSummary.Filter.Find(s => s.Dimension == Pred.Dimension && s.Element == Pred.Element);
                    if (item == null)
                    {
                        item = NewItem;
                        ConversationSummary.Filter.Add(item);
                    }
                    else
                        item.Times++;
                }

                if (Pred.Intent == IntentType.RankingTop)
                {
                    item = ConversationSummary.Top.Find(s => s.Measure == Pred.Measure && s.Dimension == Pred.Dimension);
                    if (item == null)
                    {
                        item = NewItem;
                        ConversationSummary.Top.Add(item);
                    }
                    else
                        item.Times++;
                }

                if (Pred.Intent == IntentType.RankingBottom)
                {
                    item = ConversationSummary.Bottom.Find(s => s.Measure == Pred.Measure && s.Dimension == Pred.Dimension);
                    if (item == null)
                    {
                        item = NewItem;
                        ConversationSummary.Bottom.Add(item);
                    }
                    else
                        item.Times++;
                }

            }

            ConversationHistory.Clear();

            ConversationSummary.Measure.OrderByDescending(order => order.Times);
            ConversationSummary.MeasureByDimension.OrderByDescending(order => order.Times);
            ConversationSummary.MeasureForElement.OrderByDescending(order => order.Times);
            ConversationSummary.Filter.OrderByDescending(order => order.Times);
            ConversationSummary.Top.OrderByDescending(order => order.Times);
            ConversationSummary.Bottom.OrderByDescending(order => order.Times);
        }
    }
}

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Configuration;
using System.Diagnostics;

using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types;
using Telegram.Bot.Types.InlineKeyboardButtons;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.InlineQueryResults;
using Telegram.Bot.Types.InputMessageContents;
using Telegram.Bot.Types.ReplyMarkups;

using QlikSenseEasy;
using QlikConversationService;
using QlikTelegram.Multilanguage;
using QlikLog;

namespace QlikTelegram
{
    public class Bot
    {
        private TelegramBotClient botClient = new TelegramBotClient(ConfigurationManager.AppSettings["cntBotToken"]);

        //talk to userinfobot to find out your user id
        private long cntTelegramAdministratorID = Convert.ToInt64(ConfigurationManager.AppSettings["cntTelegramAdministratorID"]);
        private long cntBotLocalAdministrator;
        private string cntBotName;
        private LogFile botLog;

        private string cntqsAppName = ConfigurationManager.AppSettings["DemoqsAppName"];
        private string cntqsAppId = ConfigurationManager.AppSettings["DemoqsAppId"];
        private string cntqsServer = ConfigurationManager.AppSettings["DemoqsServer"];
        private bool cntqsServerSSL;
        private string cntqsServerVirtualProxy;
        private string DemoqsServerHeaderAuth = ConfigurationManager.AppSettings["DemoqsServerHeaderAuth"];

        private string cntqsSingleServer = ConfigurationManager.AppSettings["DemoqsSingleServer"];
        private string cntqsSingleApp = ConfigurationManager.AppSettings["DemoqsSingleApp"];
        private string cntqsAlternativeStreams = ConfigurationManager.AppSettings["cntAlternativeStreams"];
        private string cntStreamIdPublishNewApps = ConfigurationManager.AppSettings["cntStreamIdPublishNewApps"];

        private string cntQSSheetForAnalysis = ConfigurationManager.AppSettings["cntQSSheetForAnalysis"];
        private string NPrintingImgsPath = ConfigurationManager.AppSettings["NPrintingImgsPath"];
        private string NPrintingDefaultReport = ConfigurationManager.AppSettings["NPrintingDefaultReport"];

        private string cntQlikUsersCSV = ConfigurationManager.AppSettings["cntQlikUsersCSV"];
        private QSUsers QlikUsers = new QSUsers();

        private bool cntAllowNewUsers = true;

        private TelegramConversation Conversation;

        private string LuisURL = ConfigurationManager.AppSettings["cntLuisURL"];
        private string LuisAppID = ConfigurationManager.AppSettings["cntLuisAppID"];
        private string LuisKey = ConfigurationManager.AppSettings["cntLuisKey"];

        private string cntApiAiKey = ConfigurationManager.AppSettings["cntApiAiKey"];
        private string cntApiAiLanguage = ConfigurationManager.AppSettings["cntApiAiLanguage"];

        private string cntBingSearchKey = ConfigurationManager.AppSettings["cntBingSearchKey"];

        private string cntCaptureImageApp = ConfigurationManager.AppSettings["cntCaptureImageApp"];
        private string cntCaptureWeb = ConfigurationManager.AppSettings["cntCaptureWeb"];
        private string cntCaptureTimeout = ConfigurationManager.AppSettings["cntCaptureTimeout"];

        private string cntFolderConnection = ConfigurationManager.AppSettings["cntFolderConnection"];

        private string CurrentLanguage = "es-ES";

        private bool cntCheckSDKVersion = true;
        private bool DemoMode = false;

        private double GasDistance = 5;

        private int WaitingGeoFilter = 0;
        private string WaitingGeoDimension = "";
        private double WaitingGeoDistanceKm = 0;

        //Constructor
        public Bot()
        {
            try
            {
                Console.WriteLine("Thank you for using Qlik telegram chatbot.");
                
                botClient.OnCallbackQuery += BotOnCallbackQueryReceived;
                botClient.OnMessage += BotOnMessageReceived;
                botClient.OnMessageEdited += BotOnMessageReceived;
                botClient.OnInlineQuery += BotOnInlineQueryReceived;
                botClient.OnInlineResultChosen += BotOnChosenInlineResultReceived;
                botClient.OnReceiveError += BotOnReceiveError;
                botClient.Timeout = new TimeSpan(0, 0, 5); 

                var me = botClient.GetMeAsync().Result;


                //Telegram ID for Bot Admin
                if (ConfigurationManager.AppSettings["cntBotLocalAdministrator"] != null && ConfigurationManager.AppSettings["cntBotLocalAdministrator"].Trim().Length > 0)
                    cntBotLocalAdministrator = Convert.ToInt32(ConfigurationManager.AppSettings["cntBotLocalAdministrator"]);
                else
                    cntBotLocalAdministrator = -1;
                

                //alert settup
                System.Timers.Timer AlertTimer = new System.Timers.Timer(Convert.ToInt64(ConfigurationManager.AppSettings["AlertSeconds"]) * 1000);
                AlertTimer.AutoReset = true;
                AlertTimer.Elapsed += CheckAlerts;
                AlertTimer.Start();

                //Log settup
                botLog = new LogFile(ConfigurationManager.AppSettings["logfilepath"]);
                botLog.AddBotLine("Timer for alerts started with " + AlertTimer.Interval / 1E3 + " seconds", me.Id.ToString(), me.FirstName, me.LastName, me.Username);

                //NPrinting settup
                if (NPrintingDefaultReport == null || NPrintingDefaultReport == "") NPrintingDefaultReport = "ReportSales.pdf";

                //language settup
                if (ConfigurationManager.AppSettings["cntDefaultLanguage"] != "") CurrentLanguage = ConfigurationManager.AppSettings["cntDefaultLanguage"];
                ChangeAppLanguage(CurrentLanguage);


                botLog = new LogFile(ConfigurationManager.AppSettings["logfilepath"]);
                botLog.AddLine("Started and Callback set for Bot " + cntBotName);
                botClient.StartReceiving();


                //qs server settup
                QSApp QS = new QSApp();
                QS.qsAppName = cntqsAppName;
                QS.qsAppId = cntqsAppId;
                QS.qsServer = cntqsServer;
                if (ConfigurationManager.AppSettings["DemoqsServerSSL"] != null
                    && ConfigurationManager.AppSettings["DemoqsServerSSL"].ToLower().StartsWith("y"))
                    cntqsServerSSL = true;
                cntqsServerVirtualProxy = ConfigurationManager.AppSettings["DemoqsServerVirtualProxy"];
                if (cntqsServerVirtualProxy == null) cntqsServerVirtualProxy = "";
                QS.qsSingleServer = cntqsSingleServer;
                QS.qsSingleApp = cntqsSingleApp;
                QS.qsAlternativeStreams = cntqsAlternativeStreams;
                if (ConfigurationManager.AppSettings["cntCheckSDKVersion"] != null)
                {
                    if (ConfigurationManager.AppSettings["cntCheckSDKVersion"].ToUpper().StartsWith("Y"))
                    {
                        cntCheckSDKVersion = true;
                    }
                    else
                    {
                        cntCheckSDKVersion = false;
                    }
                }
                try
                {
                    //seriously needs to be fixed
                    QS.QSConnectServerHeader("TestUser", DemoqsServerHeaderAuth, cntqsServerVirtualProxy, cntqsServerSSL, cntCheckSDKVersion);

                    QS.QSOpenApp();
                    botLog.AddBotLine("Opened the Qlik Sense app: " + QS.qsAppName, me);
                }
                catch (Exception e)
                {
                    botLog.AddBotLine(string.Format("Error opening the Qlik Sense app {0}: {1}", QS.qsAppName, e), me, LogFile.LogType.logError);
                }
                //NLP settup
                try
                {
                    //Prediction Pred = new Prediction(LuisURL, LuisAppID, LuisKey);
                    if (cntApiAiKey != null && cntApiAiKey != "")    // Priority for Google
                                                                     //if (LuisAppID == "")     // Priority for Microsoft
                    {
                        Conversation = new QlikConversationService.TelegramConversation(cntApiAiKey, cntApiAiLanguage);
                        //Pred.NLPStartApiAi(cntApiAiKey, cntApiAiLanguage);
                        botLog.AddBotLine("Created a conversation with Google Api.Ai.", me);
                    }
                    else
                    {
                        Conversation = new QlikConversationService.TelegramConversation(LuisURL, LuisAppID, LuisKey);
                        //Pred.NLPStartLUIS(LuisURL, LuisAppID, LuisKey);
                        botLog.AddBotLine("Created a conversation with Microsoft LUIS.", me);
                    }

                }
                catch (Exception e)
                {
                    botLog.AddBotLine(string.Format("Error opening the NLP engine {0}", e), me, LogFile.LogType.logError);
                }

                QlikUsers.ReadFromCSV(cntQlikUsersCSV);
                if (ConfigurationManager.AppSettings["cntAllowNewUsers"].ToUpper().StartsWith("Y"))
                    cntAllowNewUsers = true;
                else
                    cntAllowNewUsers = false;

                cntBotName = me.Username;
                Console.Title = me.Username;

                //send first message to the administrators
                try
                {
                    botClient.SendTextMessageAsync(cntTelegramAdministratorID, "The bot " + me.Username + " is now running.", replyMarkup: new ReplyKeyboardRemove()).Wait();
                    botClient.SendTextMessageAsync(cntBotLocalAdministrator, "The bot " + me.Username + " is now running.", replyMarkup: new ReplyKeyboardRemove()).Wait();
                }
                catch (Exception e)
                {
                    Console.Write("Send message to admin account Error:");
                    Console.Write(e.ToString());
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error when opening telegram bot client.");
                Console.WriteLine(e.ToString());
                return;
            }
        }

        private QSUser CheckTheUser(string UserId, Telegram.Bot.Types.User TelegramUser, string UserName = "")
        {
            QSUser Usr = QlikUsers.AddUser(UserId, UserName, LastAccess: DateTime.Now.ToString(QSUser.cntDateTimeParseFormat),
                Allowed: cntAllowNewUsers);

            if (Usr.Allowed)
            {
                if (Usr.QS == null)
                {
                    Usr.QS = new QSApp();
                    Usr.QS.qsAppName = cntqsAppName;
                    Usr.QS.qsAppId = cntqsAppId;
                    Usr.QS.qsServer = cntqsServer;
                    Usr.QS.qsSingleServer = cntqsSingleServer;
                    Usr.QS.qsSingleApp = cntqsSingleApp;
                    Usr.QS.qsAlternativeStreams = cntqsAlternativeStreams;

                    try
                    {
                        Usr.QS.QSConnectServerHeader(UserId, DemoqsServerHeaderAuth, cntqsServerVirtualProxy, cntqsServerSSL, cntCheckSDKVersion);

                        if (Usr.QS.IsConnected)
                        {
                            Usr.QS.QSOpenApp();
                            if (Usr.QS.AppIsOpen)
                            {
                                botLog.AddBotLine(string.Format("Opened the Qlik Sense app: {0} for user {1} ({2})", Usr.QS.qsAppName, Usr.UserId, Usr.UserName), TelegramUser);
                            }
                            else
                                botLog.AddBotLine(string.Format("Could not open the Qlik Sense app: {0} for user {1} ({2})", Usr.QS.qsAppName, Usr.UserId, Usr.UserName), TelegramUser);
                        }
                        else
                            botLog.AddBotLine(string.Format("Could not connect to the Qlik Sense server: {0} for user {1} ({2})", Usr.QS.qsServer, Usr.UserId, Usr.UserName), TelegramUser);
                    }
                    catch (Exception e)
                    {
                        botLog.AddBotLine(string.Format("Error opening the Qlik Sense app{0} for user {1}: {2}", Usr.QS.qsAppName, Usr.UserId, e), TelegramUser, LogFile.LogType.logError);
                    }
                }
                else
                {
                    Usr.QS.CheckConnection();
                }
            }

            return Usr;
        }

        //log error
        private void BotOnReceiveError(object sender, ReceiveErrorEventArgs receiveErrorEventArgs)
        {
            //Debugger.Break();
            botLog.AddBotLine(string.Format("BotOnReceiveError: {0}", receiveErrorEventArgs.ToString()), LogFile.LogType.logError);
        }

        //log error
          async void BotOnChosenInlineResultReceived(object sender, ChosenInlineResultEventArgs chosenInlineResultEventArgs)
        {
            botLog.AddBotLine($"Received choosen inline result: {chosenInlineResultEventArgs.ChosenInlineResult.ResultId}");
        }

        //when an inline query is received (@bot in any chat)
          async void BotOnInlineQueryReceived(object sender, InlineQueryEventArgs inlineQueryEventArgs)
        {
            Console.WriteLine("Inline query received");
            QSFoundObject[] qsFounds;
            string qsQuery;

            QSUser Usr = CheckTheUser(inlineQueryEventArgs.InlineQuery.From.Id.ToString(),
                inlineQueryEventArgs.InlineQuery.From,
                inlineQueryEventArgs.InlineQuery.From.FirstName + " " + inlineQueryEventArgs.InlineQuery.From.LastName);

            if (!Usr.Allowed)
            {
                InlineQueryResult[] results;
                results = new InlineQueryResult[1];
                results[0] = new InlineQueryResultArticle
                {
                    Id = "0",
                    Title = StringResources.secUserNotAllowed,
                    InputMessageContent = new InputTextMessageContent
                    {
                        MessageText = StringResources.secUserNotAllowed,
                        ParseMode = ParseMode.Html
                    },
                };

                try
                {
                    await botClient.AnswerInlineQueryAsync(inlineQueryEventArgs.InlineQuery.Id, results, isPersonal: true, cacheTime: 0);
                }
                catch (Exception e)
                {
                    botLog.AddBotLine(string.Format("{0} Exception caught.", e), LogFile.LogType.logError);
                }
                return;
            }
            if (DemoMode && (inlineQueryEventArgs.InlineQuery.From.Id == cntTelegramAdministratorID || inlineQueryEventArgs.InlineQuery.From.Id == cntBotLocalAdministrator))
            {
                //await BotSendTextMessage(inlineQueryEventArgs.InlineQuery.Id , StringResources.SorryDemoMode, replyMarkup: new ReplyKeyboardRemove());
                InlineQueryResult[] results;
                results = new InlineQueryResult[1];
                results[0] = new InlineQueryResultArticle
                {
                    Id = "0",
                    Title = StringResources.SorryDemoMode,
                    InputMessageContent = new InputTextMessageContent
                    {
                        MessageText = StringResources.SorryDemoMode,
                        ParseMode = ParseMode.Html
                    },
                };

                try
                {
                    await botClient.AnswerInlineQueryAsync(inlineQueryEventArgs.InlineQuery.Id, results, isPersonal: true, cacheTime: 0);
                }
                catch (Exception e)
                {
                    botLog.AddBotLine(string.Format("{0} Exception caught.", e), LogFile.LogType.logError);
                }
                return;
            }
            qsQuery = inlineQueryEventArgs.InlineQuery.Query;

            if (qsQuery.StartsWith("/"))
            {
                return;
            }

            if (qsQuery.Length > 0)
            {
                try
                {
                    qsFounds = Usr.QS.QSSearchObjects(inlineQueryEventArgs.InlineQuery.Query, true);
                }
                catch (Exception e)
                {
                    botLog.AddBotLine(string.Format("{0} Exception caught.", e), LogFile.LogType.logError);

                    qsFounds = new QSFoundObject[0];
                }

                if (qsFounds.Length == 0) return;

                InlineQueryResult[] results;
                results = new InlineQueryResult[qsFounds.Length];
                int i = 0;
                foreach (var f in qsFounds)
                {
                    results[i] = new InlineQueryResultArticle
                    {
                        Id = i.ToString(),
                        Title = f.Description,
                        InputMessageContent = new InputTextMessageContent
                        {
                            MessageText = f.HRef,
                            ParseMode = ParseMode.Html
                        },
                        Url = f.ObjectURL,
                        HideUrl = false,
                        ThumbUrl = f.ThumbURL

                        //,ThumbWidth = 25,
                        //ThumbHeight = 25
                    };
                    i++;
                }
                try
                {
                    await botClient.AnswerInlineQueryAsync(inlineQueryEventArgs.InlineQuery.Id, results, isPersonal: true, cacheTime: 0);
                }
                catch (Exception e)
                {
                    botLog.AddBotLine(string.Format("{0} Exception caught.", e), LogFile.LogType.logError);
                }
            }
        }

        //when an message is received
          async void BotOnMessageReceived(object sender, MessageEventArgs messageEventArgs)
        {
            Console.WriteLine("message received");
            var message1 = messageEventArgs.Message;
            Console.WriteLine(message1.Text);
            botClient.SendTextMessageAsync(message1.From.Id, message1.Text, replyMarkup: new ReplyKeyboardRemove()).Wait();

            try
            {
                var message = messageEventArgs.Message;

                if (message == null || message.Type == MessageType.ServiceMessage || message.Type == MessageType.UnknownMessage)
                    return;
                if (message.Entities.Count > 0 && message.Entities[0].Type == MessageEntityType.TextLink)
                    return;

                string MessageText = "";
                if (message.Text != null)
                    MessageText = message.Text;

                botLog.AddBotLine(string.Format("Message Received from {0}: {1}", message.From.Id.ToString(), MessageText), LogFile.LogType.logInfo);

                QSUser Usr = CheckTheUser(message.From.Id.ToString(),
                    message.From,
                    message.From.FirstName + " " + message.From.LastName);

                if (Usr.QS == null)
                {
                    await BotSendTextMessage(message.Chat.Id, string.Format(StringResources.secUserNotAllowed, Usr.UserName), parseMode: ParseMode.Html, replyMarkup: new ReplyKeyboardRemove());
                    return;
                }

                if (MessageText.StartsWith("/") && message.Chat.Type == ChatType.Group)
                {
                    if (MessageText.ToLower().StartsWith("/start@" + botClient.GetMeAsync().Result.Username.ToLower() + " addtogroup"))
                    {
                        await BotSendTextMessage(message.Chat.Id, string.Format(StringResources.cnvCreateGroupWelcome, Usr.UserName), parseMode: ParseMode.Html, replyMarkup: new ReplyKeyboardRemove());

                        await BotSendChartImage(message.Chat.Id, Usr.LastChart?.ObjectURL, NPrintingImgsPath + "\\" + Usr.UserId + ".png"
                            , Usr.LastChart?.Description == null ? "---" : Usr.LastChart?.Description);

                        return;
                    }
                    else
                        MessageText = MessageText.Substring(1);
                }


                if (!Usr.Allowed)
                {
                    await BotSendTextMessage(message.Chat.Id, StringResources.secUserNotAllowed, replyMarkup: new ReplyKeyboardRemove());
                    return;
                }

                if (WaitingGeoFilter > 0)
                {
                    WaitingGeoFilter--;
                    if (message.Type == MessageType.LocationMessage)
                    {
                        BotShowTypingState(message.Chat.Id);

                        string msg = "";
                        string urlfilter = "";

                        double Distance = WaitingGeoDistanceKm > 0 ? WaitingGeoDistanceKm : GetDefaultDistanceInKm();
                        QSMasterItem Dimension = (WaitingGeoDimension != null && WaitingGeoDimension != "") ? Usr.QS.GetMasterDimension(WaitingGeoDimension) : Usr.QS.LastDimension;
                        QSMasterItem Measure = Usr.QS.LastMeasure;

                        QSGeoFilter g = Usr.QS.GetGeoFilters(message.Location.Latitude, message.Location.Longitude, Distance);
                        List<QSGeoList> gl = Usr.QS.GetGeoList(g, Dimension, Measure, Usr.QS.LastFilters);

                        if (gl == null || gl.Count == 0)
                            await BotSendTextMessage(message.Chat.Id, string.Format(StringResources.geoElementNotFound, Dimension.Name, Distance)
                                , parseMode: ParseMode.Html, replyMarkup: new ReplyKeyboardRemove());
                        else
                        {
                            foreach (QSGeoList i in gl)
                            {
                                string Title = i.TextLabel + ": " + i.Text + "\n" + i.ValueLabel + ": " + i.FormattedValue;
                                await BotSendVenueMessage(message.Chat.Id, Convert.ToSingle(i.Lat), Convert.ToSingle(i.Lon), Title, i.Address,
                                    replyMarkup: new ReplyKeyboardRemove());

                                msg += Title + "\n";
                                if (i != gl.First()) urlfilter += ",";
                                urlfilter += Uri.EscapeDataString(i.Text);
                            }

                            string url = Usr.QS.qsSingleServer + "/single/?appid=" + Usr.QS.qsAppId + "&sheet=" + Usr.QS.Sheets.First().Id + "&opt=currsel";
                            url += "&select=" + Uri.EscapeDataString(Dimension.Expression) + "," + urlfilter;
                            msg = string.Format(StringResources.geoGotoAnalysisSheet, Usr.UserName, url);
                            //msg += "\n" + url;

                            await BotSendTextMessage(message.Chat.Id, msg, parseMode: ParseMode.Html, replyMarkup: new ReplyKeyboardRemove());

                        }
                    }
                }

                if (message.Type == MessageType.DocumentMessage)
                {
                    Document Doc = message.Document;
                    BotShowTypingState(message.Chat.Id);

                    string DocFile = NPrintingImgsPath + "\\" + Usr.UserId;
                    DocFile += Path.GetExtension(Doc.FileName);
                    var fileStream = System.IO.File.Create(DocFile);
                    var fts = new FileToSend(DocFile, fileStream);

                    Telegram.Bot.Types.File f = await botClient.GetFileAsync(Doc.FileId, fts.Content);
                    fts.Content.Seek(0, SeekOrigin.Begin);
                    fts.Content.CopyTo(fileStream);
                    fileStream.Close();

                    string MyAppId = Usr.QS.QSCreateApp(Path.GetFileName(DocFile), Usr.UserId, cntFolderConnection);

                    Usr.QS.QSOpenApp(MyAppId);
                    string msg;
                    if (Usr.QS.qsAppId == MyAppId)
                    {
                        msg = string.Format(StringResources.appOpenedApp, Usr.QS.qsAppName);
                        Usr.QS.QSCreateMasterItemsFromFields();

                        Response DocResp = await Conversation.Reply("#ShowAnalysis#", Usr);
                        if (DocResp.TextMessage == StringResources.nlMessageNotManaged && message.Chat.Type == ChatType.Group)
                        {
                            // Do nothing
                        }
                        else
                        {
                            if (cntStreamIdPublishNewApps.Length > 0) Usr.QS.QSPublishApp(cntStreamIdPublishNewApps);
                            ProcessConversationResponse(DocResp, message.Chat.Id, Usr);
                        }
                    }
                    else
                    {
                        msg = string.Format(StringResources.appOpenAppError, Usr.QS.qsAppId);
                    }
                    await BotSendTextMessage(message.Chat.Id, msg, parseMode: ParseMode.Html, replyMarkup: new ReplyKeyboardRemove());

                }

                if (message.Type != MessageType.TextMessage) return;

                if (DemoMode && messageEventArgs.Message.From.Id != cntTelegramAdministratorID & messageEventArgs.Message.From.Id != cntBotLocalAdministrator)
                {
                    await BotSendTextMessage(message.Chat.Id, StringResources.SorryDemoMode, replyMarkup: new ReplyKeyboardRemove());
                    return;
                }


                if (messageEventArgs.Message.From.Id == cntTelegramAdministratorID
                   || messageEventArgs.Message.From.Id == cntBotLocalAdministrator)
                {

                    if (MessageText == "/demo")
                    {
                        DemoMode = !DemoMode;
                        if (DemoMode)
                            await BotSendTextMessage(message.Chat.Id, StringResources.DemoActivated, replyMarkup: new ReplyKeyboardRemove());
                        else
                            await BotSendTextMessage(message.Chat.Id, "Demo mode deactivated", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    if (MessageText.ToLower() == "/allownewusers")
                    {
                        cntAllowNewUsers = !cntAllowNewUsers;
                        await BotSendTextMessage(message.Chat.Id, "AllowNewUsers: " + cntAllowNewUsers.ToString(), replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    if (MessageText.ToLower().StartsWith("/allowuser "))
                    {
                        string[] parm = MessageText.Split(' ');

                        if (parm.Length > 1)
                        {
                            string UserId = parm[1].Trim();
                            if (UserId == cntTelegramAdministratorID.ToString() || UserId == cntBotLocalAdministrator.ToString()) return;
                            QSUser U = QlikUsers.GetUser(UserId);
                            if (U == null)
                                await BotSendTextMessage(message.Chat.Id, "User ID not found");
                            else
                            {
                                U.Allowed = !U.Allowed;
                                QlikUsers.WriteToCSV(cntQlikUsersCSV);

                                await BotSendTextMessage(message.Chat.Id, string.Format("User {0} has been set Allowed to <{1}>", UserId, U.Allowed.ToString()));
                            }

                        }
                        return;
                    }

                    if (MessageText.ToLower().StartsWith("/masteritems"))
                    {
                        System.Text.StringBuilder msg = new System.Text.StringBuilder();

                        msg.AppendLine("Measures:");
                        msg.AppendLine("---------");
                        foreach (QSMasterItem mm in Usr.QS.MasterMeasures) msg.AppendLine(mm.Name);
                        await BotSendTextMessage(message.Chat.Id, msg.ToString());

                        msg.Clear();
                        msg.AppendLine("Dimensions:");
                        msg.AppendLine("-----------");
                        foreach (QSMasterItem md in Usr.QS.MasterDimensions) msg.AppendLine(md.Name);
                        await BotSendTextMessage(message.Chat.Id, msg.ToString());

                        return;
                    }

                    else if (MessageText.ToLower() == "send me the log")
                    {
                        await BotSendLogToAdmin(messageEventArgs.Message.From.Id);
                        return;
                    }

                    else if (MessageText.ToLower() == "send me the users")
                    {
                        await BotSendUsersToAdmin(messageEventArgs.Message.From.Id);
                        return;
                    }


                    else if (MessageText.ToLower() == "language")
                    {
                        var keyboard = new ReplyKeyboardMarkup(new[]
                        {
                                new [] // first row
                                {
                                    new KeyboardButton("Español"),
                                    new KeyboardButton("English")
                                 },
                                new []  // second row
                                {
                                   new KeyboardButton("Português"),
                                   new KeyboardButton("Italiano")
                                },
                                new []  // third row
                                {
                                   new KeyboardButton("Pусский"),
                                   new KeyboardButton("Français")
                                },
                                new []  // forth row
                                {
                                   new KeyboardButton("Português-BR")
                                }
                            });

                        await BotSendTextMessage(message.Chat.Id, "Language/Idioma", replyMarkup: keyboard);
                        return;
                    }

                    // Change Language to Spanish
                    else if (MessageText.ToLower() == "español")
                    {
                        ChangeAppLanguage("es-ES");
                        await BotSendTextMessage(message.Chat.Id, "Cambiado el idioma a Español", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    // Change Language to English
                    else if (MessageText.ToLower() == "english")
                    {
                        ChangeAppLanguage("en-US");
                        await BotSendTextMessage(message.Chat.Id, "Language changed to English", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    // Change Language to Portuguese
                    else if (MessageText.ToLower() == "português")
                    {
                        ChangeAppLanguage("pt-PT");
                        await BotSendTextMessage(message.Chat.Id, "Linguagem mudado para Português", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    // Change Language to Italiano
                    else if (MessageText.ToLower() == "italiano")
                    {
                        ChangeAppLanguage("it-IT");
                        await BotSendTextMessage(message.Chat.Id, "Cambiato la lingua Italiana per", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    // Change Language to Russian
                    else if (MessageText.ToLower() == "pусский")
                    {
                        ChangeAppLanguage("ru-RU");
                        await BotSendTextMessage(message.Chat.Id, "Язык изменен на Pусский", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    // Change Language to French
                    else if (MessageText.ToLower() == "français")
                    {
                        ChangeAppLanguage("fr-FR");
                        await BotSendTextMessage(message.Chat.Id, "Langue changée en Français", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }

                    // Change Language to Brazilian
                    else if (MessageText.ToLower() == "português-br")
                    {
                        ChangeAppLanguage("pt-BR");
                        await BotSendTextMessage(message.Chat.Id, "Linguagem mudado para Português-BR", replyMarkup: new ReplyKeyboardRemove());
                        return;
                    }
                }


                if (MessageText.ToLower().StartsWith("/start"))
                {
                    await BotSendTextMessage(message.Chat.Id, string.Format(StringResources.Welcome, message.From.FirstName), replyMarkup: new ReplyKeyboardRemove());
                    return;
                }

                //zhu else if (MessageText.ToLower().StartsWith(StringResources.ThankYou.ToLower()))
                //{
                //    await BotSendTextMessage(message.Chat.Id, StringResources.YoureWelcome, replyMarkup: new ReplyKeyboardRemove());
                //    return;
                //}




                BotShowTypingState(message.Chat.Id);

                Response Resp = await Conversation.Reply(MessageText, Usr);
                if (Resp.TextMessage == StringResources.nlMessageNotManaged && message.Chat.Type == ChatType.Group)
                {
                    // Do nothing
                }
                else
                {
                    ProcessConversationResponse(Resp, message.Chat.Id, Usr);
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("General Error in BotOnMessageReceived: {0}", e), LogFile.LogType.logError);
            }

        }

        private async void ProcessConversationResponse(Response Resp, long ChatId, QSUser Usr)
        {
            if (Resp.OtherAction.StartsWith("Language="))
            {
                if (ChatId == cntTelegramAdministratorID || ChatId == cntBotLocalAdministrator)
                {
                    string lang = Resp.OtherAction.Substring("Language=".Length);
                    ChangeAppLanguage(lang);
                }
                else
                {
                    return;
                }
            }


            if (Resp.Options.Count() > 0)
            {
                BotShowTypingState(ChatId);

                InlineKeyboardButton[][] rows = new InlineKeyboardButton[(Resp.Options.Count() + 1) / 2][];

                int r = -1;
                for (int i = 0; i < Resp.Options.Count(); i++)
                {
                    string ButtonData = Resp.Options[i].Action.ToString() + "#" + Resp.Options[i].ID;
                    //zhu
                    //var b = new InlineKeyboardButton(Resp.Options[i].Title, ButtonData);
                    if (i % 2 == 0)
                    {
                        r++;
                        if (i == Resp.Options.Count() - 1)   // Last button
                            rows[r] = new InlineKeyboardButton[1];  // new row, only 1 button
                        else
                            rows[r] = new InlineKeyboardButton[2];  // new row, two buttons

                        //rows[r][0] = b;
                        rows[r][0] = Resp.Options[i].Title;
                    }
                    //else rows[r][1] = b;
                    rows[r][1] = Resp.Options[i].Title;
                }

                var keyboard = new InlineKeyboardMarkup(rows);

                await BotSendTextMessage(ChatId, Resp.TextMessage, replyMarkup: keyboard);
                Resp.TextMessage = "";
            }


            if (Resp.OtherAction == "Help")
            {
                Resp.TextMessage.Replace("QlikSenseBot", cntBotName);
                if (ConfigurationManager.AppSettings["cntHelpInformation"] != null)
                    Resp.TextMessage = ConfigurationManager.AppSettings["cntHelpInformation"] + "\r\n" + Resp.TextMessage;
            }

            if (Resp.ErrorText != "")
                botLog.AddBotLine(Resp.ErrorText, LogFile.LogType.logError);

            if (Resp.WarningText != "")
                botLog.AddBotLine(Resp.WarningText, LogFile.LogType.logWarning);

            Resp.TextMessage = Resp.TextMessage.Replace(":-)", "😊");
            Resp.TextMessage = Resp.TextMessage.Replace(":-(", "🙁");
            Resp.TextMessage = Resp.TextMessage.Replace(";-)", "😉");

            if (Resp.TextMessage != "")
                await BotSendTextMessage(ChatId, Resp.TextMessage, parseMode: ParseMode.Html, replyMarkup: new ReplyKeyboardRemove());

            if (Resp.VoiceMessage != "")
                //zhu await Speak(ChatId, Usr, Resp.VoiceMessage);

            if (Resp.ChartFound != null && Resp.ChartFound.ObjectType != "map")
            {
                string ChartUrl = "";
                if (cntCaptureWeb != null && cntCaptureWeb.Trim().Length > 0)
                    ChartUrl = Resp.ChartFound.ObjectURL.Replace(Usr.QS.qsSingleServer, cntCaptureWeb);
                else
                    ChartUrl = Resp.ChartFound.ObjectURL;

                await BotSendChartImage(ChatId, ChartUrl, NPrintingImgsPath + "\\" + Usr.UserId + ".png"
                    , Resp.ChartFound.Description == null ? "---" : Resp.ChartFound.Description);
            }


            if (Resp.OtherAction == "ShowReports")
                //zhu await ShowReportList(ChatId);

            if (Resp.OtherAction == "GeoFilter")
            {
                WaitingGeoFilter = 3;
                WaitingGeoDimension = Resp.NLPrediction.Dimension;
                WaitingGeoDistanceKm = Resp.NLPrediction.DistanceKm;

                var keyboard = new ReplyKeyboardMarkup(new[]
                {
                    new KeyboardButton(StringResources.gasAccessLocation)
                    {
                        RequestLocation = true
                    }
                });

                await BotSendTextMessage(ChatId, StringResources.gasButtonLocation, replyMarkup: keyboard);
            }

            if (Resp.OtherAction == "CreateGroup")
            {
                await CreateChatGroup(ChatId, "addtogroup");
            }

            if (Resp.NewsSearch != "")
                await AskForNews(ChatId, Resp.NewsSearch);


        }

        private async Task CreateChatGroup(long ChatId, string GroupName)
        {
            string url = "https://telegram.me/" + botClient.GetMeAsync().Result.Username.ToLower() + "?startgroup=" + GroupName;
            string msg = string.Format(StringResources.urlCreateGroup, url);

            await BotSendTextMessage(ChatId, msg, parseMode: ParseMode.Html, disableWebPagePreview: true, replyMarkup: new ReplyKeyboardRemove());
        }

        private  async Task AskForNews(long ChatId, string NewsQuery)
        {
            if (cntBingSearchKey == null || cntBingSearchKey.Trim().Length == 0) return;

            //zhu var keyboard = new InlineKeyboardMarkup(new[]
            //{
            //    new[] // first row
            //    {
            //        new InlineKeyboardButton(StringResources.bsButtonYes, "#BingNewsYes" + NewsQuery)
            //    }
            //});

            //await BotSendTextMessage(ChatId, string.Format(StringResources.bsWantToSearch, NewsQuery), replyMarkup: keyboard);
        }
        //show a typing state
        async void BotShowTypingState(long ChatId)
        {
            try
            {
                await botClient.SendChatActionAsync(ChatId, ChatAction.Typing);
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("{0} Exception caught.", e), LogFile.LogType.logError);
            }
        }

        //when a inline keyboard button is pressed
        async void BotOnCallbackQueryReceived(object sender, CallbackQueryEventArgs callbackQueryEventArgs)
        {
            Console.WriteLine("call back query received");
        }

        private async Task<bool> BotSendChartImage(long chatID, string ChartUrl, string ImageFile, string Caption = "")
        {
            if (cntCaptureImageApp == null || ChartUrl.Trim().Length == 0 || ImageFile.Trim().Length == 0) return false;

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = cntCaptureImageApp;
            //startInfo.WindowStyle = ProcessWindowStyle.Normal;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = "\"" + ChartUrl.Replace("&opt=currsel", "") + "\" \"" + ImageFile + "\"";

            int Timeout;
            int AppTimeout;
            try
            {
                Timeout = Convert.ToInt32(cntCaptureTimeout) * 1000;
            }
            catch (Exception e)
            {
                Timeout = 5000;
            }

            if (Timeout <= 0)
                Timeout = 5000;
            AppTimeout = Timeout + 1000;

            startInfo.Arguments += " " + Timeout.ToString();

            try
            {
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit(AppTimeout);
                    if (exeProcess.HasExited)
                    {
                        await BotSendPhotoMessage(chatID, Caption, ImageFile);
                        return true;
                    }
                    else
                    {
                        botLog.AddBotLine("Error capturing chart image", LogFile.LogType.logError);
                        exeProcess.Kill();
                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed sending a chart image: {0}", e), LogFile.LogType.logError);
                return false;
            }
        }

        async Task<Message> BotSendTextMessage(long chatId, string text, bool disableWebPagePreview = false, bool disableNotification = false, int replyToMessageId = 0, IReplyMarkup replyMarkup = null, ParseMode parseMode = ParseMode.Default, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (text.Trim() == "" || text == null) return null;
            if (text.Length > 4096) text = text.Substring(0, 4090) + "...";

            Message m = new Message();

            try
            {
                m = await botClient.SendTextMessageAsync(chatId, text,parseMode,disableWebPagePreview, disableNotification, replyToMessageId, replyMarkup, cancellationToken);
                botLog.AddBotLine("Sent text message <" + text + "> to ChatID " + chatId.ToString());
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed send text message. {0} Exception caught.", e), LogFile.LogType.logError);
            }
            return m;
        }

          async Task<Message> BotSendPhotoMessage(long chatId, string text, string FilePhoto, bool disableWebPagePreview = false, bool disableNotification = false, int replyToMessageId = 0, IReplyMarkup replyMarkup = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (FilePhoto.Trim() == "" || FilePhoto == null) return null;
            Message m = new Message();

            try
            {
                using (var fileStream = new FileStream(FilePhoto, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var fts = new FileToSend(FilePhoto, fileStream);
                    m = await botClient.SendPhotoAsync(chatId, fts, text, disableNotification, replyToMessageId, replyMarkup, cancellationToken);
                    botLog.AddBotLine("Sent photo message <" + FilePhoto + "> with text <" + text + "> to ChatID " + chatId.ToString());
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed send photo message. {0} Exception caught.", e), LogFile.LogType.logError);
            }
            return m;
        }

          async Task<Message> BotSendDocMessage(long chatId, string text, string FileDoc, bool disableWebPagePreview = false, bool disableNotification = false, int replyToMessageId = 0, IReplyMarkup replyMarkup = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (FileDoc.Trim() == "" || FileDoc == null) return null;
            if (text.Length > 4096) text = text.Substring(0, 4090) + "...";

            Message m = new Message();

            try
            {
                using (var fileStream = new FileStream(FileDoc, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var fts = new FileToSend(FileDoc.Split('\\').LastOrDefault(), fileStream);
                    m = await botClient.SendDocumentAsync(chatId, fts, text, disableNotification, replyToMessageId, replyMarkup, cancellationToken);
                    botLog.AddBotLine("Sent doc message <" + FileDoc + "> with text <" + text + "> to ChatID " + chatId.ToString());
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed send doc message. {0} Exception caught.", e), LogFile.LogType.logError);
            }
            return m;
        }

          async Task<Message> BotSendVoiceMessage(long chatId, string text, string FileDoc, bool disableWebPagePreview = false, bool disableNotification = false, int replyToMessageId = 0, IReplyMarkup replyMarkup = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (FileDoc.Trim() == "" || FileDoc == null) return null;
            if (text.Length > 4096) text = text.Substring(0, 4090) + "...";

            Message m = new Message();

            try
            {
                using (var fileStream = new FileStream(FileDoc, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var fts = new FileToSend(FileDoc.Split('\\').LastOrDefault(), fileStream);
                    m = await botClient.SendVoiceAsync(chatId, fts,"", 0, disableNotification, replyToMessageId, replyMarkup, cancellationToken);
                    botLog.AddBotLine("Sent voice message <" + FileDoc + "> with text <" + text + "> to ChatID " + chatId.ToString());
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed send voice message. {0} Exception caught.", e), LogFile.LogType.logError);
            }
            return m;
        }

          async Task<Message> BotSendVenueMessage(long chatId, float Latitude, float Longitude, string Title, string Address,
            bool disableNotification = false, int replyToMessageId = 0, IReplyMarkup replyMarkup = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (float.IsNaN(Latitude) || float.IsNaN(Longitude)) return null;

            if (Title.Length > 4096) Title = Title.Substring(0, 4090) + "...";

            Message m = new Message();

            try
            {
                m = await botClient.SendVenueAsync(chatId, Latitude, Longitude, Title, Address,
                    null, disableNotification, replyToMessageId, replyMarkup, cancellationToken);
                botLog.AddBotLine("Sent venue message <" + Latitude.ToString() + "," + Longitude.ToString() + "> with Title <" + Title + "> and Address <" + Address + "> to ChatID " + chatId.ToString());
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed send Venue message. {0} Exception caught.", e), LogFile.LogType.logError);
            }
            return m;
        }

          async Task<Message> BotSendLogToAdmin(int AdminUserID)
        {
            if (AdminUserID != cntTelegramAdministratorID && AdminUserID != cntBotLocalAdministrator) return (null);

            Message m = new Message();

            try
            {
                string FileDoc = botLog.GetLogFileName();
                using (var fileStream = new FileStream(botLog.GetLogFileName(), FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var fts = new FileToSend(FileDoc.Split('\\').LastOrDefault().Replace(".log", ".txt"), fileStream);
                    m = await botClient.SendDocumentAsync(cntTelegramAdministratorID, fts, "Here you have the current log");
                    botLog.AddBotLine("Sent the log to the Administrator <" + FileDoc + ">");
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed send Log to administrator. {0} Exception caught.", e), LogFile.LogType.logError);
            }
            return m;
        }

          async Task<Message> BotSendUsersToAdmin(int AdminUserID)
        {
            if (AdminUserID != cntTelegramAdministratorID && AdminUserID != cntBotLocalAdministrator) return (null);

            Message m = new Message();

            try
            {
                using (var fileStream = new FileStream(cntQlikUsersCSV, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var fts = new FileToSend(cntQlikUsersCSV.Split('\\').LastOrDefault().Replace(".csv", ".txt"), fileStream);
                    m = await botClient.SendDocumentAsync(cntTelegramAdministratorID, fts, "Here you have the current user list");
                    botLog.AddBotLine("Sent the user list to the Administrator <" + cntQlikUsersCSV + ">");
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Failed send user list to administrator. {0} Exception caught.", e), LogFile.LogType.logError);
            }
            return m;
        }

          async void CheckAlerts(Object source, System.Timers.ElapsedEventArgs et)
        {
        
        }

          string[] GetReportList()
        {
            string[] ReportFiles;

            try
            {
                ReportFiles = Directory.GetFiles(NPrintingImgsPath, "*.pdf", SearchOption.TopDirectoryOnly);
                for (int i = 0; i < ReportFiles.Length; i++)
                {
                    ReportFiles[i] = Path.GetFileNameWithoutExtension(ReportFiles[i]);
                }
            }
            catch (Exception e)
            {
                botLog.AddBotLine(string.Format("Error looking for reports in {0}: {1}", NPrintingImgsPath, e.ToString(), e), LogFile.LogType.logError);
                ReportFiles = new string[0];
            }

            return ReportFiles;
        }

          void ChangeAppLanguage(string LanguageName)
        {
            CultureInfo NewCulture = CultureInfo.CreateSpecificCulture(LanguageName);

            Thread.CurrentThread.CurrentUICulture = NewCulture;
            Thread.CurrentThread.CurrentCulture = NewCulture;
            CultureInfo.DefaultThreadCurrentUICulture = NewCulture;
            CultureInfo.DefaultThreadCurrentCulture = NewCulture;

            CurrentLanguage = LanguageName;

            try
            {
                LuisURL = ConfigurationManager.AppSettings["cntLuisURL" + "-" + LanguageName.ToLower()];
                LuisAppID = ConfigurationManager.AppSettings["cntLuisAppID" + "-" + LanguageName.ToLower()];
                LuisKey = ConfigurationManager.AppSettings["cntLuisKey" + "-" + LanguageName.ToLower()];

                if (LuisURL == "" || LuisAppID == "" || LuisKey == "") throw new System.Exception();
            }
            catch (Exception e)
            {
                LuisURL = ConfigurationManager.AppSettings["cntLuisURL"];
                LuisAppID = ConfigurationManager.AppSettings["cntLuisAppID"];
                LuisKey = ConfigurationManager.AppSettings["cntLuisKey"];
            }

        }

          double GetDefaultDistanceInKm()
        {
            double Distance;

            NumberFormatInfo provider = new NumberFormatInfo();
            provider.NumberDecimalSeparator = ".";
            provider.NumberGroupSeparator = ",";
            provider.NumberGroupSizes = new int[] { 3 };

            try
            {
                string strDistance = ConfigurationManager.AppSettings["GeoKmSelection"];
                Distance = Convert.ToDouble(ConfigurationManager.AppSettings["GeoKmSelection"], provider);
            }
            catch (Exception e)
            {
                Distance = 5;
            }

            return Distance;
        }

    }
}

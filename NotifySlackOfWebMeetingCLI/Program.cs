using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using CommandLine;
using System.Text.RegularExpressions;
using System.IO;
using System.Net.Http;
using Newtonsoft.Json;
using NotifySlackOfWebMeetingCLI.Settings;
using NotifySlackOfWebMeetingCLI.WebMeetings;
using JsonSerializer = System.Text.Json.JsonSerializer;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace NotifySlackOfWebMeetingCLI
{
    class Program
    {
        /// <summary>
        /// HTTPクライアント
        /// </summary>
        private static HttpClient s_HttpClient = new HttpClient();

        [Verb("setting", HelpText = "Register Slack channel information and create a configuration file.")]
        public class SettingOptions
        {
            [Option('n', "name", HelpText = "The Slack channel name.", Required = true)]
            public string Name { get; set; }

            [Option('u', "url", HelpText = "The web service endpoint url.", Required = true)]
            public string EndpointUrl { get; set; }

            [Option('w', "webhookUrl", HelpText = "The web hook url.", Required = true)]
            public string WebhookUrl { get; set; }

            [Option('r', "register", HelpText = "The registered name.", Required = true)]
            public string RegisteredBy { get; set; }
        }
        [Verb("register", HelpText = "Register the web conference information to be notified.")]
        public class RegisterOptions
        {
            [Option('f', "filepath", HelpText = "Setting file path.", Default = "./setting.json")]
            public string Filepath { get; set; }
        }

        static int Main(string[] args)
        {
            Func<SettingOptions, int> RunSettingAndReturnExitCode = opts =>
            {
                Console.WriteLine("Run Setting");

                return 1;
            };
            Func<RegisterOptions, int> RunRegisterAndReturnExitCode = opts =>
            {
                Console.WriteLine("Run Register");
                Console.WriteLine($"filepath:{opts.Filepath}");

                var application = new Outlook.Application();

                #region ログインユーザーのOutlookから、翌稼働日の予定を取得

                // ログインユーザーのOutlookの予定表フォルダを取得
                Outlook.Folder calFolder =
                    application.Session.GetDefaultFolder(
                            Outlook.OlDefaultFolders.olFolderCalendar)
                        as Outlook.Folder;

                DateTime start = DateTime.Today.AddDays(1);
                DateTime end = start.AddDays(1);
                Outlook.Items nextOperatingDayAppointments = GetAppointmentsInRange(calFolder, start, end);

                #endregion

                #region 取得した予定一覧の中からWeb会議情報を含む予定を抽出

                var webMeetingAppointments = new List<Outlook.AppointmentItem>();

                // ZoomURLを特定するための正規表現
                var zoomUrlRegexp = @"https?://[^(?!.*(/|.|\n).*$)]*\.?zoom\.us/[A-Za-z0-9/?=]+";

                foreach (Outlook.AppointmentItem nextOperatingDayAppointment in nextOperatingDayAppointments)
                {
                    // ZoomURLが本文に含まれる予定を正規表現で検索し、リストに詰める
                    if (Regex.IsMatch(nextOperatingDayAppointment.Body, zoomUrlRegexp))
                    {
                        webMeetingAppointments.Add(nextOperatingDayAppointment);
                    }
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているSlackチャンネル情報のIDと抽出した予定を使い、Web会議情報を作成

                // jsonファイルから設定を取り出す
                var fileContent = string.Empty;
                using (var sr = new StreamReader(opts.Filepath, Encoding.GetEncoding("utf-8")))
                {
                    fileContent = sr.ReadToEnd();
                }

                var setting = JsonConvert.DeserializeObject<Setting>(fileContent);

                // 追加する会議情報の一覧を作成
                var addWebMettings = new List<WebMeeting>();
                foreach (var webMeetingAppintment in webMeetingAppointments)
                {
                    var url = Regex.Match(webMeetingAppintment.Body, zoomUrlRegexp).Value;
                    var name = webMeetingAppintment.Subject;
                    var startDateTime = webMeetingAppintment.Start;
                    var addWebMetting = new WebMeeting()
                    {
                        Name = name,
                        StartDateTime = startDateTime,
                        Url = url,
                        RegisteredBy = setting.RegisteredBy,
                        SlackChannelId = setting.SlackChannelId
                    };
                    addWebMettings.Add(addWebMetting);
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を削除



                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を登録

                var postUrl = $"{((string)setting.EndpointUrl)}{"WebMeetings"}";

                // Web会議情報を登録
                foreach (var addWebMetting in addWebMettings)
                {
                    var postData = JsonConvert.SerializeObject(addWebMetting);
                    var content = new StringContent(postData, Encoding.UTF8, "application/json");
                    var response = s_HttpClient.PostAsync(postUrl, content).ConfigureAwait(true);
                }

                #endregion

                return 1;
            };

            return CommandLine.Parser.Default.ParseArguments<SettingOptions, RegisterOptions>(args)
                .MapResult(
                    (SettingOptions opts) => RunSettingAndReturnExitCode(opts),
                    (RegisterOptions opts) => RunRegisterAndReturnExitCode(opts),
                    errs => 1);
        }

        /// <summary>
        /// 指定したOutlookフォルダから指定期間の予定を取得する
        /// </summary>
        /// <param name="folder">Outlookフォルダ</param>
        /// <param name="startTime">開始日時</param>
        /// <param name="endTime">終了日時</param>
        /// <returns>Outlook.Items</returns>
        private static Outlook.Items GetAppointmentsInRange(
            Outlook.Folder folder, DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '"
                            + startTime.ToString("g")
                            + "' AND [End] <= '"
                            + endTime.ToString("g") + "'";
            Debug.WriteLine(filter);
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }
    }
}

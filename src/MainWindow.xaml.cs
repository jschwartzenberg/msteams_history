using MSTeamsHistory.Helpers;
using MSTeamsHistory.Models.Db;
using MSTeamsHistory.Models.Graph;
using MSTeamsHistory.Models.Graph.Chats;
using MSTeamsHistory.Models.Graph.Members;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using MSTeamsHistory.Helpers;
using System.Text.RegularExpressions;
using MSTeamsHistory.ShareGateModels;
using System.Collections;
using System.Reflection;

namespace MSTeamsHistory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string path = $@"C:\Users\{Environment.UserName}\Desktop\MSTeamsHistory";

        //Set the scope for API call to user.read
        string[] scopes = new string[] { 
            "user.read", 
            "Chat.Read",
            "User.ReadBasic.All"
        };


        public MainWindow()
        {
            InitializeComponent();
            HistoryText.Text = path;
        }

        /// <summary>
        /// Call AcquireToken - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            LogText.Text = string.Empty;
            LogText.IsReadOnly = true;
            if (string.IsNullOrEmpty(HistoryText.Text))
            {
                HistoryText.Text = path;
            }

            var accounts = await app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();


            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(accounts.FirstOrDefault())
                        .WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    LogText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                LogText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {
                TokenHolder tokenHolder = new TokenHolder(app, scopes, firstAccount, authResult);
                this.SignOutButton.Visibility = Visibility.Visible;
                LogText.Text = "Loading data ...";

                var me = await LoadItem<User>("https://graph.microsoft.com/beta/me", tokenHolder);

                var path = HistoryText.Text;
                var dbPath = Path.Combine(path, authResult.Account.Username); ;
                if (!System.IO.Directory.Exists(dbPath))
                {
                    System.IO.Directory.CreateDirectory(dbPath);
                }

                var chatsList = new List<Models.Graph.Chats.Chat>();
                var url = new Uri("https://graph.microsoft.com/beta/me/chats");
                do
                {
                    var chatsObj = await LoadItems<Models.Graph.Chats.Chat>(url.OriginalString, tokenHolder);
                    url = chatsObj.OdataNextLink;
                    chatsList.AddRange(chatsObj.Value);
                } while (url != null);

                System.IO.File.WriteAllText(Path.Combine(dbPath, "chats.json"), JsonConvert.SerializeObject(chatsList));
                var chatsPath = Path.Combine(dbPath, "chats"); ;
                if (!System.IO.Directory.Exists(chatsPath))
                {
                    System.IO.Directory.CreateDirectory(chatsPath);
                }


                int i = 0;
                foreach (var chat in chatsList)
                {
                    LogText.Text = $"Loading messages for chat {++i}/{chatsList.Count}";
                    var chatDirPath = Path.Combine(chatsPath, chat.Id.SHA1());
                    if (!System.IO.Directory.Exists(chatDirPath))
                    {
                        System.IO.Directory.CreateDirectory(chatDirPath);
                    }

                    System.IO.File.WriteAllText(Path.Combine(chatDirPath, "chat.json"), JsonConvert.SerializeObject(chat));

                    var messagesPath = Path.Combine(chatDirPath, "messages.json");

                    var listMessages = new List<Message>();
                    var messages = await LoadItems<Message>($"https://graph.microsoft.com/beta/me/chats/{chat.Id}/messages", tokenHolder);
                    if (messages.OdataCount > 0)
                    {
                        listMessages.AddRange(messages.Value);
                        do
                        {
                            var newMessages = await LoadItems<Message>(messages.OdataNextLink.ToString(), tokenHolder);
                            if (messages.OdataNextLink != null && newMessages.OdataNextLink != null && messages.OdataNextLink.ToString().Equals(newMessages.OdataNextLink.ToString()))
                            {
                                LogText.Text = $"MS Graph API loop detected! Affected chat is {chat.Id.SHA1()}.";
                                break;
                            }
                            messages = newMessages;
                            if (messages.OdataCount==0)
                            {
                                break;
                            }
                            listMessages.AddRange(messages.Value);
                        }while(true);
                    }

                    var x_messages = listMessages.OrderBy(x => x.CreatedDateTime).ToList();
                    System.IO.File.WriteAllText(messagesPath, JsonConvert.SerializeObject(x_messages));

                    var members = await LoadItems<Member>($"https://graph.microsoft.com/beta/me/chats/{chat.Id}/members", tokenHolder);

                    if (members.Value!=null&&
                        members.Value.Count>0&& x_messages.Count>0)
                    {
                        System.IO.File.WriteAllText(Path.Combine(chatDirPath, "members.json"), JsonConvert.SerializeObject(members.Value));
                        var arr = members.Value.Where(x => x.UserId.ToString() != me.Id).ToList();
                        var history = string.Empty;
                        if (arr.Count > 0)
                        {
                            history = string.Join(",", arr.Select(x => x.DisplayName));
                            if (history.Length>100)
                            {
                                history = history.Remove(100)+"...";
                            }
                        }
                        if (string.IsNullOrEmpty(history))
                        {
                            history = "history";
                        }
                            var data = x_messages.Select(x =>
                            {
                                var text = System.Net.WebUtility.HtmlDecode(x.Body.Content.StripHTML());
                                text = Regex.Replace(text, @"^\s+$[\r\n]*", string.Empty, RegexOptions.Multiline);
                                var str = $"{x.CreatedDateTime?.ToString("yyyy-MM-dd HH:mm:ss")} {text}";
                                return str;
                            });
                            System.IO.File.WriteAllLines(Path.Combine(chatDirPath, $"{history}.txt"), data);
                    }

                    var shareGateChatDirPath = Path.Combine(chatDirPath, "sharegate");
                    if (!System.IO.Directory.Exists(shareGateChatDirPath))
                    {
                        System.IO.Directory.CreateDirectory(shareGateChatDirPath);
                    }

                    var shareGateChatDirAttachmentsPath = Path.Combine(shareGateChatDirPath, "Messages Attachments");
                    if (!System.IO.Directory.Exists(shareGateChatDirAttachmentsPath))
                    {
                        System.IO.Directory.CreateDirectory(shareGateChatDirAttachmentsPath);
                    }

                    var shareGateMessagesPath = Path.Combine(shareGateChatDirPath, "Messages.json");
                    var sharegate_messages = ConvertToShareGate(x_messages, shareGateChatDirAttachmentsPath, tokenHolder);
                    var sharegate_messagesDotJson = WrapShareGateMessages(await sharegate_messages);
                    var json_string = UnDoDoubleEscaping(JsonConvert.SerializeObject(sharegate_messagesDotJson));
                    System.IO.File.WriteAllText(shareGateMessagesPath, json_string);

                    var shareGateMessagesDotAspxTemplatePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Messages.aspx.template");
                    var shareGateMessagesDotAspxTemplate = System.IO.File.ReadAllText(shareGateMessagesDotAspxTemplatePath);

                    var shareGateMessagesDotAspxPath = Path.Combine(shareGateChatDirPath, "Messages.aspx");
                    System.IO.File.WriteAllText(shareGateMessagesDotAspxPath, shareGateMessagesDotAspxTemplate.Replace("####INSERT_JSON_HERE####", json_string));
                }
                LogText.Text = "Done.";
            }
        }

        private string UnDoDoubleEscaping(string v)
        {
            return v.Replace("\\\\", "\\").Replace("\\\\", "\\\\\\");
        }

        private List<SgMessagesDotJsonElement> WrapShareGateMessages(List<List<SgMessage>> sharegate_messages) =>
            sharegate_messages.Select(messageList =>
                {
                    var element = new SgMessagesDotJsonElement();
                    element.Message = ShareGateEscaped(JsonConvert.SerializeObject(messageList.First()));
                    element.Replies = messageList.Skip(1).Select(message =>
                        {
                            return ShareGateEscaped(JsonConvert.SerializeObject(message));
                        }
                    ).ToList();
                    return element;
                }
            ).ToList();

        private string ShareGateEscaped(string v)
        {
            return v.Replace("\"", "\\u0022")
                    .Replace("&", "\\u0026")
                    .Replace("'", "\\u0027")
                    .Replace("<", "\\u003c")
                    .Replace(">", "\\u003e")
                    .Replace(@"\n", "\\u005c\\u006e")
                    .Replace(@"\t", "\\u005c\\u0074")
                    ;
        }

        private async Task<List<List<SgMessage>>> ConvertToShareGate(List<Message> x_messages, string attachmentsPath, TokenHolder accessToken)
        {
            var retList = new List<List<SgMessage>>();
            DateTime previousCurrentDay = DateTime.MinValue;
            List<SgMessage> currentDayList = null;
            var pictureCacheIds = new Dictionary<string, string>();
            foreach (var message in x_messages)
            {
                var messageDay = message.LastModifiedDateTime.Value.DateTime.Date;
                var newDay = !previousCurrentDay.Equals(messageDay);
                previousCurrentDay = messageDay;

                var sg_message = ConvertOneMessageToShareGate(message, pictureCacheIds, attachmentsPath, accessToken);

                if (newDay)
                {
                    currentDayList = new List<SgMessage>();
                    retList.Add(currentDayList);
                }
                currentDayList.Add(await sg_message);
            }
            return retList;
        }

        private async Task<SgMessage> ConvertOneMessageToShareGate(Message message, Dictionary<string, string> pictureCache, string attachmentsPath, TokenHolder accessToken)
        {
            var sg_message = new SgMessage();
            sg_message.Subject = message.Subject != null ? message.Subject : "";

            sg_message.Body = new SgMessageBody();
            sg_message.Body.ContentType = "html";

            var attachmentCode = "";

            sg_message.Attachments = message.Attachments.Select(attachment =>
            {
                var sg_attachment = new SgAttachment();
                sg_attachment.Id = attachment.Id;
                sg_attachment.Content = "querying attachment not implemented";
                sg_attachment.ContentType = attachment.ContentType;
                if (attachment.ContentType.Equals("reference"))
                {
                    sg_attachment.Content = attachment.AdditionalData["contentUrl"].ToString();
                    attachmentCode += $"<div>Attachment: <a href=\"{attachment.AdditionalData["contentUrl"].ToString()}\">${attachment.Name} - {attachment.AdditionalData["contentUrl"].ToString()}</a></div>";
                }
                if (attachment.ContentType.Equals("application/vnd.microsoft.card.adaptive"))
                {
                    sg_attachment.Content = "adaptive card not migrated";
                    //                    dynamic adaptiveCard = JsonConvert.DeserializeObject(attachment.AdditionalData["content"].ToString());
                    //                    attachmentCode += $"<div>Attachment: <img src=\"{adaptiveCard.url.ToString()}\" alt=\"{adaptiveCard.altText.ToString()}\" /> </div>";
                }
                sg_attachment.Name = attachment.Name;
                sg_attachment.ExportedAttachmentContentUrl = "";
                return sg_attachment;
            }).ToList();

            sg_message.Body.Content = await TransformToShareGateMessageBodyContent(message, pictureCache, attachmentsPath, accessToken, attachmentCode);

            sg_message.Mentions = new List<object>();
            sg_message.Importance = message.Importance.ToString().ToLower();
            return sg_message;
        }

        private async Task<string> TransformToShareGateMessageBodyContent(Message message, Dictionary<string, string> pictureCache, string attachmentsPath, TokenHolder accessToken, string attachmentCode) {
            var sender = (Newtonsoft.Json.Linq.JObject)message.From.AdditionalData["user"]; // null for bots
            if (sender != null && !pictureCache.Keys.Contains(sender["id"].ToString()))
            {
                pictureCache.Add(sender["id"].ToString(), await TryToFetchPicture(sender["id"].ToString(), attachmentsPath, accessToken));
            }

            string pictureFilename = sender != null ? pictureCache[sender["id"].ToString()] : null;
            var pictureCode = pictureFilename != null ? $"<img src=\"Messages Attachments/{pictureFilename}\" width=\"32\" height=\"32\" style=\"vertical-align:top; width:32px; height:32px;\">" : "";
            string messageBody;
            if (message.Body.ContentType.ToString().Equals("Html"))
            {
                messageBody = await FetchImages(message.Body.Content, attachmentsPath, message.Id, accessToken);;
            }
            else
            {
                messageBody = message.Body.Content;
            }
            string senderDisplayName = sender != null ? sender["displayName"].ToString() : "exporter encountered unsupported sender (probably a bot?)";
            return "<div style=\"display: flex; margin-top: 10px\">"
                        + "<div style=\"flex: none; overflow: hidden; border-radius: 50%; height: 32px; width: 32px; margin: 0 10px 10px 0\">"
                            + pictureCode
                        + "</div>"
                        + "<div style=\"flex: 1; overflow: hidden;\">"
                            + "<div style=\"font-size:1.2rem; white-space:nowrap; text-overflow:ellipsis; overflow: hidden;\">"
                            + $"<span style=\"font-weight:700;\">{senderDisplayName}</span>"
                            + $"<span style=\"margin-left:1rem;\">{message.CreatedDateTime}</span>"
                            + "</div>"
                            + $"<div>{messageBody}</div>"
                        + "</div>"
                  + "</div>"
                + attachmentCode;
        }

        private async Task<string> FetchImages(string content, string attachmentsPath, string messageId, TokenHolder accessToken)
        {
            Regex graphFiles = new Regex(@"https:\/\/graph\.microsoft\.com\/[^\""]+\/\$value");
            MatchCollection urlsToFetch = graphFiles.Matches(content);
            string transformedContent = content;
            for (int count = 0; count < urlsToFetch.Count; count++)
            {
                var graphUrl = urlsToFetch[count].Value;
                var attachmentFileNamePrefix = messageId + "_" + count;
                var inlineMessageAttachmentFilename = await FetchAttachmentWithToken(graphUrl.Replace("/$value", ""), attachmentsPath, attachmentFileNamePrefix, accessToken);

                transformedContent = transformedContent.Replace(graphUrl, $"Messages Attachments/{inlineMessageAttachmentFilename}");
            }
            return transformedContent;
        }

        private Task<string> TryToFetchPicture(string v, string attachmentsPath, TokenHolder accessToken)
        {
            return FetchAttachmentWithToken($"https://graph.microsoft.com/beta/users/{v}/photos/48x48", attachmentsPath, v, accessToken);
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, TokenHolder token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                int backOffMultiplier = 1;
            begin:
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.getToken());
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                if (response.IsSuccessStatusCode)
                {
                    return content;
                }
                else
                {
                    if (response.StatusCode.Equals(System.Net.HttpStatusCode.NotFound))
                    {
                        // resource probably belongs to another tenant
                        return content;
                    }
                    if (response.StatusCode.Equals(System.Net.HttpStatusCode.Forbidden))
                    {
                        // client id lacks permissions
                        LogText.Text = "ClientId lacks permission to fetch: " + url;
                        return content;
                    }
                    if (response.StatusCode.Equals(System.Net.HttpStatusCode.Unauthorized))
                    {
                        LogText.Text = "Token expired?? Not authorized to fetch: " + url;
                        token.refreshToken();
                        goto begin;
                    }
                    else
                    {
                       LogText.Text = $"Error statuscode: {response.StatusCode} content: {response.Content}" + Environment.NewLine
                                + $"sleeping for {30*backOffMultiplier}sec..";
                        Thread.Sleep(backOffMultiplier * 30 * 1000);
                        LogText.Text = "Continuing...";
                        backOffMultiplier *= 2;
                        goto begin;
                    }
                }
            }
            catch (Exception ex)
            {
                LogText.Text = ex.ToString();
                return ex.ToString();
            }
        }

        public static string getExtensionForMimeType(string mimeType)
        {
            switch(mimeType)
            {
                case "image/jpeg":
                    return "jpeg";
                case "image/png":
                    return "png";
                case "image/gif":
                    return "gif";
                default:
                    var slashIndex = mimeType.IndexOf("/");
                    return mimeType.Substring(slashIndex + 1);
            }
        }

        public async Task<string> FetchAttachmentWithToken(string url, string attachmentsPath, string filenamePrefix, TokenHolder token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            try
            {
                var fileMetadataStr = await GetHttpContentWithToken(url, token);
                Metadata obj = null;
                if (fileMetadataStr != null)
                {
                    obj = JsonConvert.DeserializeObject<Metadata>(fileMetadataStr);
                }
                int backOffMultiplier = 1;
            begin:
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url + "/$value");
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.getToken());
                var response = httpClient.SendAsync(request);
                if (response.Result.IsSuccessStatusCode)
                {
                    var content = response.Result.Content.ReadAsByteArrayAsync().Result;
                    string filenameSuffix = GetFilenameSuffix(response.Result.Content.Headers.ContentType.MediaType, obj);

                    System.IO.File.WriteAllBytes(Path.Combine(attachmentsPath, $"{filenamePrefix}{filenameSuffix}"), content);

                    return $"{filenamePrefix}{filenameSuffix}";
                }
                else
                {
                    if (response.Result.StatusCode.Equals(System.Net.HttpStatusCode.NotFound))
                    {
                        return null;
                    }
                    if (fileMetadataStr.Contains("ErrorInsufficientPermissionsInAccessToken"))
                    {
                        LogText.Text = "ClientId lacks permission to fetch: " + url;
                        return null;
                    }
                    if (response.Result.StatusCode.Equals(System.Net.HttpStatusCode.Unauthorized))
                    {
                        token.refreshToken();
                        goto begin;
                    }
                    LogText.Text = $"Error statuscode: {response.Result.StatusCode} content: {response.Result.Content}" + Environment.NewLine
                            + $"sleeping for {30*backOffMultiplier}sec..";
                    Thread.Sleep(backOffMultiplier * 30 * 1000);
                    backOffMultiplier *= 2;
                    goto begin;
                }
                return null; // no picture found
            }
            catch (Exception ex)
            {
                LogText.Text = ex.ToString();
                return null; // probably no profile picture was found
            }
        }

        private static string GetFilenameSuffix(string contentType, Metadata contentType2)
        {
            string filenameSuffix;
            if (contentType != null && !contentType.Equals(""))
            {
                filenameSuffix = "." + getExtensionForMimeType(contentType);
            }
            else
            {
                if (contentType2 != null && !contentType2.ContentType.Equals(""))
                {
                    filenameSuffix = "." + getExtensionForMimeType(contentType2.ContentType);
                }
                else
                {
                    filenameSuffix = "";
                }
            }
            return filenameSuffix;
        }

        public async Task<Items<T>> LoadItems<T>(string url, TokenHolder token) where T : new()
        {
            int backOffMultiplier = 1;
        begin:
            var str = await GetHttpContentWithToken(url,token);
            var obj = JsonConvert.DeserializeObject<Items<T>>(str);
            if (obj.Error!=null)
            {
                if (obj.Error.Code== "TooManyRequests")
                {
                    LogText.Text = "too many requests to server," + Environment.NewLine
                        + $"sleeping for {30*backOffMultiplier}sec..";
                    Thread.Sleep(backOffMultiplier * 30 * 1000);
                    backOffMultiplier *= 2;
                    goto begin;
                }
            }
            return obj;
        }

        public async Task<T> LoadItem<T>(string url, TokenHolder token) where T : new()
        {
            var str = await GetHttpContentWithToken(url, token);
            var obj = JsonConvert.DeserializeObject<T>(str);
            return obj;
        }


        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    this.LogText.Text = "User has signed-out";
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {
                    LogText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }
    }
}

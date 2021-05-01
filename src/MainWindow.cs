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
//using System.Windows.Interop;
using Gtk;
using System.Text.RegularExpressions;

namespace MSTeamsHistory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Gtk.Window
    {
        TextView HistoryText = new TextView();
        TextView LogText = new TextView();
        Button CallGraphButton = new Button("Export `Microsoft Teams` history");
        Button SignOutButton = new Button("Sign-Out");

        //string path = $@"C:\Users\{Environment.UserName}\Desktop\MSTeamsHistory";
        string path = $@"/tmp/MSTeamsHistory";

        //Set the scope for API call to user.read
        string[] scopes = new string[] { 
            "user.read", 
            "Chat.Read"
        };


        public MainWindow() : base(Gtk.WindowType.Toplevel)
        {
            //InitializeComponent();
            //Build();
            HistoryText.Buffer.Text = path;
            CallGraphButton.Clicked += new EventHandler(CallGraphButton_Click);
            SignOutButton.Clicked += new EventHandler(SignOutButton_Click);

            VBox vbox = new VBox(true, 0);

            vbox.Add(HistoryText);
            vbox.Add(LogText);
            vbox.Add(CallGraphButton);
            vbox.Add(SignOutButton);

            this.Add(vbox);

            this.ShowAll();
        }

        protected void OnDeleteEvent(object sender, DeleteEventArgs a)
        {
           Gtk.Application.Quit();
            a.RetVal = true;
        }

        /// <summary>
        /// Call AcquireToken - to acquire a token requiring user to sign-in
        /// </summary>
        //private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        private async void CallGraphButton_Click(object sender, EventArgs e)
        {
            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            LogText.Buffer.Text = string.Empty;
            LogText.Editable = false;
            if (string.IsNullOrEmpty(HistoryText.Buffer.Text))
            {
                HistoryText.Buffer.Text = path;
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

                        //.WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                        .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    LogText.Buffer.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                LogText.Buffer.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {
                this.SignOutButton.Visible = true;
                LogText.Buffer.Text = "Loading data ...";

                var me = await LoadItem<User>("https://graph.microsoft.com/beta/me", authResult.AccessToken);

                var path = HistoryText.Buffer.Text;
                var dbPath = System.IO.Path.Combine(path, authResult.Account.Username); ;
                if (!System.IO.Directory.Exists(dbPath))
                {
                    System.IO.Directory.CreateDirectory(dbPath);
                }

                var chatsObj = await LoadItems<Models.Graph.Chats.Chat>("https://graph.microsoft.com/beta/me/chats", authResult.AccessToken);

                System.IO.File.WriteAllText(System.IO.Path.Combine(dbPath, "chats.json"), JsonConvert.SerializeObject(chatsObj.Value));
                var chatsPath = System.IO.Path.Combine(dbPath, "chats"); ;
                if (!System.IO.Directory.Exists(chatsPath))
                {
                    System.IO.Directory.CreateDirectory(chatsPath);
                }


                int i = 0;
                foreach (var chat in chatsObj.Value)
                {
                    LogText.Buffer.Text = $"Loading messages for chat {++i}/{chatsObj.Value.Count}";
                    var chatDirPath = System.IO.Path.Combine(chatsPath, chat.Id.SHA1());
                    if (!System.IO.Directory.Exists(chatDirPath))
                    {
                        System.IO.Directory.CreateDirectory(chatDirPath);
                    }

                    System.IO.File.WriteAllText(System.IO.Path.Combine(chatDirPath, "chat.json"), JsonConvert.SerializeObject(chat));

                    var messagesPath = System.IO.Path.Combine(chatDirPath, "messages.json");

                    var listMessages = new List<Message>();
                    var messages = await LoadItems<Message>($"https://graph.microsoft.com/beta/me/chats/{chat.Id}/messages", authResult.AccessToken);
                    if (messages.OdataCount > 0)
                    {
                        listMessages.AddRange(messages.Value);
                        do
                        {
                            messages = await LoadItems<Message>(messages.OdataNextLink.ToString(), authResult.AccessToken);
                            if (messages.OdataCount==0)
                            {
                                break;
                            }
                            listMessages.AddRange(messages.Value);
                        }while(true);
                    }

                    var x_messages = listMessages.OrderBy(x => x.CreatedDateTime).ToList();
                    System.IO.File.WriteAllText(messagesPath, JsonConvert.SerializeObject(x_messages));

                    var members = await LoadItems<Member>($"https://graph.microsoft.com/beta/me/chats/{chat.Id}/members", authResult.AccessToken);

                    if (members.Value!=null&&
                        members.Value.Count>0&& x_messages.Count>0)
                    {
                        System.IO.File.WriteAllText(System.IO.Path.Combine(chatDirPath, "members.json"), JsonConvert.SerializeObject(members.Value));
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
                            System.IO.File.WriteAllLines(System.IO.Path.Combine(chatDirPath, $"{history}.txt"), data);
                    }
                }
                LogText.Buffer.Text = "Done.";
            }
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        public async Task<Items<T>> LoadItems<T>(string url, string token) where T : new()
        {
            begin:
            var str = await GetHttpContentWithToken(url,token);
            var obj = JsonConvert.DeserializeObject<Items<T>>(str);
            if (obj.Error!=null)
            {
                if (obj.Error.Code== "TooManyRequests")
                {
                    LogText.Buffer.Text = "too many requests to server," + Environment.NewLine
                        + "sleeping for the 30sec..";
                    Thread.Sleep(30 * 1000);
                    goto begin;
                }
            }
            return obj;
        }

        public async Task<T> LoadItem<T>(string url, string token) where T : new()
        {
            var str = await GetHttpContentWithToken(url, token);
            var obj = JsonConvert.DeserializeObject<T>(str);
            return obj;
        }


    /// <summary>
    /// Sign out the current user
    /// </summary>
    //private async void SignOutButton_Click(object sender, RoutedEventArgs e)
    private async void SignOutButton_Click(object sender, EventArgs e)
    {
        var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    this.LogText.Buffer.Text = "User has signed-out";
                    this.CallGraphButton.Visible = true;
                    this.SignOutButton.Visible = false;
                }
                catch (MsalException ex)
                {
                    LogText.Buffer.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }
    }
}

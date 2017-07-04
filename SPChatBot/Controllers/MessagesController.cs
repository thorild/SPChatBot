using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using Microsoft.Bot.Connector;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Web;
using Newtonsoft.Json.Linq;
using Microsoft.SharePoint.Client;
using System.Security;

namespace SPChatBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        /// 


        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {

            StateClient stateClient = activity.GetStateClient();
            BotData userData = await stateClient.BotState.GetUserDataAsync(activity.ChannelId, activity.From.Id);


            if (activity.Type == ActivityTypes.Message)
            {


                ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));


                int length = (activity.Text ?? string.Empty).Length;



                if (userData.GetProperty<bool>("SentGreeting") && activity.Text.ToLower() != "ja")
                {

                    userData.SetProperty<string>("ListName", activity.Text);
                    await stateClient.BotState.SetUserDataAsync(activity.ChannelId, activity.From.Id, userData);


                    await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Grymt namn!"));

                    await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Vänta så skapar jag upp den"));

                    string x = await CreateList("https://acmebiz.sharepoint.com/sites/it", activity.Text, userData.GetProperty<string>("Type"));
                    await connector.Conversations.ReplyToActivityAsync(activity.CreateReply(x));


                    await connector.Conversations.ReplyToActivityAsync(activity.CreateReply("Vill du börja om?"));

                }
                else
                {

                    if (null == null)
                    {
                        switch (activity.Text)
                        {
                            case "Uppgiftslista":
                                await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Aha coolt en uppgiftslista!"));
                                await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Ska bli!, behöver bara lite mer info"));
                                await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Vad ska listan heta?"));

                                userData.SetProperty<bool>("SentGreeting", true);
                                userData.SetProperty<string>("Type", "Task");
                                await stateClient.BotState.SetUserDataAsync(activity.ChannelId, activity.From.Id, userData);


                                break;
                            case "Customlista":
                                await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Aha tungt! En customlista!"));
                                await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Ska bli!, behöver bara lite mer info"));
                                await connector.Conversations.ReplyToActivityAsync(activity.CreateReply($"Vad ska listan heta?"));

                                userData.SetProperty<bool>("SentGreeting", true);
                                userData.SetProperty<string>("Type", "Custom");
                                await stateClient.BotState.SetUserDataAsync(activity.ChannelId, activity.From.Id, userData);

                                break;
                            default:
                                Activity reply = activity.CreateReply($"Hej, vill du ha en lista?");
                                await connector.Conversations.ReplyToActivityAsync(reply);
                                SubmitStartSiteCreationForm(activity);

                                userData.SetProperty<bool>("SentGreeting", false);
                                await stateClient.BotState.SetUserDataAsync(activity.ChannelId, activity.From.Id, userData);
                                break;

                        }
                    }
                }




                // return our reply to the user

            }
            else
            {
                HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
        }

        private async void SubmitStartSiteCreationForm(Activity activity)
        {
            ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));
            Activity replyToConversation = activity.CreateReply("Skapa lista:");
            replyToConversation.Recipient = activity.From;
            replyToConversation.Type = "message";
            replyToConversation.Attachments = new List<Microsoft.Bot.Connector.Attachment>();
            List<CardImage> cardImages = new List<CardImage>();
            cardImages.Add(new CardImage(url: "https://www.studentlegal.ucla.edu/assets/img/index_resources.png"));
            List<CardAction> cardButtons = new List<CardAction>();

            CardAction plButton = new CardAction()
            {
                Value = "Customlista",
                Type = "imBack",
                Title = "Customlista"
            };

            CardAction pl1Button = new CardAction()
            {
                Value = "Uppgiftslista",
                Type = "imBack",
                Title = "Uppgiftslista"
            };

            cardButtons.Add(pl1Button);
            cardButtons.Add(plButton);
            HeroCard plCard = new HeroCard()
            {
                Title = "Klart du ska ha en lista - SharePointlistor rockar!",
                Subtitle = "Välj typ av lista nedan:",
                Images = cardImages,
                Buttons = cardButtons
            };
            Microsoft.Bot.Connector.Attachment plAttachment = plCard.ToAttachment();
            replyToConversation.Attachments.Add(plAttachment);
            var reply = await connector.Conversations.SendToConversationAsync(replyToConversation);
        }

        private async Task<string> CreateList(string url, string title, string type)
        {
            using (var cc = GetContext(url))
            {

                List l;
                if (type == "Task")
                  l=  cc.Web.CreateList(ListTemplateType.Tasks, title, false, true, "Lists/" + title.Replace(" ", ""), false);
                else
                   l= cc.Web.CreateList(ListTemplateType.GenericList, title, false, true, "Lists/" + title.Replace(" ", ""), false);

                l.OnQuickLaunch = true;
                cc.Load(l);
                cc.Web.Update();

                cc.ExecuteQuery();




            }

            return "Klart!!!";
        }

        private Activity HandleSystemMessage(Activity message)
        {
            if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing tha the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {
            }

            return null;
        }

        public static ClientContext GetContext(string url)
        {

            string adminUrl = url;
            string userName = "*******@*******.onmicrosoft.com";

            SecureString securePassword = new SecureString();
            string psw = "************";
            foreach (char c in psw)
            {
                securePassword.AppendChar(c);
            }


            try
            {

                using (var cc = new ClientContext(adminUrl))
                {
                    cc.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                    cc.Load(cc.Web, w => w.Title);
                    cc.ExecuteQuery();
                    return cc;

                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}
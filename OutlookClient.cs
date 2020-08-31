using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace OutlookWelkinSyncFunction
{
    public class OutlookClient // TODO: pagination
    {
        private readonly OutlookConfig config;
        private readonly string token;
        private readonly GraphServiceClient graphClient;
        private readonly ILogger logger;

        public OutlookClient(OutlookConfig config, ILogger logger)
        {
            this.config = config;
            this.logger = logger;
            IConfidentialClientApplication app = 
                        ConfidentialClientApplicationBuilder
                            .Create(config.ClientId)
                            .WithClientSecret(config.ClientSecret)
                            .WithAuthority(new Uri(config.Authority))
                            .Build();
                                                    
            string[] scopes = new string[] { $"{config.ApiUrl}.default" }; 
            
            AuthenticationResult result = app.AcquireTokenForClient(scopes).ExecuteAsync().GetAwaiter().GetResult();
            this.token = result.AccessToken;

            if (string.IsNullOrEmpty(this.token))
            {
                throw new ArgumentException($"Unable to retrieve a valid token using the credentials in env");
            }
            
            this.graphClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) => {
                requestMessage
                    .Headers
                    .Authorization = new AuthenticationHeaderValue("Bearer", this.token);

                return Task.FromResult(0);
            }));
        }

        public IEnumerable<User> GetAllUsers()
        {
            // TODO: Pagination a la https://docs.microsoft.com/en-us/graph/sdks/paging?tabs=csharp
            return this.graphClient.Users.Request().GetAsync().GetAwaiter().GetResult();
        }

        public IEnumerable<Event> GetEventsForUserScheduledBetween(User user, DateTime start, DateTime end, string extensionsNamespace = null)
        {
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("startdatetime", start.ToString("o")),
                new QueryOption("enddatetime", end.ToString("o"))
            };

            ICalendarEventsCollectionRequest request = 
                        this.graphClient
                            .Users[user.UserPrincipalName]
                            .Calendar
                            .Events
                            .Request(queryOptions);

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }

            return request
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
        }

        public Event CreateOutlookEventFromWelkinEvent(User outlookUser, WelkinEvent welkinEvent, WelkinPractitioner welkinUser)
        {
            // Create and associate a new Outlook event
            Event outlookEvent = new Event
            {
                Subject = "Placeholder for appointment in Welkin",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = $"See your Welkin calendar (user {welkinUser.Email}) for details."
                },
                IsAllDay = welkinEvent.IsAllDay,
                Start = new DateTimeTimeZone
                {
                    DateTime = welkinEvent.IsAllDay 
                        ? welkinEvent.Day.Value.Date.ToString() // Midnight day of
                        : welkinEvent.Start.Value.ToString(), // Will be UTC
                    TimeZone = welkinUser.Timezone
                },
                End = new DateTimeTimeZone
                {
                    DateTime = welkinEvent.IsAllDay 
                        ? welkinEvent.Day.Value.Date.AddDays(1).ToString() // Midnight day after
                        : welkinEvent.Start.Value.ToString(), // Will be UTC
                    TimeZone = welkinUser.Timezone
                }
            };

            Event createdEvent = this.graphClient
                                        .Users[outlookUser.UserPrincipalName]
                                        .Calendar
                                        .Events
                                        .Request()
                                        .AddAsync(outlookEvent)
                                        .GetAwaiter()
                                        .GetResult();

            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            keyValuePairs[Constants.LinkedWelkinEventIdKey] = welkinEvent.Id;
            this.SetOpenExtensionPropertiesOnEvent(outlookUser, createdEvent, keyValuePairs, Constants.OutlookEventExtensionsNamespace);

            return createdEvent;
        }

        public IEnumerable<Event> GetEventsForUserUpdatedSince(User user, TimeSpan ago, string extensionsNamespace = null)
        {
            DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;
            string filter = $"lastModifiedDateTime lt {end.ToString("o")} and lastModifiedDateTime gt {start.ToString("o")}";

            ICalendarEventsCollectionRequest request = 
                        this.graphClient
                            .Users[user.UserPrincipalName]
                            .Calendar
                            .Events
                            .Request()
                            .Filter(filter);

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }
            
            return request
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
        }

        public Event GetEventForUserWithId(User user, string id, string extensionsNamespace = null)
        {
            IEventRequest request = 
                        this.graphClient
                            .Users[user.UserPrincipalName]
                            .Calendar
                            .Events[id]
                            .Request();

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }
            
            return request
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
        }

        public void SetOpenExtensionPropertiesOnEvent(User usr, Event evt, IDictionary<string, object> keyValuePairs, string extensionsNamespace)
        {
            IEventExtensionsCollectionRequest request = 
                        this.graphClient
                            .Users[usr.UserPrincipalName]
                            .Calendar
                            .Events[evt.Id]
                            .Extensions
                            .Request();
            OpenTypeExtension ext = new OpenTypeExtension();
            ext.ExtensionName = extensionsNamespace;
            ext.AdditionalData = keyValuePairs;
            string parameterString = (keyValuePairs != null) ? string.Join(", ", keyValuePairs.Select(kv => kv.Key + "=" + kv.Value).ToArray()) : "NULL";

            request.AddAsync(ext).GetAwaiter().OnCompleted(() => this.logger.LogInformation($"Successfully added an extension with values {parameterString}."));
        }
    }
}
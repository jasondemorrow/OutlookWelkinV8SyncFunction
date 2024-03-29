namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;

    public class OutlookClient
    {
        private MemoryCache internalCache = new MemoryCache(new MemoryCacheOptions()
        {
            SizeLimit = 1024
        });
        private readonly MemoryCacheEntryOptions cacheEntryOptions = 
            new MemoryCacheEntryOptions()
                .SetAbsoluteExpiration(TimeSpan.FromSeconds(180))
                .SetSize(1);
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
            
            try
            {
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
            } catch (Exception e) {
                this.logger.LogError(e, "While constructing Outlook client");
            }
        }

        public static bool IsPlaceHolderEvent(Event outlookEvent)
        {
            Extension? extensionForWelkin = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(Constants.OutlookEventExtensionsNamespace))?.FirstOrDefault();
            if (extensionForWelkin?.AdditionalData != null && extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookPlaceHolderEventKey))
            {
                return true;
            }

            return false;
        }

        private ICalendarRequestBuilder CalendarRequestBuilderFrom(Event outlookEvent, string? userPrincipal, string calendarId = null)
        {
            if (userPrincipal == null)
            {
                User? outlookUser = outlookEvent.AdditionalData[Constants.OutlookUserObjectKey] as User;
                userPrincipal = outlookUser?.UserPrincipalName;
            }
            
            return CalendarRequestBuilderFrom(userPrincipal, calendarId);
        }

        private ICalendarRequestBuilder CalendarRequestBuilderFrom(string userPrincipal, string? calendarId = null)
        {
            IUserRequestBuilder userBuilder = this.graphClient.Users[userPrincipal];
            
            if (calendarId != null)
            {
                return userBuilder.Calendars[calendarId];
            }
            else
            {
                return userBuilder.Calendar;  // Use default calendar
            }
        }

        public Event? RetrieveEventWithICalId(
            User owningUser, 
            string iCalId, 
            string? extensionsNamespace = null, 
            string? calendarId = null)
        {
            Event? found;
            if (this.internalCache.TryGetValue(iCalId, out found))
            {
                return found;
            }

            string filter = $"iCalUId eq '{iCalId}'";

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(owningUser.UserPrincipalName, calendarId)
                            .Events
                            .Request()
                            .Filter(filter);

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }
            
            found = request
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult()
                    .FirstOrDefault();
            
            if (found != null)
            {
                this.internalCache.Set(iCalId, found, this.cacheEntryOptions);
                if (found.AdditionalData != null)
                {
                    found.AdditionalData[Constants.OutlookUserObjectKey] = owningUser;
                }
            }

            return found;
        }

        public IEnumerable<Event> RetrieveEventsForUserUpdatedSince(string userPrincipal, TimeSpan ago, string extensionsNamespace = null, string calendarId = null)
        {
            DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;
            string filter = $"lastModifiedDateTime lt {end.ToString("o")} and lastModifiedDateTime gt {start.ToString("o")}";

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(userPrincipal, calendarId)
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

        public IEnumerable<Event> RetrieveEventsForUserScheduledBetween(User outlookUser, DateTimeOffset start, DateTimeOffset end, string extensionsNamespace = null, string calendarId = null)
        {
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("startdatetime", start.ToString("o")),
                new QueryOption("enddatetime", end.ToString("o"))
            };

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(outlookUser.UserPrincipalName, calendarId)
                            .Events
                            .Request(queryOptions);

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }

            IEnumerable<Event> events = request
                                        .GetAsync()
                                        .GetAwaiter()
                                        .GetResult();

            // Cache for later individual retrieval by ICalUId
            foreach (Event outlookEvent in events)
            {
                this.internalCache.Set(outlookEvent.ICalUId, outlookEvent, this.cacheEntryOptions);
                if (outlookEvent.AdditionalData != null)
                {
                    outlookEvent.AdditionalData[Constants.OutlookUserObjectKey] = outlookUser;
                }
            }

            return events;
        }

        public ISet<string> RetrieveAllDomainsInCompany()
        {
            HashSet<string> domains;
            string key = "domains";

            if (this.internalCache.TryGetValue(key, out domains))
            {
                return domains;
            }

            try {
                var page = this.graphClient.Domains.Request().GetAsync().GetAwaiter().GetResult();
                domains = page.Select(r => r.Id).ToHashSet();
            } catch(Exception e) {
                logger.LogError(e, "In RetrieveAllDomainsInCompany");
            }

            this.internalCache.Set(key, domains, this.cacheEntryOptions);
            return domains;
        }
        
        public Event? UpdateEvent(Event outlookEvent, string? userName = null, string? calendarId = null)
        {
            this.logger.LogInformation("Not updating outlook event " + outlookEvent.ICalUId);
            return CalendarRequestBuilderFrom(outlookEvent, userName, calendarId)
                .Events[outlookEvent.Id]
                .Request()
                .UpdateAsync(outlookEvent)
                .GetAwaiter()
                .GetResult();
            return null;
        }

        public void DeleteEvent(Event outlookEvent, string userName = null, string calendarId = null)
        {
            this.logger.LogInformation("Not deleting outlook event " + outlookEvent.ICalUId);
            CalendarRequestBuilderFrom(outlookEvent, userName, calendarId)
                .Events[outlookEvent.Id]
                .Request()
                .DeleteAsync()
                .GetAwaiter()
                .GetResult();
        }

        public string? LinkedWelkinEventIdFrom(Event outlookEvent)
        {
            Extension? extensionForWelkin = outlookEvent?.Extensions?.Where(e => e.Id?.EndsWith(Constants.OutlookEventExtensionsNamespace) ?? false)?.FirstOrDefault();
            if (extensionForWelkin?.AdditionalData == null || !extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookLinkedWelkinEventIdKey))
            {
                this.logger.LogInformation($"No linked Welkin event for Outlook event {outlookEvent.ICalUId}");
                return null;
            }

            string? linkedEventId = extensionForWelkin.AdditionalData[Constants.OutlookLinkedWelkinEventIdKey]?.ToString();
            if (string.IsNullOrEmpty(linkedEventId))
            {
                this.logger.LogInformation($"Null or empty linked Welkin event ID for Outlook event {outlookEvent.ICalUId}");
                return null;
            }

            return linkedEventId;
        }

        public Microsoft.Graph.Calendar? RetrieveOwningUserDefaultCalendar(Event childEvent)
        {
            if (!childEvent.AdditionalData.ContainsKey(Constants.WelkinWorkerEmailKey))
            {
                return null;
            }
            
            return CalendarRequestBuilderFrom(
                childEvent, 
                childEvent.AdditionalData[Constants.WelkinWorkerEmailKey].ToString())
                    .Request()
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
        }

        public User? RetrieveOwningUser(Event outlookEvent)
        {
            return RetrieveUser(outlookEvent.AdditionalData[Constants.WelkinWorkerEmailKey]?.ToString());
        }

        public User? RetrieveUser(string? email)
        {
            if (email == null) {
                return null;
            }

            User retrieved;
            if (internalCache.TryGetValue(email, out retrieved))
            {
                return retrieved;
            }

            IUserRequestBuilder? userRequestBuilder = this.graphClient.Users[email];
            if (userRequestBuilder == null)
            {
                return null;
            }

            try
            {
                retrieved = userRequestBuilder.Request().GetAsync().GetAwaiter().GetResult();
            }
            catch (NullReferenceException)
            {
                return null;
            }

            if (retrieved != null)
            {
                internalCache.Set(email, retrieved, this.cacheEntryOptions);
            }

            return retrieved;
        }

        public User? FindUserCorrespondingTo(WelkinUser welkinWorker)
        {
            User? retrieved;
            string key = "outlookuserfor:" + welkinWorker.Email;
            if (internalCache.TryGetValue(key, out retrieved))
            {
                return retrieved;
            }

            ISet<string> domains = this.RetrieveAllDomainsInCompany();
            ISet<string> candidateEmails = ProducePrincipalCandidates(welkinWorker, domains);
            foreach (string email in candidateEmails)
            {
                try
                {
                    retrieved = this.RetrieveUser(email);
                    if (retrieved != null)
                    {
                        internalCache.Set(key, retrieved, this.cacheEntryOptions);
                        return retrieved;
                    }
                }
                catch (ServiceException)
                {
                }
            }
            return null;
        }

        private static ISet<string> ProducePrincipalCandidates(WelkinUser user, ISet<string> domains)
        {
            HashSet<string> candidates = new HashSet<string>();
            int idxIdAt = user.UserName.IndexOf("@");
            string idAt = (idxIdAt > -1) ? user.UserName.Substring(0, idxIdAt) : null;
            int idxIdPlus = user.UserName.IndexOf("+");
            string idPlus = (idxIdPlus > -1) ? user.UserName.Substring(0, idxIdPlus) : null;
            int idxEmailAt = user.Email.IndexOf("@");
            string emailAt = (idxEmailAt > -1) ? user.Email.Substring(0, idxEmailAt) : null;
            int idxEmailPlus = user.Email.IndexOf("+");
            string emailPlus = (idxEmailPlus > -1) ? user.Email.Substring(0, idxEmailPlus) : null;

            foreach (string domain in domains)
            {
                if (!string.IsNullOrEmpty(idAt))
                {
                    candidates.Add($"{idAt}@{domain}");
                }
                if (!string.IsNullOrEmpty(idPlus))
                {
                    candidates.Add($"{idPlus}@{domain}");
                }
                if (!string.IsNullOrEmpty(emailAt))
                {
                    candidates.Add($"{emailAt}@{domain}");
                }
                if (!string.IsNullOrEmpty(emailPlus))
                {
                    candidates.Add($"{emailPlus}@{domain}");
                }
            }

            return candidates;
        }

        public void SetOpenExtensionPropertiesOnEvent(Event? outlookEvent, IDictionary<string, object> keyValuePairs, string extensionsNamespace, string calendarId = null)
        {
            this.logger.LogInformation("Not setting extension properties outlook event " + outlookEvent.ICalUId);
            IEventExtensionsCollectionRequest request = 
                        CalendarRequestBuilderFrom(outlookEvent, null, calendarId)
                            .Events[outlookEvent.Id]
                            .Extensions
                            .Request();
            OpenTypeExtension ext = new OpenTypeExtension();
            ext.ExtensionName = extensionsNamespace;
            ext.AdditionalData = keyValuePairs;
            string parameterString = (keyValuePairs != null) ? string.Join(", ", keyValuePairs.Select(kv => kv.Key + "=" + kv.Value).ToArray()) : "NULL";

            request.AddAsync(ext).GetAwaiter().OnCompleted(() => this.logger.LogInformation($"Successfully added an extension with values {parameterString}."));
        }

        public void MergeOpenExtensionPropertiesOnEvent(Event outlookEvent, IDictionary<string, object> keyValuePairs, string extensionsNamespace)
        {
            Extension? extension = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(extensionsNamespace))?.FirstOrDefault();
            if (extension?.AdditionalData != null)
            {
                extension.AdditionalData.ToList().ForEach(x => 
                {
                    if (!keyValuePairs.ContainsKey(x.Key))
                    {
                        keyValuePairs[x.Key] = x.Value;
                    }
                });
            }

            this.SetOpenExtensionPropertiesOnEvent(outlookEvent, keyValuePairs, extensionsNamespace);
        }

        public bool SetLastSyncDateTime(Event evt, DateTimeOffset? lastSync = null)
        {
            if (lastSync == null)
            {
                lastSync = DateTimeOffset.UtcNow.AddSeconds(Constants.SecondsToAccountForEventualConsistency);
            }

            IDictionary<string, object> keyValuePairs = new Dictionary<string, object>
            {
                {Constants.OutlookLastSyncDateTimeKey , lastSync.Value.ToString("o", CultureInfo.InvariantCulture)}
            };

            try
            {
                this.MergeOpenExtensionPropertiesOnEvent(evt, keyValuePairs, Constants.OutlookEventExtensionsNamespace);
            }
            catch (Exception e)
            {
                this.logger.LogError(string.Format("While setting sync date-time for event {0}", evt.ICalUId), e);
                return false;
            }

            return true;
        }

        public static DateTime? GetLastSyncDateTime(Event outlookEvent)
        {
            Extension? extensionForWelkin = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(Constants.OutlookEventExtensionsNamespace))?.FirstOrDefault();
            if (extensionForWelkin?.AdditionalData != null && extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookLastSyncDateTimeKey))
            {
                string? lastSync = extensionForWelkin.AdditionalData[Constants.OutlookLastSyncDateTimeKey].ToString();
                return string.IsNullOrEmpty(lastSync) ? null : new DateTime?(DateTime.ParseExact(lastSync, "o", CultureInfo.InvariantCulture).ToUniversalTime());
            }

            return null;
        }

        public Event CreateOutlookEventFromWelkinEvent(WelkinEvent welkinEvent, WelkinUser welkinUser, WelkinPatient welkinPatient, string calendarId = null)
        {
            User? outlookUser = this.FindUserCorrespondingTo(welkinUser);
            if (outlookUser == null)
            {
                this.logger.LogWarning($"Couldn't find Outlook user corresponding to Welkin user {welkinUser.Email}. " +
                                       $"Can't create an Outlook event from Welkin event {welkinEvent.Id}.");
                return null;
            }
            return this.CreateOutlookEventFromWelkinEvent(welkinEvent, welkinUser, outlookUser, welkinPatient, calendarId);
        }

        public Event CreateOutlookEventFromWelkinEvent(WelkinEvent welkinEvent, WelkinUser welkinUser, User outlookUser, WelkinPatient welkinPatient, string calendarId = null)
        {
            // Create and associate a new Outlook event
            Event outlookEvent = new Event
            {
                Subject = $"Welkin Appointment: {welkinEvent.EventMode} with {welkinPatient.FirstName} {welkinPatient.LastName} for {welkinUser.FirstName} {welkinUser.LastName}",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = $"Event synchronized from Welkin. See Welkin calendar (user {welkinUser.Email}) for details."
                },
                IsAllDay = welkinEvent.IsAllDay,
                Start = new DateTimeTimeZone
                {
                    DateTime = welkinEvent.IsAllDay 
                        ? welkinEvent.Start.Value.Date.ToString() // Midnight day of
                        : welkinEvent.Start.Value.ToString(), // Will be UTC
                    TimeZone = "UTC"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = welkinEvent.IsAllDay 
                        ? welkinEvent.Start.Value.Date.AddDays(1).ToString() // Midnight day after
                        : welkinEvent.End.Value.ToString(), // Will be UTC
                    TimeZone = "UTC"
                },
                Organizer = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Name = welkinUser.FirstName + " " + welkinUser.LastName,
                        Address = welkinUser.Email
                    }
                }
            };

            this.logger.LogInformation("Not creating new outlook event " + outlookEvent.Subject);
            ICalendarRequestBuilder calendarRequestBuilder = CalendarRequestBuilderFrom(outlookUser.UserPrincipalName, calendarId);
            ICalendarEventsCollectionRequest eventsCollectionRequest = calendarRequestBuilder.Events.Request();
            Task<Event> eventResult = eventsCollectionRequest.AddAsync(outlookEvent);
            Event createdEvent = eventResult.GetAwaiter().GetResult();
            createdEvent.AdditionalData[Constants.OutlookUserObjectKey] = outlookUser;

            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            keyValuePairs[Constants.OutlookLinkedWelkinEventIdKey] = welkinEvent.Id;
            keyValuePairs[Constants.OutlookPlaceHolderEventKey] = true;
            this.SetOpenExtensionPropertiesOnEvent(createdEvent, keyValuePairs, Constants.OutlookEventExtensionsNamespace);

            return createdEvent;
            
        }

        public Microsoft.Graph.Calendar? RetrieveCalendar(string userPrincipal, string calendarName)
        {
            List<Microsoft.Graph.Calendar> calendars = new List<Microsoft.Graph.Calendar>();
            IUserCalendarsCollectionPage page = this.graphClient
                .Users[userPrincipal]
                .Calendars
                .Request()
                .GetAsync()
                .GetAwaiter().GetResult();
            calendars.AddRange(page.ToList());
            while (page.NextPageRequest != null)
            {
                page = page.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                calendars.AddRange(page.ToList());
            }
            Microsoft.Graph.Calendar? calendar = calendars
                                                    .Where(c => c.Name.ToLowerInvariant().Equals(calendarName.ToLowerInvariant()))
                                                    .FirstOrDefault();
            if (calendar == null)
            {
                string foundCalendarNames = string.Join(',', calendars.Select(c => c.Name));
                this.logger.LogWarning(
                    $"Couldn't find calendar named {calendarName} but did find the following calendars for user {userPrincipal}: {foundCalendarNames}.");
            }

            return calendar;
        }
    }
}
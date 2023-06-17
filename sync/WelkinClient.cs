namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Ninject;
    using RestSharp;

    public class WelkinClient
    {
        private MemoryCache internalCache = new MemoryCache(new MemoryCacheOptions()
        {
            SizeLimit = 1024
        });
        private readonly MemoryCacheEntryOptions cacheEntryOptions = 
            new MemoryCacheEntryOptions()
                .SetAbsoluteExpiration(TimeSpan.FromSeconds(180))
                .SetSize(1);
        private readonly WelkinConfig config;
        private readonly ILogger logger;
        private readonly string token;
        private readonly string dummyPatientId;
        private readonly string baseEndpointUrl;
        private readonly string adminEndpointUrl;

        public WelkinClient(
            WelkinConfig config, 
            ILogger logger, 
            [Named(Constants.DummyPatientEnvVarName)] string dummyPatientId, 
            [Named(Constants.WelkinUseSandboxKey)] bool useSandbox,
            [Named(Constants.WelkinTenantNameKey)] string tenantName,
            [Named(Constants.WelkinInstanceNameKey)] string instanceName)
        {
            this.config = config;
            this.logger = logger;
            this.dummyPatientId = dummyPatientId;
            string baseUrl = useSandbox ? "https://api.sandbox.welkincloud.io" : "https://api.live.welkincloud.io";
            this.adminEndpointUrl = $"{baseUrl}/{tenantName}/admin/";
            string authUrl = $"{this.adminEndpointUrl}api_clients/{this.config.ClientId}";
            this.baseEndpointUrl = $"{baseUrl}/{tenantName}/{instanceName}/";
            
            Dictionary<string, string> values = new Dictionary<string, string> 
            {
                { "secret", config.ClientSecret }
            };
            string json = JsonConvert.SerializeObject(values);
            StringContent data = new StringContent(json, Encoding.UTF8, "application/json");

            using (var httpClient = new HttpClient())
            {
                HttpResponseMessage postResponse = httpClient.PostAsync(authUrl, data)
                    .GetAwaiter()
                    .GetResult();
                string content = postResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                dynamic resp = JObject.Parse(content);
                this.token = resp.token;
            }

            if (string.IsNullOrEmpty(this.token))
            {
                throw new ArgumentException($"Unable to retrieve a valid token using the credentials in env");
            }
        }

        private T? CreateOrUpdateObject<T>(T obj, string path, string id = null) where T : class
        {
            string url = (id == null) ? $"{this.baseEndpointUrl}{path}" : $"{this.baseEndpointUrl}{path}/{id}";
            this.logger.LogInformation("Not creating or updating object " + url);
            return null;
            var client = new RestClient(url);

            Method method = (id == null) ? Method.POST : Method.PUT;
            var request = new RestRequest(method);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            request.AddParameter("application/json", JsonConvert.SerializeObject(obj), ParameterType.RequestBody);

            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK && response.StatusCode != System.Net.HttpStatusCode.Created)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }

            JObject? data = JsonConvert.DeserializeObject(response.Content) as JObject;
            if (data != null && data.ContainsKey("data"))
            {
                data = data["data"].ToObject<JProperty>()?.Value.ToObject<JObject>();
            }
            T? updated = (data == null) ? default(T) : JsonConvert.DeserializeObject<T>(data.ToString());

            if (updated != null) {
                internalCache.Set(url, updated, cacheEntryOptions);
            }

            return updated;
        }

        private T RetrieveObject<T>(string id, string path, Dictionary<string, string> parameters = null, bool serializeTopLevelObject = false)
        {
            string url = $"{this.baseEndpointUrl}{path}/{id}";
            T retrieved = default(T);
            if (internalCache.TryGetValue(url, out retrieved))
            {
                return retrieved;
            }

            var client = new RestClient(url);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");

            if (parameters != null)
            {
                foreach (KeyValuePair<string, string> kvp in parameters)
                {
                    request.AddParameter(kvp.Key, kvp.Value);
                }
            }

            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }

            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            if (serializeTopLevelObject) 
            {
                retrieved = JsonConvert.DeserializeObject<T>(result.ToString());
            }
            else 
            {
                JProperty body = result.First.ToObject<JProperty>();
                retrieved = JsonConvert.DeserializeObject<T>(body.Value.ToString());
            }

            internalCache.Set(url, retrieved, cacheEntryOptions);
            return retrieved;
        }

        private void DeleteObject(string id, string path)
        {
            string url = $"{this.baseEndpointUrl}{path}/{id}";
            this.logger.LogInformation("Not creating or updating object " + url);
            var client = new RestClient(url);

            Method method = Method.DELETE;
            var request = new RestRequest(method);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");

            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }

            internalCache.Remove(url);
        }

        private IEnumerable<T> SearchObjects<T>(string path, Dictionary<string, string> parameters = null, bool adminEndpoint = false)
        {
            string url = adminEndpoint ? $"{this.adminEndpointUrl}{path}" : $"{this.baseEndpointUrl}{path}";

            if (typeof(IAdminEntity).IsAssignableFrom(typeof(T)))
            {
                url = $"{this.adminEndpointUrl}{path}";
            }

            string key = url + "?";
            if (parameters != null && parameters.Count > 0)
            {
                key += string.Join("&", parameters.Select(e => $"{e.Key}={e.Value}"));
            }
            IEnumerable<T> found;

            if (internalCache.TryGetValue(key, out found))
            {
                return found;
            }

            var retrieved = new List<T>();
            var client = new RestClient(url);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");

            foreach (KeyValuePair<string, string> kvp in parameters ?? Enumerable.Empty<KeyValuePair<string, string>>())
            {
                request.AddParameter(kvp.Key, kvp.Value);
            }

            var response = client.Execute(request);
            //this.logger.LogInformation($"GET {key} yields {response.Content}");

            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }

            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JArray data = null;
            //this.logger.LogInformation($"GET {key} yields {response.Content}");

            if (result.ContainsKey("data"))
            {
                data = result["data"].ToObject<JArray>();
            }
            else if (result.ContainsKey("content"))
            {
                data = result["content"].ToObject<JArray>();
            }
            else
            {
                return null;
            }

            IEnumerable<T> page = JsonConvert.DeserializeObject<IEnumerable<T>>(data.ToString());
            retrieved.AddRange(page);
            int totalPages = 1;
            int currentPage = 1;

            if (result.ContainsKey("totalPages"))
            {
                totalPages = result["totalPages"].ToObject<int>();
            }
            else if(result.ContainsKey("metaInfo"))
            {
                JObject metaInfo = result["metaInfo"].ToObject<JObject>();
                totalPages = metaInfo["totalPages"].ToObject<int>();
            }
            else
            {
                this.logger.LogWarning($"Total pages not found at {url}.");
            }

            while (currentPage < totalPages)
            {
                currentPage++;
                string nextUrl = $"{url}?page={currentPage}";
                client = new RestClient(nextUrl);
                request = new RestRequest("/");
                request.AddHeader("authorization", "Bearer " + this.token);
                request.AddHeader("cache-control", "no-cache");
                response = client.Execute(request);
                result = JsonConvert.DeserializeObject(response.Content) as JObject;
                if (result.ContainsKey("data"))
                {
                    data = result["data"].ToObject<JArray>();
                }
                else if (result.ContainsKey("content"))
                {
                    data = result["content"].ToObject<JArray>();
                }
                else
                {
                    break;
                }
                page = JsonConvert.DeserializeObject<List<T>>(data.ToString());
                retrieved.AddRange(page);
            }

            internalCache.Set(key, retrieved, cacheEntryOptions);
            return retrieved;
        }

        public WelkinEvent? CreateOrUpdateEvent(WelkinEvent evt, string id = null)
        {
            evt.LocalEnd = null;   // Welkin will return "bad request"
            evt.LocalStart = null; // if we try to set local times.
            return this.CreateOrUpdateObject(evt, Constants.V8CalendarEventResourceName, id);
        }

        public WelkinEvent RetrieveEvent(string eventId)
        {
            return this.RetrieveObject<WelkinEvent>(eventId, Constants.V8CalendarEventResourceName);
        }

        public IEnumerable<WelkinEvent> RetrieveEventsOccurring(DateTime start, DateTime end)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["from"] = start.ToFormattedString("o3");
            parameters["to"] = end.ToFormattedString("o3");
            IEnumerable<WelkinEvent> retrieved = SearchObjects<WelkinEvent>(Constants.V8CalendarEventResourceName, parameters);
            // De-dupe since Welkin might send us duplicate events
            HashSet<string> uniqueIds = new HashSet<string>();
            IList<WelkinEvent> results = new List<WelkinEvent>();
            foreach(WelkinEvent evt in retrieved)
            {
                if(IsValid(evt) && !uniqueIds.Contains(evt.Id))
                {
                    results.Add(evt);
                    uniqueIds.Add(evt.Id);
                    // Cache these for later event requests
                    string key = $"{this.baseEndpointUrl}{Constants.V8CalendarEventResourceName}/{evt.Id}";
                    internalCache.Set(key, evt, cacheEntryOptions);
                }
            }
            return results;
        }

        private bool IsValid(WelkinEvent evt)
        {
            return
                evt?.Patient != null &&
                !(evt.EventStatus != null && evt.EventStatus.Equals(Constants.WelkinEventStatusCancelled));
        }

        public void DeleteEvent(WelkinEvent welkinEvent)
        {
            this.DeleteObject(welkinEvent.Id, Constants.V8CalendarEventResourceName);
        }

        public WelkinEvent? CancelEvent(WelkinEvent welkinEvent)
        {
            welkinEvent.EventStatus = Constants.WelkinEventStatusCancelled;
            return this.CreateOrUpdateObject(welkinEvent, Constants.V8CalendarEventResourceName, welkinEvent.Id);
        }

        public WelkinPatient RetrievePatient(string patientId)
        {
            return this.RetrieveObject<WelkinPatient>(patientId, Constants.WelkinPatientResourceName, null, true);
        }

        public WelkinUser? RetrieveUser(string userId)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["ids"] = userId;
            IEnumerable<WelkinUser> userResults = this.SearchObjects<WelkinUser>(Constants.WelkinUserResourceName, parameters, true);
            return userResults.FirstOrDefault();
        }

        public IEnumerable<WelkinUser> RetrieveAllUsers()
        {
            return this.SearchObjects<WelkinUser>(Constants.WelkinUserResourceName, null, true);
        }

        public WelkinUser? FindUser(string email)
        {
            if (string.IsNullOrEmpty(email))
            {
                return null;
            }

            WelkinUser? user;

            if (internalCache.TryGetValue(email.ToLowerInvariant(), out user))
            {
                return user;
            }

            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["email"] = email;
            IEnumerable<WelkinUser> found = SearchObjects<WelkinUser>(Constants.WelkinUserResourceName, parameters);
            user = found.FirstOrDefault();
            if (user != null)
            {
                internalCache.Set(user.Email.ToLowerInvariant(), user, cacheEntryOptions);
            }
            return user;
        }

        public WelkinEvent GeneratePlaceholderEventForHost(WelkinUser host, Event outlookEvent)
        {
            WelkinEvent evt = new WelkinEvent();
            evt.HostId = host.Id;
            evt.IsAllDay = true;
            evt.EventTitle = outlookEvent.Subject;
            evt.Start = DateTime.UtcNow.Date;
            evt.EventStatus = Constants.WelkinEventStatusScheduled;
            evt.EventMode = Constants.WelkinEventModeInPerson;
            WelkinEventParticipant practitioner = new WelkinEventParticipant();
            practitioner.ParticipantId = host.Id;
            practitioner.ParticipantRole = Constants.WelkinParticipantRolePsm;
            practitioner.Attended = false;
            WelkinEventParticipant patient = new WelkinEventParticipant();
            patient.ParticipantId = this.dummyPatientId;
            patient.ParticipantRole = Constants.WelkinParticipantRolePatient;
            patient.Attended = false;
            evt.Participants = new List<WelkinEventParticipant>{ practitioner, patient };

            return evt;
        }

        public bool IsPlaceHolderEvent(WelkinEvent evt)
        {
            string? patientId = evt?.Patient?.Id;
            return !string.IsNullOrEmpty(patientId) && patientId.Equals(this.dummyPatientId);
        }
    }
}
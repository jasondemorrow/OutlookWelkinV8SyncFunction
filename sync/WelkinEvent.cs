namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    public class WelkinEvent
    {
        public bool SyncWith(Event outlookEvent)
        {
            bool keepMine = 
                (outlookEvent.LastModifiedDateTime == null) || 
                (this.UpdatedAt != null && this.UpdatedAt.Value.ToUniversalTime() > outlookEvent.LastModifiedDateTime);

            if (keepMine)
            {
                outlookEvent.IsAllDay = this.IsAllDay;
                if (this.IsAllDay)
                {
                    DateTimeOffset dayUtc = this.Start.Value.ToUniversalTime();
                    outlookEvent.Start.DateTime = dayUtc.DateTime.Date.ToString("o");
                    outlookEvent.End.DateTime = dayUtc.AddDays(1).DateTime.Date.ToString("o");
                }
                else 
                {
                    outlookEvent.Start.DateTime = this.Start.Value.ToUniversalTime().DateTime.ToString("o");
                    outlookEvent.End.DateTime = this.End.Value.ToUniversalTime().DateTime.ToString("o");
                }
                outlookEvent.Start.TimeZone = Constants.OutlookUtcTimezoneLabel;
                outlookEvent.End.TimeZone = Constants.OutlookUtcTimezoneLabel;
            }
            else
            {
                this.IsAllDay = outlookEvent.IsAllDay.HasValue? outlookEvent.IsAllDay.Value : false;
                
                if (this.IsAllDay)
                {
                    this.Start = DateTime.Parse(outlookEvent.Start.DateTime);
                    this.End = this.Start.Value.AddDays(1);
                }
                else 
                {
                    this.Start = outlookEvent.StartUtc();
                    this.End = outlookEvent.EndUtc();
                }
            }

            return !keepMine; // was changed
        }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("allDayEvent")]
        public bool IsAllDay { get; set; }

        [JsonProperty("updatedAt", NullValueHandling=NullValueHandling.Ignore)]
        public DateTimeOffset? UpdatedAt { get; set; }

        [JsonProperty("createdAt", NullValueHandling=NullValueHandling.Ignore)]
        public DateTimeOffset? CreatedAt { get; set; }

        [JsonProperty("updatedBy")]
        public string UpdatedBy { get; set; }

        [JsonProperty("createdBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("hostId")]
        public string HostId { get; set; }

        [JsonProperty("eventTitle")]
        public string EventTitle { get; set; }

        [JsonProperty("eventDescription")]
        public string EventDescription { get; set; }

        [JsonProperty("eventType")]
        public string EventType { get; set; }

        [JsonProperty("eventStatus")]
        public string EventStatus { get; set; }

        [JsonProperty("eventMode")]
        public string EventMode { get; set; }

        [JsonProperty("eventColor")]
        public string EventColor { get; set; }

        [JsonProperty("startDateTime")]
        public DateTimeOffset? Start { get; set; }

        [JsonProperty("endDateTime")]
        public DateTimeOffset? End { get; set; }

        [JsonProperty("localStartDateTime")]
        public DateTimeOffset? LocalStart { get; set; }

        [JsonProperty("localEndDateTime")]
        public DateTimeOffset? LocalEnd { get; set; }

        [JsonProperty("participants")]
        public List<WelkinEventParticipant> Participants { get; set; }

        [JsonProperty("additionalInfo")]
        public Dictionary<string, string> AdditionalInfo { get; set; }

        public WelkinEventParticipant Patient
        {
            get
            {
                return this.Participants?.Where(
                    p => 
                        !string.IsNullOrEmpty(p.ParticipantRole) && 
                        p.ParticipantRole.Equals(Constants.WelkinParticipantRolePatient, StringComparison.InvariantCultureIgnoreCase)
                    ).FirstOrDefault();
            }
        }

        public string LinkedOutlookEventId
        {
            get
            {
                if (this.AdditionalInfo == null || !this.AdditionalInfo.ContainsKey(Constants.WelkinLinkedOutlookEventIdKey))
                {
                    return null;
                }
                return this.AdditionalInfo[Constants.WelkinLinkedOutlookEventIdKey];
            }
            set
            {
                if (this.AdditionalInfo == null)
                {
                    this.AdditionalInfo = new Dictionary<string, string>();
                }
                this.AdditionalInfo[Constants.WelkinLinkedOutlookEventIdKey] = value;
            }
        }

        public DateTimeOffset? LastSyncDateTime
        {
            get
            {
                if (this.AdditionalInfo == null || !this.AdditionalInfo.ContainsKey(Constants.WelkinEventLastSyncKey))
                {
                    return null;
                }
                string dateTimeString = this.AdditionalInfo[Constants.WelkinEventLastSyncKey];
                return DateTimeOffset.ParseExact(dateTimeString, "o", CultureInfo.InvariantCulture);
            }
            set
            {
                string dateTimeString = 
                    value.HasValue ? value.Value.ToString("o", CultureInfo.InvariantCulture) : null;
                if (this.AdditionalInfo == null)
                {
                    this.AdditionalInfo = new Dictionary<string, string>();
                }
                this.AdditionalInfo[Constants.WelkinEventLastSyncKey] = dateTimeString;
            }
        }

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }

        [JsonIgnore]
        public bool IsCancelled
        {
            get
            {
                return !string.IsNullOrEmpty(this.EventStatus) &&
                       this.EventStatus.Equals(Constants.WelkinEventStatusCancelled, StringComparison.InvariantCultureIgnoreCase);
            }
        }
    }
}
namespace OutlookWelkinSync
{
    
    using System;
    using System.Collections.Generic;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    public class WelkinEventParticipant
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("participantId")]
        public string ParticipantId { get; set; }

        [JsonProperty("participantRole")]
        public string ParticipantRole { get; set; }

        [JsonProperty("participationStatus")]
        public string ParticipationStatus { get; set; }

        [JsonProperty("attended")]
        public bool Attended { get; set; }
    }
}
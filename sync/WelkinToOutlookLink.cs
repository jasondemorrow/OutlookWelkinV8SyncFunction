namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public class WelkinToOutlookLink
    {
        private readonly OutlookClient outlookClient;
        private readonly WelkinClient welkinClient;
        private readonly WelkinEvent sourceWelkinEvent;
        private readonly Event targetOutlookEvent;
        protected readonly ILogger logger;

        public WelkinToOutlookLink(OutlookClient outlookClient, WelkinClient welkinClient, WelkinEvent sourceWelkinEvent, Event targetOutlookEvent, ILogger logger)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.sourceWelkinEvent = sourceWelkinEvent;
            this.targetOutlookEvent = targetOutlookEvent;
            this.logger = logger;
        }

        /// <summary>
        /// Create a link from the source Welkin event to the destination Outlook event if not already linked.
        /// </summary>
        /// <returns>True if a new link was created, otherwise false.</returns>
        public bool CreateIfMissing()
        {
            string linkedOutlookId = this.sourceWelkinEvent.LinkedOutlookEventId;

            if (string.IsNullOrEmpty(this.sourceWelkinEvent.LinkedOutlookEventId))
            {
                this.logger.LogInformation($"Linking Welkin event {this.sourceWelkinEvent.Id} to Outlook event {this.targetOutlookEvent.ICalUId}.");
                this.sourceWelkinEvent.LinkedOutlookEventId = this.targetOutlookEvent.ICalUId;
                WelkinEvent savedEvent = this.welkinClient.CreateOrUpdateEvent(this.sourceWelkinEvent, this.sourceWelkinEvent.Id);
                string outlookICalId = savedEvent.LinkedOutlookEventId;

                if (outlookICalId != null && outlookICalId.Equals(this.targetOutlookEvent.ICalUId))
                {
                    this.logger.LogInformation($"Created link from Welkin event {this.sourceWelkinEvent.Id} to Outlook event {this.targetOutlookEvent.ICalUId}");
                    return true;
                }
            }

            return false;
        }
    }
}
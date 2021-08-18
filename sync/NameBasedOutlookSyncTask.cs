namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Ninject;

    /// <summary>
    /// For the outlook event given, look for a linked welkin event and sync if it exists. 
    /// If not, get user that created the outlook event. If they have a welkin user with 
    /// the same user name, create a new, corresponding event in that welkin user's 
    /// schedule and link it with the outlook event.
    /// </summary>
    public class NameBasedOutlookSyncTask : OutlookSyncTask
    {
        private static readonly IList<string> whiteListedOutlookUserEmails = Whitelisted.Emails(Constants.OutlookUserWhitelistedEmailsKey);

        public NameBasedOutlookSyncTask(
            Event outlookEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger,
            [Named(Constants.OutlookUserWhitelistedEmailsKey)] IList<string> whiteListedOutlookUserEmails)
        : base(outlookEvent, outlookClient, welkinClient, logger)
        {
        }

        public override WelkinEvent Sync()
        {
            if (!this.ShouldSync())
            {
                return null;
            }

            string organizerEmail = this.outlookEvent.Organizer?.EmailAddress?.Address?.ToLowerInvariant().Trim();
            string linkedWelkinEventId = this.outlookClient.LinkedWelkinEventIdFrom(this.outlookEvent);
            WelkinEvent syncedTo = null;

            if (whiteListedOutlookUserEmails != null && whiteListedOutlookUserEmails.Count > 0 && !whiteListedOutlookUserEmails.Contains(organizerEmail))
            {
                this.logger.LogWarning($"Skipping sync of Outlook event {this.outlookEvent.ICalUId} for user {organizerEmail} since they are not whitelisted for sync.");
                return null; // There's a whitelist, and this user isn't on it.
            }


            if (!string.IsNullOrEmpty(linkedWelkinEventId))
            {
                syncedTo = this.welkinClient.RetrieveEvent(linkedWelkinEventId);
                if (syncedTo.SyncWith(this.outlookEvent)) // Welkin needs to be updated
                {
                    syncedTo = this.welkinClient.CreateOrUpdateEvent(syncedTo, syncedTo.Id);
                }
                else // Outlook needs to be updated
                {
                    this.outlookClient.UpdateEvent(this.outlookEvent);
                }
            }
            else // Welkin needs to be created
            {
                // Find the Welkin user for the Outlook event owner
                string eventOwnerEmail = this.outlookEvent.AdditionalData[Constants.WelkinWorkerEmailKey].ToString().ToLowerInvariant().Trim();
                WelkinUser practitioner = this.welkinClient.FindUser(eventOwnerEmail);
                Throw.IfAnyAreNull(eventOwnerEmail, practitioner);

                // Generate and save a placeholder event in Welkin with a dummy patient
                WelkinEvent placeholderEvent = this.welkinClient.GeneratePlaceholderEventForHost(practitioner, this.outlookEvent);
                placeholderEvent.SyncWith(this.outlookEvent);
                placeholderEvent = this.welkinClient.CreateOrUpdateEvent(placeholderEvent, placeholderEvent.Id);

                // Link the Outlook and Welkin events using external metadata fields
                OutlookToWelkinLink outlookToWelkinLink = new OutlookToWelkinLink(
                    this.outlookClient, this.welkinClient, this.outlookEvent, placeholderEvent, this.logger);

                if (outlookToWelkinLink.CreateIfMissing())
                {
                    // Link did not previously exist and needs to be created from Welkin to Outlook as well
                    WelkinToOutlookLink welkinToOutlookLink = new WelkinToOutlookLink(
                        this.outlookClient, this.welkinClient, placeholderEvent, this.outlookEvent, this.logger);
                    
                    if (!welkinToOutlookLink.CreateIfMissing())
                    {
                        outlookToWelkinLink.Rollback();
                        throw new LinkException(
                            $"Failed to create link from Welkin event {placeholderEvent.Id} " +
                            $"to Outlook event {this.outlookEvent.ICalUId}.");
                    }

                    syncedTo = placeholderEvent;
                }
            }

            this.outlookClient.SetLastSyncDateTime(this.outlookEvent);
            return syncedTo;
        }
    }
}
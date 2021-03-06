namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Ninject;

    /// <summary>
    /// For the welkin event given, look for a linked outlook event and sync if it exists. 
    /// If not, get user that created the welkin event. If they have an outlook user with 
    /// the same user name, create a new, corresponding event in that outlook user's 
    /// calendar and link it with the welkin event.
    /// </summary>
    public class NameBasedWelkinSyncTask : WelkinSyncTask
    {
        private static readonly IList<string> whiteListedWelkinUserEmails = Whitelisted.Emails(Constants.WelkinUserWhitelistedEmailsKey);

        public NameBasedWelkinSyncTask(
            WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger,
            [Named(Constants.WelkinUserWhitelistedEmailsKey)] IList<string> whiteListedWelkinUserEmails)
        : base(welkinEvent, outlookClient, welkinClient, logger)
        {
        }

        public override Event Sync()
        {
            if (!this.ShouldSync())
            {
                return null;
            }

            WelkinUser practitioner = this.welkinClient.RetrieveUser(welkinEvent.HostId);
            string welkinEmail = practitioner.Email?.ToLowerInvariant().Trim();

            if (whiteListedWelkinUserEmails != null && whiteListedWelkinUserEmails.Count > 0 && !whiteListedWelkinUserEmails.Contains(welkinEmail))
            {
                this.logger.LogWarning($"Skipping sync of Welkin event {this.welkinEvent.Id} for user {welkinEmail} since they are not whitelisted for sync.");
                return null; // There's a whitelist, and this user isn't on it.
            }

            string syncedOutlookEventId = this.welkinEvent.ExternalId;
            Event syncedTo = null;

            // If there's already an Outlook event linked to this Welkin event
            if (!string.IsNullOrEmpty(this.welkinEvent.ExternalId))
            {
                string outlookICalId = this.welkinEvent.ExternalId;
                this.logger.LogInformation($"Found Outlook event {outlookICalId} associated with Welkin event {welkinEvent.Id}.");
                User outlookUser = this.outlookClient.FindUserCorrespondingTo(practitioner);
                syncedTo = this.outlookClient.RetrieveEventWithICalId(outlookUser, outlookICalId);
                if (this.welkinEvent.SyncWith(syncedTo)) // Welkin needs to be updated
                {
                    this.welkinEvent = this.welkinClient.CreateOrUpdateEvent(this.welkinEvent, this.welkinEvent.Id);
                }
                else // Outlook needs to be updated
                {
                    this.outlookClient.UpdateEvent(syncedTo);
                }
            }
            else // An Outlook event needs to be created and linked
            {
                WelkinPatient patient = this.welkinClient.RetrievePatient(this.welkinEvent.Patient.ParticipantId);
                // This will also create and persist the Outlook->Welkin link
                syncedTo = this.outlookClient.CreateOutlookEventFromWelkinEvent(this.welkinEvent, practitioner, patient);
                this.logger.LogInformation($"Successfully created a new Outlook placeholder event {syncedTo.ICalUId} in user calendar for {practitioner.Email}.");

                if (syncedTo == null)
                {
                    throw new SyncException(
                        $"Failed to create Outlook event for Welkin event {this.welkinEvent.Id}, probably because a " +
                        $"corresponding Outlook user wasn't found for Welkin worker {practitioner.Email}");
                }
                
                WelkinToOutlookLink welkinToOutlookLink = new WelkinToOutlookLink(
                    this.outlookClient, this.welkinClient, this.welkinEvent, syncedTo, this.logger);

                if (!welkinToOutlookLink.CreateIfMissing())
                {
                    // Failed for some reason, need to roll back
                    this.outlookClient.DeleteEvent(syncedTo);
                    throw new LinkException(
                        $"Failed to create link from Welkin event {this.welkinEvent.Id} " +
                        $"to Outlook event {syncedTo.ICalUId}.");
                }
            }

            this.welkinEvent.ExternalIdUpdatedAt = DateTimeOffset.UtcNow.AddSeconds(Constants.SecondsToAccountForEventualConsistency);
            this.welkinClient.CreateOrUpdateEvent(this.welkinEvent, this.welkinEvent.Id);
            return syncedTo;
        }

        public override void Cleanup()
        {
            if (this.welkinClient.IsPlaceHolderEvent(this.welkinEvent))
            {
                WelkinUser practitioner = this.welkinClient.RetrieveUser(this.welkinEvent.HostId);
                User outlookUser = this.outlookClient.FindUserCorrespondingTo(practitioner);
                string outlookICalId = this.welkinEvent.ExternalId;
                Event outlookEvent = null;

                if (!string.IsNullOrEmpty(outlookICalId) && outlookUser != null)
                {
                    try
                    {
                        outlookEvent = this.outlookClient.RetrieveEventWithICalId(outlookUser, outlookICalId);
                    }
                    catch (ServiceException)
                    {
                        outlookEvent = null;
                    }
                }

                // If we can't find the externally mapped Outlook event for this placeholder event, clean it up
                if (!string.IsNullOrEmpty(outlookICalId) && outlookEvent == null)
                {
                    this.logger.LogWarning($"Welkin event {this.welkinEvent.Id} is an orphaned placeholder event for Outlook user " + 
                                           $"{outlookUser.UserPrincipalName} and will be cancelled. Event details: {welkinEvent.ToString()}.");
                    this.welkinClient.CancelEvent(this.welkinEvent);
                }
            }
        }
    }
}
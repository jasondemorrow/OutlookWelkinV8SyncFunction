namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Ninject;

    /// <summary>
    /// For the welkin event given, look for a linked outlook event in the configured 
    /// shared calendar (by user name and calendar name) and sync if it exists. 
    /// If no corresponding event exists in the shared calendar, create it and 
    /// and link it with the welkin event.
    /// </summary>
    public class SharedCalendarWelkinSyncTask : WelkinSyncTask
    {
        private readonly string sharedCalendarUser;
        private readonly string sharedCalendarName;
        private readonly User sharedCalendarOutlookUser;
        private readonly Calendar sharedOutlookCalendar;

        public SharedCalendarWelkinSyncTask(
            WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger,
            [Named(Constants.SharedCalUserEnvVarName)] string sharedCalendarUser,
            [Named(Constants.SharedCalNameEnvVarName)] string sharedCalendarName
            ) : base(welkinEvent, outlookClient, welkinClient, logger)
        {
            this.sharedCalendarUser = sharedCalendarUser;
            this.sharedCalendarName = sharedCalendarName;
            this.sharedCalendarOutlookUser = this.outlookClient.RetrieveUser(this.sharedCalendarUser);
            this.sharedOutlookCalendar = this.outlookClient.RetrieveCalendar(this.sharedCalendarUser, this.sharedCalendarName);
        }

        public override Event Sync()
        {
            // 1. Look for an external link to an outlook event
            // 2. If one exists, look for it in the shared calendar
            // 3. If it can be retrieved from the shared calendar, sync it
            // 4. If linked event isn't in shared calendar or no link exists, 
            //    make a new event in the shared calendar and sync it
            // 5. Create or update the external link
            if (!this.ShouldSync())
            {
                return null;
            }

            Event linkedOutlookEvent = null; // From the configured shared calendar
            string syncedOutlookEventId = this.welkinEvent.ExternalId;
            WelkinUser worker = this.welkinClient.RetrieveUser(this.welkinEvent.HostId);

            if (!string.IsNullOrEmpty(this.welkinEvent.ExternalId))
            {
                string outlookICalId = this.welkinEvent.ExternalId;
                linkedOutlookEvent = this.outlookClient.RetrieveEventWithICalId(
                    this.sharedCalendarOutlookUser, 
                    outlookICalId, 
                    Constants.OutlookEventExtensionsNamespace, 
                    this.sharedOutlookCalendar.Id);
            }

            if (linkedOutlookEvent != null)
            {
                if (this.welkinEvent.SyncWith(linkedOutlookEvent)) // Welkin needs to be updated
                {
                    this.welkinEvent = this.welkinClient.CreateOrUpdateEvent(this.welkinEvent, this.welkinEvent.Id);
                }
                else // Outlook needs to be updated
                {
                    this.outlookClient.UpdateEvent(linkedOutlookEvent);
                }
            }
            else // An Outlook event needs to be created and linked
            {
                WelkinPatient patient = this.welkinClient.RetrievePatient(this.welkinEvent.Patient.Id);
                // This will also create and persist the Outlook->Welkin link
                linkedOutlookEvent = this.outlookClient.CreateOutlookEventFromWelkinEvent(
                    this.welkinEvent, worker, this.sharedCalendarOutlookUser, patient, this.sharedOutlookCalendar.Id);
                this.logger.LogInformation($"Successfully created a new Outlook placeholder event {linkedOutlookEvent.ICalUId} in shared calendar {this.sharedOutlookCalendar.Name}.");
                WelkinToOutlookLink welkinToOutlookLink = new WelkinToOutlookLink(
                    this.outlookClient, this.welkinClient, this.welkinEvent, linkedOutlookEvent, this.logger);

                if (!welkinToOutlookLink.CreateIfMissing())
                {
                    // Failed for some reason, need to roll back
                    this.outlookClient.DeleteEvent(linkedOutlookEvent);
                    throw new LinkException(
                        $"Failed to create link from Welkin event {this.welkinEvent.Id} " +
                        $"to Outlook event {linkedOutlookEvent.ICalUId}.");
                }
            }

            this.welkinEvent.ExternalIdUpdatedAt = DateTimeOffset.UtcNow.AddSeconds(Constants.SecondsToAccountForEventualConsistency);
            this.welkinClient.CreateOrUpdateEvent(this.welkinEvent, this.welkinEvent.Id);
            return linkedOutlookEvent;
        }
        public override void Cleanup()
        {
            if (this.welkinClient.IsPlaceHolderEvent(this.welkinEvent))
            {
                User outlookUser = this.sharedCalendarOutlookUser;
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

                // If we can't find either the externally mapped Outlook event for this placeholder event, clean it up
                if (!string.IsNullOrEmpty(outlookICalId) && outlookEvent == null)
                {
                    this.logger.LogWarning($"Welkin event {this.welkinEvent.Id} is an orphaned placeholder event for Outlook user " + 
                                           $"{outlookUser.UserPrincipalName} and will be deleted. Event details: {welkinEvent.ToString()}.");
                    this.welkinClient.CancelEvent(this.welkinEvent);
                }
            }
        }
    }
}
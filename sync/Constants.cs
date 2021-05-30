using System;

namespace OutlookWelkinSync
{
    public static class Constants
    {
        public const string OutlookEventExtensionsNamespace = "sync.outlook.welkinhealth.com";
        public const string WelkinWorkerEmailKey = "sync.welkin.worker.email";
        public const string OutlookUserObjectKey = "sync.welkin.outlook.user.object";
        public const string WelkinPatientExtensionNamespace = "patient_placeholders_sync_outlook_welkinhealth_com";
        public const string WelkinEventExtensionNamespacePrefix = "sync_outlook_";
        public const string WelkinLastSyncExtensionNamespace = "sync_last_datetime";
        public const string WelkinClientVersionKey = "WelkinClientVersion";
        public const string WelkinTenantNameKey = "WelkinV8TenantName";
        public const string WelkinInstanceNameKey = "WelkinV8InstanceName";
        public const string WelkinUseSandboxKey = "WelkinV8UseSandbox";
        public const string WelkinEventLastSyncKey = "WelkinEventLastSync";
        public const string WelkinLinkedOutlookEventIdKey = "LinkedOutlookEventId";
        public const string OutlookLinkedWelkinEventIdKey = "LinkedWelkinEventId";
        public const string OutlookPlaceHolderEventKey = "IsOutlookPlaceHolderEvent";
        public const string OutlookLastSyncDateTimeKey = "LastSyncDateTime";
        public const string OutlookUtcTimezoneLabel = "UTC";
        public const string DefaultModality = "call";
        public const string DefaultAppointmentType = "intake_call";
        public const string WelkinEventStatusCancelled = "Cancelled";
        public const string WelkinEventStatusScheduled = "Scheduled";
        public const string WelkinEventModeInPerson = "IN-PERSON";
        public const string CalendarEventResourceName = "calendar_events";
        public const string V8CalendarEventResourceName = "calendar/events";
        public const string CalendarResourceName = "calendars";
        public const string WelkinCalendarResourceName = "calendar";
        public const string WelkinParticipantRolePatient = "patient";
        public const string WelkinParticipantRolePsm = "psm";
        public const string ExternalIdResourceName = "external_ids";
        public const string WelkinUserResourceName = "users";
        public const string WelkinPatientResourceName = "patients";
        public const string SyncNamespaceDateSeparator = ":::";
        public const string DummyPatientEnvVarName = "WelkinDummyPatientId";
        public const string SharedCalUserEnvVarName = "OutlookSharedCalendarUser";
        public const string SharedCalNameEnvVarName = "OutlookSharedCalendarName";
        public const int SecondsToAccountForEventualConsistency = 3;
    }
}
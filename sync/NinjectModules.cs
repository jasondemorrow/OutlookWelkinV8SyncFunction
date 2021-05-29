namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Ninject.Modules;

    public class NinjectModules
    {
        public static NinjectModule CurrentModule { get; set; } = new ProdModule();
        public static ILogger CurrentLogger { get; set; } = null;

        public class ProdModule : NinjectModule
        {
            public override void Load()
            {
                Bind<ILogger>().ToConstant(CurrentLogger);
                Bind<WelkinConfig>().To<WelkinConfig>(); // just create a default instance, don't modify
                Bind<OutlookConfig>().To<OutlookConfig>(); // just create a default instance, don't modify
                Bind<string>()
                    .ToMethod((context) => Environment.GetEnvironmentVariable(Constants.DummyPatientEnvVarName))
                    .InSingletonScope()
                    .Named("DummyPatientId");
                Bind<OutlookClient>().To<OutlookClient>().InSingletonScope();
                string sandboxMode = Environment.GetEnvironmentVariable(Constants.WelkinUseSandboxKey)?.ToLowerInvariant() ?? "false";
                bool useSandbox = Boolean.Parse(sandboxMode);
                Bind<bool>()
                    .ToConstant(useSandbox)
                    .InSingletonScope()
                    .Named(Constants.WelkinUseSandboxKey);
                Bind<string>()
                    .ToMethod((context) => Environment.GetEnvironmentVariable(Constants.WelkinTenantNameKey))
                    .InSingletonScope()
                    .Named(Constants.WelkinTenantNameKey);
                Bind<string>()
                    .ToMethod((context) => Environment.GetEnvironmentVariable(Constants.WelkinInstanceNameKey))
                    .InSingletonScope()
                    .Named(Constants.WelkinInstanceNameKey);
                Bind<WelkinClient>().To<WelkinClient>().InSingletonScope();
                Bind<OutlookSyncTask>().To<NameBasedOutlookSyncTask>();

                string sharedCalendarUser = Environment.GetEnvironmentVariable(Constants.SharedCalUserEnvVarName);
                string sharedCalendarName = Environment.GetEnvironmentVariable(Constants.SharedCalNameEnvVarName);

                if (!string.IsNullOrEmpty(sharedCalendarUser) && !string.IsNullOrEmpty(sharedCalendarName))
                {
                    Bind<string>()
                        .ToConstant(sharedCalendarUser)
                        .InSingletonScope()
                        .Named(Constants.SharedCalUserEnvVarName);
                    Bind<string>()
                        .ToConstant(sharedCalendarName)
                        .InSingletonScope()
                        .Named(Constants.SharedCalNameEnvVarName);
                    Bind<WelkinSyncTask>().To<SharedCalendarWelkinSyncTask>();
                    Bind<OutlookEventRetrieval>().To<SharedCalendarOutlookEventRetrieval>();
                }
                else
                {
                    Bind<WelkinSyncTask>().To<NameBasedWelkinSyncTask>();
                    Bind<OutlookEventRetrieval>().To<WelkinWorkerOutlookEventRetrieval>();
                }
            }
        }
    }
}
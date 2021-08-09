namespace OutlookWelkinSync
{
    using Newtonsoft.Json.Converters;

    public class WelkinUtcDateTimeConverter : IsoDateTimeConverter 
    {
        public WelkinUtcDateTimeConverter()
        {
            base.DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffZ";
        }
    }
}
namespace OutlookWelkinSync
{
    using Newtonsoft.Json.Converters;

    public class WelkinLocalDateTimeConverter : IsoDateTimeConverter 
    {
        public WelkinLocalDateTimeConverter()
        {
            base.DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffzzz";
        }
    }
}
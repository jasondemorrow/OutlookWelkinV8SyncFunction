namespace OutlookWelkinSync
{
    using System;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;
    
    public class IsoDateTimeAccuracyConverter : IsoDateTimeConverter
    {
        private readonly int accuracy;

        public IsoDateTimeAccuracyConverter(int accuracy)
        {
            this.accuracy = accuracy;
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(DateTimeOffset) || objectType == typeof(DateTimeOffset?) ||
                   objectType == typeof(DateTime) || objectType == typeof(DateTime?);
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            Type objectType = value.GetType();
            if (objectType == typeof(DateTimeOffset) || objectType == typeof(DateTimeOffset?))
            {
                var dateTimeValue =  (DateTimeOffset)value;
                base.WriteJson(writer, dateTimeValue.UtcDateTime.ToFormattedString("o" + this.accuracy), serializer);
            }
            else
            {
                var dateTimeValue = (DateTime)value;
                base.WriteJson(writer, dateTimeValue.ToFormattedString("o" + this.accuracy), serializer);
            }
        }
    }
}
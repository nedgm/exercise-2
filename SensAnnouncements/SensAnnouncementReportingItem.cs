namespace SensAnnouncements
{
    public class SensAnnouncementReportingItem
    {
        public string Title { get; set; }
        public string Url { get; set; }
        public string Sid { get; set; }
        public string Timestamp { get; set; }

        public SensAnnouncementReportingItem(string title, string url, string sidTimestamp)
        {
            Title = title;
            Url = url;
            PopulateSidAndTimestamp(sidTimestamp);
        }

        private void PopulateSidAndTimestamp(string sidTimestamp)
        {
            const char sidTimeStampDelimiter = '-';
            var delimiterIndex = sidTimestamp.IndexOf(sidTimeStampDelimiter);
            if (delimiterIndex != -1)
            {
                Sid = sidTimestamp.Substring(0, delimiterIndex).Trim();
                Timestamp = sidTimestamp.Substring(delimiterIndex + 1).Trim();
            }
        }
    }
}
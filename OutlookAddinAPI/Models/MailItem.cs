using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace OutlookAddinAPI.Models
{
    // To parse this JSON data, add NuGet 'Newtonsoft.Json' then do:
    //
    //    using QuickType;
    //
    //    var mailItem = MailItem.FromJson(jsonString);

    public partial class MailItem
    {
        [JsonProperty("_data$p$0")]
        public MailItemDataP0 DataP0 { get; set; }

        [JsonProperty("_body$p$1")]
        public object BodyP1 { get; set; }

        [JsonProperty("_notificationMessages$p$0")]
        public object NotificationMessagesP0 { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("bodytext")]
        public string Bodytext { get; set; }
    }

    public partial class MailItemDataP0
    {
        [JsonProperty("_data$p$0")]
        public DataP0__DataP0 DataP0 { get; set; }

        [JsonProperty("_permissionLevel$p$0")]
        public long PermissionLevelP0 { get; set; }
    }

    public partial class DataP0__DataP0
    {
        [JsonProperty("permissionLevel")]
        public long PermissionLevel { get; set; }

        [JsonProperty("timeZoneOffsets")]
        public TimeZoneOffset[] TimeZoneOffsets { get; set; }

        [JsonProperty("hostVersion")]
        public string HostVersion { get; set; }

        [JsonProperty("owaView")]
        public string OwaView { get; set; }

        [JsonProperty("marketplaceAssetId")]
        public object MarketplaceAssetId { get; set; }

        [JsonProperty("extensionId")]
        public string ExtensionId { get; set; }

        [JsonProperty("marketplaceContentMarket")]
        public object MarketplaceContentMarket { get; set; }

        [JsonProperty("entryPointUrl")]
        public string EntryPointUrl { get; set; }

        [JsonProperty("consentMetadata")]
        public object ConsentMetadata { get; set; }

        [JsonProperty("endNodeUrl")]
        public object EndNodeUrl { get; set; }

        [JsonProperty("isRead")]
        public bool IsRead { get; set; }

        [JsonProperty("userProfileType")]
        public string UserProfileType { get; set; }

        [JsonProperty("overrideWindowOpen")]
        public bool OverrideWindowOpen { get; set; }

        [JsonProperty("userDisplayName")]
        public string UserDisplayName { get; set; }

        [JsonProperty("userEmailAddress")]
        public string UserEmailAddress { get; set; }

        [JsonProperty("userTimeZone")]
        public string UserTimeZone { get; set; }

        [JsonProperty("ewsUrl")]
        public string EwsUrl { get; set; }

        [JsonProperty("restUrl")]
        public string RestUrl { get; set; }

        [JsonProperty("itemType")]
        public long ItemType { get; set; }

        [JsonProperty("conversationId")]
        public string ConversationId { get; set; }

        [JsonProperty("internetMessageId")]
        public string InternetMessageId { get; set; }

        [JsonProperty("from")]
        public From From { get; set; }

        [JsonProperty("sender")]
        public From Sender { get; set; }

        [JsonProperty("to")]
        public From[] To { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("itemClass")]
        public string ItemClass { get; set; }

        [JsonProperty("dateTimeCreated")]
        public long DateTimeCreated { get; set; }

        [JsonProperty("dateTimeModified")]
        public long DateTimeModified { get; set; }

        [JsonProperty("dateTimeSent")]
        public long DateTimeSent { get; set; }

        [JsonProperty("entities")]
        public Entities Entities { get; set; }

        [JsonProperty("filteredEntities")]
        public object FilteredEntities { get; set; }

        [JsonProperty("subject")]
        public string Subject { get; set; }

        [JsonProperty("normalizedSubject")]
        public string NormalizedSubject { get; set; }

        [JsonProperty("attachments")]
        public Attachment[] Attachments { get; set; }
    }

    public partial class Attachment
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("contentType")]
        public string ContentType { get; set; }

        [JsonProperty("size")]
        public long Size { get; set; }

        [JsonProperty("attachmentType")]
        public long AttachmentType { get; set; }

        [JsonProperty("isInline")]
        public bool IsInline { get; set; }
    }

    public partial class Entities
    {
        [JsonProperty("Addresses")]
        public object Addresses { get; set; }

        [JsonProperty("Urls")]
        public object Urls { get; set; }

        [JsonProperty("PhoneNumbers")]
        public object PhoneNumbers { get; set; }

        [JsonProperty("Contacts")]
        public object Contacts { get; set; }

        [JsonProperty("EmailAddresses")]
        public object EmailAddresses { get; set; }

        [JsonProperty("MeetingSuggestions")]
        public object MeetingSuggestions { get; set; }

        [JsonProperty("IsLegacyEntityExtraction")]
        public bool IsLegacyEntityExtraction { get; set; }

        [JsonProperty("TaskSuggestions")]
        public object TaskSuggestions { get; set; }
    }

    public partial class From
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("address")]
        public string Address { get; set; }

        [JsonProperty("recipientType")]
        public long RecipientType { get; set; }
    }

    public partial class TimeZoneOffset
    {
        [JsonProperty("start")]
        public long Start { get; set; }

        [JsonProperty("end")]
        public long End { get; set; }

        [JsonProperty("offset")]
        public long Offset { get; set; }
    }

    public partial class MailItem
    {
        public static MailItem FromJson(string json) => JsonConvert.DeserializeObject<MailItem>(json, Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this MailItem self) => JsonConvert.SerializeObject(self, Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters = {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }
}
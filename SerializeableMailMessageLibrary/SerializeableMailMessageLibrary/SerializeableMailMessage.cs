using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace SerializeableMailMessageLibrary
{
    ///

    /// Serializeable mailmessage object
    ///
    [Serializable]
    public class SerializeableMailMessage
    {
        public Guid MessageID { get; set; }
        public String Department { get; set; }
        public Boolean areAttachmentsEmbedded { get; set; }
        public String UserEmail { get; set; }
        Boolean IsBodyHtml { get; set; }
        String Body { get; set; }
        SerializeableMailAddress From { get; set; }
        List<SerializeableMailAddress> To = new List<SerializeableMailAddress>();
        List<SerializeableMailAddress> CC = new List<SerializeableMailAddress>();
        List<SerializeableMailAddress> Bcc = new List<SerializeableMailAddress>();
        List<SerializeableMailAddress> ReplyToList = new List<SerializeableMailAddress>();
        SerializeableMailAddress Sender { get; set; }
        String Subject { get; set; }
        List<SerializeableAttachment> Attachments = new List<SerializeableAttachment>();
        string BodyEncoding = string.Empty;
        string SubjectEncoding = string.Empty;
        DeliveryNotificationOptions DeliveryNotificationOptions;
        SerializeableCollection Headers;
        MailPriority Priority;
        List<SerializeableAlternateView> AlternateViews = new List<SerializeableAlternateView>();


        ///

        /// Creates a new serializeable mailmessage based on a MailMessage object
        ///

        /// 
        public SerializeableMailMessage(MailMessage _MailMessage, Guid _MessageID, string Department)
        {
            if (_MessageID == null || _MessageID == Guid.Empty)
                this.MessageID = Guid.NewGuid();
            else
                this.MessageID = _MessageID;

            if (string.IsNullOrEmpty(Department))
                this.Department = string.Empty;
            else
                this.Department = Department;

    
            IsBodyHtml = _MailMessage.IsBodyHtml;
            Body = _MailMessage.Body;
            Subject = _MailMessage.Subject;
            From = SerializeableMailAddress.GetSerializeableMailAddress(_MailMessage.From);
            To = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in _MailMessage.To)
            {
                To.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }

            CC = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in _MailMessage.CC)
            {
                CC.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }

            Bcc = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in _MailMessage.Bcc)
            {
                Bcc.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }

            Attachments = new List<SerializeableAttachment>();
            foreach (Attachment att in _MailMessage.Attachments)
            {
                Attachments.Add(SerializeableAttachment.GetSerializeableAttachment(att));
            }

            ReplyToList = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in _MailMessage.ReplyToList)
            {
                ReplyToList.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }
            if (_MailMessage.BodyEncoding != null)
            {
                BodyEncoding = _MailMessage.BodyEncoding.GetType().ToString();
            }

            DeliveryNotificationOptions = _MailMessage.DeliveryNotificationOptions;
            Headers = SerializeableCollection.GetSerializeableCollection(_MailMessage.Headers);
            Priority = _MailMessage.Priority;
            Sender = SerializeableMailAddress.GetSerializeableMailAddress(_MailMessage.Sender);


            if (_MailMessage.SubjectEncoding != null)
            {
                SubjectEncoding = _MailMessage.SubjectEncoding.GetType().ToString();
            }

            foreach (AlternateView av in _MailMessage.AlternateViews)
                AlternateViews.Add(SerializeableAlternateView.GetSerializeableAlternateView(av));
        }

        public MailMessage GetMailMessage()
        {
            MailMessage mm = new MailMessage();

            mm.IsBodyHtml = IsBodyHtml;
            mm.Body = Body;
            mm.Subject = Subject;
            if (From != null)
                mm.From = From.GetMailAddress();

            foreach (SerializeableMailAddress ma in To)
            {
                mm.To.Add(ma.GetMailAddress());
            }

            foreach (SerializeableMailAddress ma in CC)
            {
                mm.CC.Add(ma.GetMailAddress());
            }

            foreach (SerializeableMailAddress ma in Bcc)
            {
                mm.Bcc.Add(ma.GetMailAddress());
            }

            foreach (SerializeableMailAddress ma in ReplyToList)
            {
                mm.ReplyToList.Add(ma.GetMailAddress());
            }

            foreach (SerializeableAttachment att in Attachments)
            {
                mm.Attachments.Add(att.GetAttachment());
            }


            switch (BodyEncoding)
            {
                case "System.Text.ASCIIEncoding":
                    mm.BodyEncoding = Encoding.ASCII;
                    break;
                case "System.Text.UnicodeEncoding":
                    mm.BodyEncoding = Encoding.Unicode;
                    break;
                case "System.Text.UTF32Encoding":
                    mm.BodyEncoding = Encoding.UTF32;
                    break;
                case "System.Text.UTF7Encoding":
                    mm.BodyEncoding = Encoding.UTF7;
                    break;
                case "System.Text.UTF8Encoding":
                    mm.BodyEncoding = Encoding.UTF8;
                    break;
                default: mm.BodyEncoding = null;
                    break;
            }


            mm.DeliveryNotificationOptions = DeliveryNotificationOptions;
            Headers.SetColletion(mm.Headers);
            mm.Priority = Priority;

            if (Sender != null)
                mm.Sender = Sender.GetMailAddress();


            switch (SubjectEncoding)
            {
                case "System.Text.ASCIIEncoding":
                    mm.SubjectEncoding = Encoding.ASCII;
                    break;
                case "System.Text.UnicodeEncoding":
                    mm.SubjectEncoding = Encoding.Unicode;
                    break;
                case "System.Text.UTF32Encoding":
                    mm.SubjectEncoding = Encoding.UTF32;
                    break;
                case "System.Text.UTF7Encoding":
                    mm.SubjectEncoding = Encoding.UTF7;
                    break;
                case "System.Text.UTF8Encoding":
                    mm.SubjectEncoding = Encoding.UTF8;
                    break;
                default: mm.SubjectEncoding = null;
                    break;
            }


            foreach (SerializeableAlternateView av in AlternateViews)
                mm.AlternateViews.Add(av.GetAlternateView());

            return mm;
        }

    }



    [Serializable]
    internal class SerializeableLinkedResource
    {
        String ContentId;
        Uri ContentLink;
        Stream ContentStream;
        SerializeableContentType ContentType;
        TransferEncoding TransferEncoding;

        internal static SerializeableLinkedResource GetSerializeableLinkedResource(LinkedResource lr)
        {
            if (lr == null)
                return null;

            SerializeableLinkedResource slr = new SerializeableLinkedResource();
            slr.ContentId = lr.ContentId;
            slr.ContentLink = lr.ContentLink;

            if (lr.ContentStream != null)
            {
                byte[] bytes = new byte[lr.ContentStream.Length];
                lr.ContentStream.Read(bytes, 0, bytes.Length);
                slr.ContentStream = new MemoryStream(bytes);
            }

            slr.ContentType = SerializeableContentType.GetSerializeableContentType(lr.ContentType);
            slr.TransferEncoding = lr.TransferEncoding;

            return slr;

        }

        internal LinkedResource GetLinkedResource()
        {
            LinkedResource slr = new LinkedResource(ContentStream);
            slr.ContentId = ContentId;
            slr.ContentLink = ContentLink;

            slr.ContentType = ContentType.GetContentType();
            slr.TransferEncoding = TransferEncoding;

            return slr;
        }
    }

    [Serializable]
    internal class SerializeableAlternateView
    {

        Uri BaseUri;
        String ContentId;
        Stream ContentStream;
        SerializeableContentType ContentType;
        List<SerializeableLinkedResource> LinkedResources = new List<SerializeableLinkedResource>();
        TransferEncoding TransferEncoding;

        internal static SerializeableAlternateView GetSerializeableAlternateView(AlternateView av)
        {
            if (av == null)
                return null;

            SerializeableAlternateView sav = new SerializeableAlternateView();

            sav.BaseUri = av.BaseUri;
            sav.ContentId = av.ContentId;

            if (av.ContentStream != null)
            {
                byte[] bytes = new byte[av.ContentStream.Length];
                av.ContentStream.Read(bytes, 0, bytes.Length);
                sav.ContentStream = new MemoryStream(bytes);
            }

            sav.ContentType = SerializeableContentType.GetSerializeableContentType(av.ContentType);

            foreach (LinkedResource lr in av.LinkedResources)
                sav.LinkedResources.Add(SerializeableLinkedResource.GetSerializeableLinkedResource(lr));

            sav.TransferEncoding = av.TransferEncoding;
            return sav;
        }

        internal AlternateView GetAlternateView()
        {

            AlternateView sav = new AlternateView(ContentStream);

            sav.BaseUri = BaseUri;
            sav.ContentId = ContentId;

            sav.ContentType = ContentType.GetContentType();

            foreach (SerializeableLinkedResource lr in LinkedResources)
                sav.LinkedResources.Add(lr.GetLinkedResource());

            sav.TransferEncoding = TransferEncoding;
            return sav;
        }
    }

    [Serializable]
    internal class SerializeableMailAddress
    {
        String User;
        String Host;
        String Address;
        String DisplayName;

        internal static SerializeableMailAddress GetSerializeableMailAddress(MailAddress ma)
        {
            if (ma == null)
                return null;
            SerializeableMailAddress sma = new SerializeableMailAddress();

            sma.User = ma.User;
            sma.Host = ma.Host;
            sma.Address = ma.Address;
            sma.DisplayName = ma.DisplayName;
            return sma;
        }

        internal MailAddress GetMailAddress()
        {
            return new MailAddress(Address, DisplayName);
        }
    }

    [Serializable]
    internal class SerializeableContentDisposition
    {
        DateTime CreationDate;
        String DispositionType;
        String FileName;
        Boolean Inline;
        DateTime ModificationDate;
        SerializeableCollection Parameters;
        DateTime ReadDate;
        long Size;

        internal static SerializeableContentDisposition GetSerializeableContentDisposition(System.Net.Mime.ContentDisposition cd)
        {
            if (cd == null)
                return null;

            SerializeableContentDisposition scd = new SerializeableContentDisposition();
            scd.CreationDate = cd.CreationDate;
            scd.DispositionType = cd.DispositionType;
            scd.FileName = cd.FileName;
            scd.Inline = cd.Inline;
            scd.ModificationDate = cd.ModificationDate;
            scd.Parameters = SerializeableCollection.GetSerializeableCollection(cd.Parameters);
            scd.ReadDate = cd.ReadDate;
            scd.Size = cd.Size;

            return scd;
        }

        internal void SetContentDisposition(ContentDisposition scd)
        {
            scd.CreationDate = CreationDate;
            scd.DispositionType = DispositionType;
            scd.FileName = FileName;
            scd.Inline = Inline;
            scd.ModificationDate = ModificationDate;
            Parameters.SetColletion(scd.Parameters);

            scd.ReadDate = ReadDate;
            scd.Size = Size;
        }
    }

    [Serializable]
    internal class SerializeableContentType
    {
        String Boundary;
        String CharSet;
        String MediaType;
        String Name;
        SerializeableCollection Parameters;

        internal static SerializeableContentType GetSerializeableContentType(System.Net.Mime.ContentType ct)
        {
            if (ct == null)
                return null;

            SerializeableContentType sct = new SerializeableContentType();

            sct.Boundary = ct.Boundary;
            sct.CharSet = ct.CharSet;
            sct.MediaType = ct.MediaType;
            sct.Name = ct.Name;
            sct.Parameters = SerializeableCollection.GetSerializeableCollection(ct.Parameters);

            return sct;
        }

        internal ContentType GetContentType()
        {

            ContentType sct = new ContentType();

            sct.Boundary = Boundary;
            sct.CharSet = CharSet;
            sct.MediaType = MediaType;
            sct.Name = Name;

            Parameters.SetColletion(sct.Parameters);

            return sct;
        }
    }

    [Serializable]
    internal class SerializeableAttachment
    {
        String ContentId;
        SerializeableContentDisposition ContentDisposition;
        SerializeableContentType ContentType;
        MemoryStream ContentStream;
        System.Net.Mime.TransferEncoding TransferEncoding;
        String Name;
        // Encoding NameEncoding = Encoding.ASCII;
        string NameEncoding = string.Empty;
        internal static SerializeableAttachment GetSerializeableAttachment(Attachment att)
        {
            if (att == null)
                return null;

            SerializeableAttachment saa = new SerializeableAttachment();
            saa.ContentId = att.ContentId;
            saa.ContentDisposition = SerializeableContentDisposition.GetSerializeableContentDisposition(att.ContentDisposition);

            if (att.ContentStream != null)
            {
                byte[] bytes = new byte[att.ContentStream.Length];
                att.ContentStream.Read(bytes, 0, bytes.Length);

                saa.ContentStream = new MemoryStream(bytes);
            }

            saa.ContentType = SerializeableContentType.GetSerializeableContentType(att.ContentType);
            saa.Name = att.Name;
            saa.TransferEncoding = att.TransferEncoding;

            if (att.NameEncoding != null)
            {
                saa.NameEncoding = att.NameEncoding.GetType().ToString();
            }

            return saa;
        }

        internal Attachment GetAttachment()
        {
            Attachment saa = new Attachment(ContentStream, Name);
            saa.ContentId = ContentId;
            this.ContentDisposition.SetContentDisposition(saa.ContentDisposition);

            saa.ContentType = ContentType.GetContentType();
            saa.Name = Name;
            saa.TransferEncoding = TransferEncoding;

            switch (NameEncoding)
            {
                case "System.Text.ASCIIEncoding":
                    saa.NameEncoding = Encoding.ASCII;
                    break;
                case "System.Text.UnicodeEncoding":
                    saa.NameEncoding = Encoding.Unicode;
                    break;
                case "System.Text.UTF32Encoding":
                    saa.NameEncoding = Encoding.UTF32;
                    break;
                case "System.Text.UTF7Encoding":
                    saa.NameEncoding = Encoding.UTF7;
                    break;
                case "System.Text.UTF8Encoding":
                    saa.NameEncoding = Encoding.UTF8;
                    break;
                default: saa.NameEncoding = null;
                    break;
            }

            return saa;
        }
    }

    [Serializable]
    internal class SerializeableCollection
    {
        Dictionary<string, string> Collection = new Dictionary<string, string>();

        internal static SerializeableCollection GetSerializeableCollection(NameValueCollection col)
        {

            if (col == null)
                return null;

            SerializeableCollection scol = new SerializeableCollection();
            foreach (String key in col.Keys)
                scol.Collection.Add(key, col[key]);

            return scol;
        }

        internal static SerializeableCollection GetSerializeableCollection(StringDictionary col)
        {
            if (col == null)
                return null;

            SerializeableCollection scol = new SerializeableCollection();
            foreach (String key in col.Keys)
                scol.Collection.Add(key, col[key]);

            return scol;
        }

        internal void SetColletion(NameValueCollection scol)
        {

            foreach (String key in Collection.Keys)
            {
                scol.Add(key, this.Collection[key]);
            }

        }

        internal void SetColletion(StringDictionary scol)
        {
            foreach (String key in Collection.Keys)
            {
                if (scol.ContainsKey(key))
                    scol[key] = Collection[key];
                else
                    scol.Add(key, this.Collection[key]);
            }
        }
    }

}

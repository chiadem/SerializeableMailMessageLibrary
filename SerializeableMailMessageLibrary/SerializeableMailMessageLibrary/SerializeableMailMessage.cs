using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace SerializeableMailMessageLibrary
{
    ///

    /// Serializeable mailmessage object
    ///
    [Serializable]
    public class SerializeableMailMessage
    {
        public Guid MessageId { get; set; }
        public String Department { get; set; }
        public Boolean AreAttachmentsEmbedded { get; set; }
        public String UserEmail { get; set; }
        Boolean IsBodyHtml { get; set; }
        String Body { get; set; }
        SerializeableMailAddress From { get; set; }
        readonly List<SerializeableMailAddress> _to = new List<SerializeableMailAddress>();
        readonly List<SerializeableMailAddress> _cc = new List<SerializeableMailAddress>();
        readonly List<SerializeableMailAddress> _bcc = new List<SerializeableMailAddress>();
        readonly List<SerializeableMailAddress> _replyToList = new List<SerializeableMailAddress>();
        SerializeableMailAddress Sender { get; set; }
        String Subject { get; set; }
        readonly List<SerializeableAttachment> _attachments = new List<SerializeableAttachment>();
        readonly string _bodyEncoding = string.Empty;
        readonly string _subjectEncoding = string.Empty;
        readonly DeliveryNotificationOptions _deliveryNotificationOptions;
        readonly SerializeableCollection _headers;
        readonly MailPriority _priority;
        readonly List<SerializeableAlternateView> _alternateViews = new List<SerializeableAlternateView>();


        ///

        /// Creates a new serializeable mailmessage based on a MailMessage object
        ///

        /// 
        public SerializeableMailMessage(MailMessage mailMessage, Guid messageId, string department)
        {
            MessageId = messageId == Guid.Empty ? Guid.NewGuid() : messageId;

            Department = string.IsNullOrEmpty(department) ? string.Empty : department;


            IsBodyHtml = mailMessage.IsBodyHtml;
            Body = mailMessage.Body;
            Subject = mailMessage.Subject;
            From = SerializeableMailAddress.GetSerializeableMailAddress(mailMessage.From);
            _to = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in mailMessage.To)
            {
                _to.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }

            _cc = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in mailMessage.CC)
            {
                _cc.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }

            _bcc = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in mailMessage.Bcc)
            {
                _bcc.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }

            _attachments = new List<SerializeableAttachment>();
            foreach (Attachment att in mailMessage.Attachments)
            {
                _attachments.Add(SerializeableAttachment.GetSerializeableAttachment(att));
            }

            _replyToList = new List<SerializeableMailAddress>();
            foreach (MailAddress ma in mailMessage.ReplyToList)
            {
                _replyToList.Add(SerializeableMailAddress.GetSerializeableMailAddress(ma));
            }
            if (mailMessage.BodyEncoding != null)
            {
                _bodyEncoding = mailMessage.BodyEncoding.GetType().ToString();
            }

            _deliveryNotificationOptions = mailMessage.DeliveryNotificationOptions;
            _headers = SerializeableCollection.GetSerializeableCollection(mailMessage.Headers);
            _priority = mailMessage.Priority;
            Sender = SerializeableMailAddress.GetSerializeableMailAddress(mailMessage.Sender);


            if (mailMessage.SubjectEncoding != null)
            {
                _subjectEncoding = mailMessage.SubjectEncoding.GetType().ToString();
            }

            foreach (AlternateView av in mailMessage.AlternateViews)
                _alternateViews.Add(SerializeableAlternateView.GetSerializeableAlternateView(av));
        }

        public MailMessage GetMailMessage()
        {
            MailMessage mm = new MailMessage { IsBodyHtml = IsBodyHtml, Body = Body, Subject = Subject };

            if (From != null)
                mm.From = From.GetMailAddress();

            foreach (SerializeableMailAddress ma in _to)
            {
                mm.To.Add(ma.GetMailAddress());
            }

            foreach (SerializeableMailAddress ma in _cc)
            {
                mm.CC.Add(ma.GetMailAddress());
            }

            foreach (SerializeableMailAddress ma in _bcc)
            {
                mm.Bcc.Add(ma.GetMailAddress());
            }

            foreach (SerializeableMailAddress ma in _replyToList)
            {
                mm.ReplyToList.Add(ma.GetMailAddress());
            }

            foreach (SerializeableAttachment att in _attachments)
            {
                mm.Attachments.Add(att.GetAttachment());
            }


            switch (_bodyEncoding)
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


            mm.DeliveryNotificationOptions = _deliveryNotificationOptions;
            _headers.SetColletion(mm.Headers);
            mm.Priority = _priority;

            if (Sender != null)
                mm.Sender = Sender.GetMailAddress();


            switch (_subjectEncoding)
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


            foreach (SerializeableAlternateView av in _alternateViews)
                mm.AlternateViews.Add(av.GetAlternateView());

            return mm;
        }

    }



    [Serializable]
    internal class SerializeableLinkedResource : IDisposable
    {
        String _contentId;
        Uri _contentLink;
        Stream _contentStream;
        SerializeableContentType _contentType;
        TransferEncoding _transferEncoding;

        internal static SerializeableLinkedResource GetSerializeableLinkedResource(LinkedResource lr)
        {
            if (lr == null)
                return null;
            using (SerializeableLinkedResource slr = new SerializeableLinkedResource
                                                     {
                                                         _contentId = lr.ContentId,
                                                         _contentLink = lr.ContentLink
                                                     })
            {
                if (lr.ContentStream != null)
                {
                    byte[] bytes = new byte[lr.ContentStream.Length];
                    lr.ContentStream.Read(bytes, 0, bytes.Length);
                    slr._contentStream = new MemoryStream(bytes);
                }

                slr._contentType = SerializeableContentType.GetSerializeableContentType(lr.ContentType);
                slr._transferEncoding = lr.TransferEncoding;
                return slr;
            }
        }

        internal LinkedResource GetLinkedResource()
        {
            LinkedResource slr = new LinkedResource(_contentStream)
            {
                ContentId = _contentId,
                ContentLink = _contentLink,
                ContentType = _contentType.GetContentType(),
                TransferEncoding = _transferEncoding
            };
            return slr;
        }

        public void Dispose()
        {
            if (_contentStream != null)
            {
                _contentStream.Dispose();
                _contentStream = null;
            }
        }
    }

    [Serializable]
    internal class SerializeableAlternateView : IDisposable
    {
        Uri _baseUri;
        String _contentId;
        Stream _contentStream;
        SerializeableContentType _contentType;
        readonly List<SerializeableLinkedResource> _linkedResources = new List<SerializeableLinkedResource>();
        TransferEncoding _transferEncoding;

        internal static SerializeableAlternateView GetSerializeableAlternateView(AlternateView av)
        {
            if (av == null)
                return null;
            using (SerializeableAlternateView sav = new SerializeableAlternateView
                                                    {
                                                        _baseUri = av.BaseUri,
                                                        _contentId = av.ContentId
                                                    })
            {
                if (av.ContentStream != null)
                {
                    byte[] bytes = new byte[av.ContentStream.Length];
                    av.ContentStream.Read(bytes, 0, bytes.Length);
                    sav._contentStream = new MemoryStream(bytes);
                }

                sav._contentType = SerializeableContentType.GetSerializeableContentType(av.ContentType);

                foreach (LinkedResource lr in av.LinkedResources)
                    sav._linkedResources.Add(SerializeableLinkedResource.GetSerializeableLinkedResource(lr));

                sav._transferEncoding = av.TransferEncoding;
                return sav;
            }
        }

        internal AlternateView GetAlternateView()
        {
            AlternateView sav = new AlternateView(_contentStream)
            {
                BaseUri = _baseUri,
                ContentId = _contentId,
                ContentType = _contentType.GetContentType()
            };

            foreach (SerializeableLinkedResource lr in _linkedResources)
                sav.LinkedResources.Add(lr.GetLinkedResource());

            sav.TransferEncoding = _transferEncoding;
            return sav;
        }
        public void Dispose()
        {
            if (_contentStream != null)
            {
                _contentStream.Dispose();
                _contentStream = null;
            }
        }
    }

    [Serializable]
    internal class SerializeableMailAddress
    {
        String _user;
        String _host;
        String _address;
        String _displayName;

        internal static SerializeableMailAddress GetSerializeableMailAddress(MailAddress ma)
        {
            if (ma == null)
                return null;
            SerializeableMailAddress sma = new SerializeableMailAddress
            {
                _user = ma.User,
                _host = ma.Host,
                _address = ma.Address,
                _displayName = ma.DisplayName
            };
            return sma;
        }

        internal MailAddress GetMailAddress()
        {
            return new MailAddress(_address, _displayName);
        }
    }

    [Serializable]
    internal class SerializeableContentDisposition
    {
        DateTime _creationDate;
        String _dispositionType;
        String _fileName;
        Boolean _inline;
        DateTime _modificationDate;
        SerializeableCollection _parameters;
        DateTime _readDate;
        long _size;

        internal static SerializeableContentDisposition GetSerializeableContentDisposition(ContentDisposition cd)
        {
            if (cd == null)
                return null;

            SerializeableContentDisposition scd = new SerializeableContentDisposition
            {
                _creationDate = cd.CreationDate,
                _dispositionType = cd.DispositionType,
                _fileName = cd.FileName,
                _inline = cd.Inline,
                _modificationDate = cd.ModificationDate,
                _parameters = SerializeableCollection.GetSerializeableCollection(cd.Parameters),
                _readDate = cd.ReadDate,
                _size = cd.Size
            };

            return scd;
        }

        internal void SetContentDisposition(ContentDisposition scd)
        {
            scd.CreationDate = _creationDate;
            scd.DispositionType = _dispositionType;
            scd.FileName = _fileName;
            scd.Inline = _inline;
            scd.ModificationDate = _modificationDate;
            _parameters.SetColletion(scd.Parameters);

            scd.ReadDate = _readDate;
            scd.Size = _size;
        }
    }

    [Serializable]
    internal class SerializeableContentType
    {
        String _boundary;
        String _charSet;
        String _mediaType;
        String _name;
        SerializeableCollection _parameters;

        internal static SerializeableContentType GetSerializeableContentType(ContentType ct)
        {
            if (ct == null)
                return null;

            SerializeableContentType sct = new SerializeableContentType
            {
                _boundary = ct.Boundary,
                _charSet = ct.CharSet,
                _mediaType = ct.MediaType,
                _name = ct.Name,
                _parameters = SerializeableCollection.GetSerializeableCollection(ct.Parameters)
            };

            return sct;
        }

        internal ContentType GetContentType()
        {

            ContentType sct = new ContentType
            {
                Boundary = _boundary,
                CharSet = _charSet,
                MediaType = _mediaType,
                Name = _name
            };

            _parameters.SetColletion(sct.Parameters);

            return sct;
        }
    }

    [Serializable]
    internal class SerializeableAttachment : IDisposable
    {
        String _contentId;
        SerializeableContentDisposition _contentDisposition;
        SerializeableContentType _contentType;
        MemoryStream _contentStream;
        TransferEncoding _transferEncoding;
        String _name;
        // Encoding NameEncoding = Encoding.ASCII;
        string _nameEncoding = string.Empty;

        internal static SerializeableAttachment GetSerializeableAttachment(Attachment att)
        {
            if (att == null)
                return null;
            using (SerializeableAttachment saa = new SerializeableAttachment
                                                 {
                                                     _contentId = att.ContentId,
                                                     _contentDisposition =
                                                         SerializeableContentDisposition.GetSerializeableContentDisposition(att.ContentDisposition)
                                                 })
            {
                if (att.ContentStream != null)
                {
                    byte[] bytes = new byte[att.ContentStream.Length];
                    att.ContentStream.Read(bytes, 0, bytes.Length);

                    saa._contentStream = new MemoryStream(bytes);
                }

                saa._contentType = SerializeableContentType.GetSerializeableContentType(att.ContentType);
                saa._name = att.Name;
                saa._transferEncoding = att.TransferEncoding;

                if (att.NameEncoding != null)
                {
                    saa._nameEncoding = att.NameEncoding.GetType().ToString();
                }

                return saa;
            }
        }

        internal Attachment GetAttachment()
        {
            Attachment saa = new Attachment(_contentStream, _name) { ContentId = _contentId };
            _contentDisposition.SetContentDisposition(saa.ContentDisposition);

            saa.ContentType = _contentType.GetContentType();
            saa.Name = _name;
            saa.TransferEncoding = _transferEncoding;

            switch (_nameEncoding)
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

        public void Dispose()
        {
            if (_contentStream != null)
            {
                _contentStream.Dispose();
                _contentStream = null;
            }
        }
    }

    [Serializable]
    internal class SerializeableCollection
    {
        readonly Dictionary<string, string> _collection = new Dictionary<string, string>();

        internal static SerializeableCollection GetSerializeableCollection(NameValueCollection col)
        {

            if (col == null)
                return null;

            SerializeableCollection scol = new SerializeableCollection();
            foreach (String key in col.Keys)
                scol._collection.Add(key, col[key]);

            return scol;
        }

        internal static SerializeableCollection GetSerializeableCollection(StringDictionary col)
        {
            if (col == null)
                return null;

            SerializeableCollection scol = new SerializeableCollection();
            foreach (String key in col.Keys)
                scol._collection.Add(key, col[key]);

            return scol;
        }

        internal void SetColletion(NameValueCollection scol)
        {

            foreach (String key in _collection.Keys)
            {
                scol.Add(key, _collection[key]);
            }

        }

        internal void SetColletion(StringDictionary scol)
        {
            foreach (String key in _collection.Keys)
            {
                if (scol.ContainsKey(key))
                    scol[key] = _collection[key];
                else
                    scol.Add(key, _collection[key]);
            }
        }
    }

}

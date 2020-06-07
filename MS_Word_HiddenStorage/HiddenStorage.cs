using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml.Linq;
using Interop = Microsoft.Office.Interop.Word;

namespace OfficeTools {
    public class HiddenStorage {
        public HiddenStorage(Interop.Document document, string key) {
            if (key == null) throw new ArgumentNullException(nameof(key));
            if (string.IsNullOrWhiteSpace(key)) {
                throw new ArgumentException($"Argument '{key}' cannot be null or white space", nameof(key));
            }
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _key = SecurityElement.Escape(key);
        }

        readonly string _key;
        readonly Interop.Document _document;
        string _id;

        public string Read() {
            if (!FindAndSetIdByKey()) {
                throw new KeyNotFoundException();
            }

            CustomXMLPart xmlPart = _document.CustomXMLParts.SelectByID(_id);

            XElement root = XDocument.Parse(xmlPart.XML).Root;

            return root?.Value == null ? null : FromBase64(root.Value);
        }

        public void Write(string @object) {
            XElement root = new XElement(_key) {
                Value = ToBase64(@object)
            };
            string data = root.ToString();

            if (FindAndSetIdByKey()) {
                _document.CustomXMLParts.SelectByID(_id).Delete();
            }

            _id = _document.CustomXMLParts.Add(data).Id;
        }

        public void Delete() {
            if (!FindAndSetIdByKey()) {
                throw new KeyNotFoundException();
            }

            CustomXMLPart xmlPart = _document.CustomXMLParts.SelectByID(_id);

            xmlPart.Delete();
        }

        bool FindAndSetIdByKey() {
            _id = _document.CustomXMLParts
                .Cast<CustomXMLPart>()
                .SingleOrDefault(p => XDocument.Parse(p.XML).Root?.Name == _key)
                ?.Id;

            return _id != null;
        }

        static string FromBase64(string value) =>
            Encoding.UTF8.GetString(Convert.FromBase64String(value));

        static string ToBase64(string value) =>
            Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
    }
}
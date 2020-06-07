using Microsoft.Office.Interop.Word;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeTools.Tests {
    [TestFixture]
    public class OfficeToolsTests {
        [SetUp]
        public void Setup() {
            _application = new Application {
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };
        }

        [TearDown]
        public void TearDown() {
            int processId = 0;
            try {
                if (_application.Windows.Count > 0) {
                    int hWnd = _application.Windows[1].Hwnd;
                    GetWindowThreadProcessId((IntPtr)hWnd, out processId);
                }

                _application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                _application.Quit(WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
            finally {
                try {
                    Thread.Sleep(3000);
                    Process.GetProcessById(processId).Kill();
                }
                catch { }
            }
            try {
                if (_tmpFile != null && File.Exists(_tmpFile)) File.Delete(_tmpFile);
            }
            catch { }
        }

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        Application _application;
        string _tmpFile;

        [Test]
        public void Ctor_Test_ThrowsNREForNullKey() {
            Document document = _application.Documents.Add();
            Assert.Throws<ArgumentNullException>(() => _ = new HiddenStorage(document, null));
        }

        [Test]
        public void Ctor_Test_ThrowsArgumentExceptionForEmptyKey() {
            Document document = _application.Documents.Add();
            Assert.Throws<ArgumentException>(() => _ = new HiddenStorage(document, " "));
        }

        [Test]
        public void Ctor_Test_ThrowsNREForNullDocument() {
            Assert.Throws<ArgumentNullException>(() => _ = new HiddenStorage(null, "test"));
        }

        [Test]
        public void Read_Test_WrongKey_ThrowsKeyNotFoundException() {
            Document document = _application.Documents.Add();
            Assert.Throws<KeyNotFoundException>(() => new HiddenStorage(document, "nonExistingKey").Read());
        }

        [Test]
        public void Write_Read_Test() {
            string expected = @"{ test: ""testvalue""}";

            Document document = _application.Documents.Add();
            string key = "testKey";
            HiddenStorage storage = new HiddenStorage(document, key);

            storage.Write(expected);
            string actual = storage.Read();

            CollectionAssert.AreEqual(expected, actual);
        }

        [Test]
        public void Write_CloseFile_OpenFile_Read_Test() {
            string expected = @"{ test: ""testvalue""}";
            _tmpFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");

            Document document = _application.Documents.Add();
            string key = "testKey";
            HiddenStorage storage = new HiddenStorage(document, key);
            storage.Write(expected);
            document.SaveAs(_tmpFile, WdSaveFormat.wdFormatXMLDocument);
            document.Close();

            Document documentReopened = _application.Documents.Open(_tmpFile);
            HiddenStorage storageReopened = new HiddenStorage(documentReopened, key);
            string actual = storageReopened.Read();

            CollectionAssert.AreEqual(expected, actual);
        }

        [Test]
        public void Write_Delete_Read_Test_ThrowsKeyNotFoundException() {
            string data = @"_";

            Document document = _application.Documents.Add();
            string key = "testKey";
            HiddenStorage storage = new HiddenStorage(document, key);

            storage.Write(data);
            storage.Delete();
            Assert.Throws<KeyNotFoundException>(() => storage.Read());
        }
    }
}
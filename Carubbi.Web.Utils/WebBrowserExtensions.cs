using Carubbi.JavascriptWatcher;
using Carubbi.WindowsAppHelper;
using Microsoft.Win32;
using mshtml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace Carubbi.Web.Utils
{
    /// <summary>
    /// Estrutura de dados com opções para manipulação da página html
    /// </summary>
    internal class WebBrowserParams
    {
        /// <summary>
        /// Quantidade de Postbacks encontrados
        /// </summary>
        internal int PostbackCounts { get; set; }

     
        /// <summary>
        /// Textos encontrados em Alerts
        /// </summary>
        internal StringCollection MonitoredAlertTexts { get; set; }

        /// <summary>
        /// Objeto responsável por monitorar de eventos javascript em uma página
        /// </summary>
        internal JavascriptWatcher.JavascriptWatcher JsWatcher { get; set; }

        /// <summary>
        /// Indica se um alert foi disparado
        /// </summary>
        public bool AlertRaised { get; set; }

        /// <summary>
        /// Texto do alert disparado
        /// </summary>
        public string AlertRaisedText { get; set; }


        /// <summary>
        /// Url do window.open disparado
        /// </summary>
        public string WindowOpenUrl { get; set; }
    }


    /// <summary>
    /// Extension Methods da classe WebBrowser
    /// </summary>
    public static class WebBrowserExtensions
    {
        private static Dictionary<string, WebBrowserParams> _webBrowserInstances = new Dictionary<string, WebBrowserParams>();

        /// <summary>
        /// Recupera a quantidade corrente de postbacks efetuados no browser
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        public static int GetPostbackCounter(this WebBrowser instance)
        {
            return _webBrowserInstances.ContainsKey(instance.Name) ? _webBrowserInstances[instance.Name].PostbackCounts : 0;
        }

        /// <summary>
        /// Inicia o processo de contar a quantidade de postbacks do browser para poder recuperar com o método WaitForPostbacks
        /// </summary>
        /// <param name="instance"></param>
        public static void StartMonitoringPostbacks(this WebBrowser instance)
        {
            instance.DocumentCompleted += Instance_DocumentCompletedWaitPostback;
        }

        internal static void Instance_DocumentCompletedWaitPostback(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

            if (_webBrowserInstances.ContainsKey(((WebBrowser)sender).Name))
                _webBrowserInstances[((WebBrowser)sender).Name].PostbackCounts++;
            else
            {
                _webBrowserInstances.Add(((WebBrowser)sender).Name, new WebBrowserParams());
            }
        }

        /// <summary>
        /// Inicia o processo de monitoramento de alerts no browser
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="supressAlert"></param>
        /// <param name="supressWindowOpen"></param>
        public static void StartMonitoringJavascript(this WebBrowser instance, bool supressAlert, bool supressWindowOpen)
        {
            if (!_webBrowserInstances.ContainsKey(instance.Name))
            {
                _webBrowserInstances.Add(instance.Name, new WebBrowserParams());
                _webBrowserInstances[instance.Name].MonitoredAlertTexts = new StringCollection();
                _webBrowserInstances[instance.Name].WindowOpenUrl = string.Empty;
            }
            if (_webBrowserInstances[instance.Name] == null)
            {
                _webBrowserInstances[instance.Name] = new WebBrowserParams
                {
                    MonitoredAlertTexts = new StringCollection()
                };
            }
            _webBrowserInstances[instance.Name].JsWatcher = new JavascriptWatcher.JavascriptWatcher(instance);
            _webBrowserInstances[instance.Name].JsWatcher.Start(supressAlert, supressWindowOpen);
            _webBrowserInstances[instance.Name].JsWatcher.AlertIntercepted += JsWatcher_AlertIntercepted;
            _webBrowserInstances[instance.Name].JsWatcher.WindowOpenIntercepted += JsWatcher_WindowOpenIntercepted;

        }

        private static void JsWatcher_WindowOpenIntercepted(object sender, WindowOpenInterceptedEventArgs e)
        {
            var webBrowser = (WebBrowser)sender;
            _webBrowserInstances[webBrowser.Name].WindowOpenUrl = e.Url.ToString();
        }

        private static void JsWatcher_AlertIntercepted(object sender, AlertInterceptedEventArgs e)
        {
            var webBrowser = (WebBrowser)sender;
            _webBrowserInstances[webBrowser.Name].MonitoredAlertTexts.Add(e.AlertText);
            _webBrowserInstances[webBrowser.Name].AlertRaised = true;
            _webBrowserInstances[webBrowser.Name].AlertRaisedText = e.AlertText;
        }

        /// <summary>
        /// Ignora apenas alerts de erro de script
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="supressByDefault"></param>
        public static void SuppressScriptErrorsOnly(this WebBrowser instance, bool supressByDefault)
        {
            // Ensure that ScriptErrorsSuppressed is set to false.
            instance.ScriptErrorsSuppressed = supressByDefault;

            // Handle DocumentCompleted to gain access to the Document object.
            instance.DocumentCompleted += browser_DocumentCompleted;
        }

        private static void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            ((WebBrowser)sender).InvokeIfRequired(i =>
            {
                var htmlDocument = ((WebBrowser) i).Document;
                if (htmlDocument == null) return;
                if (htmlDocument.Window != null)
                    htmlDocument.Window.Error += Window_Error;
            });
        }

        private static void Window_Error(object sender, HtmlElementErrorEventArgs e)
        {
            // Ignore the error and suppress the error dialog box. 
            e.Handled = true;
        }

        /// <summary>
        /// Aguarda por alert ou postback por um determinado tempo
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="timeout"></param>
        /// <returns>ResponseType indica se foi lançado alert ou postback, responseText indica o texto do alert</returns>
        public static ResponseResult WaitResponse(this WebBrowser instance, int timeout)
        {
            var cronometer = new Stopwatch();
            cronometer.Start();

            if (!_webBrowserInstances.ContainsKey(instance.Name))
            {
                _webBrowserInstances.Add(instance.Name, new WebBrowserParams());
            }

            while (instance.GetPostbackCounter() < 1
                && !_webBrowserInstances[instance.Name].AlertRaised
                && string.IsNullOrEmpty(_webBrowserInstances[instance.Name].WindowOpenUrl)
                && cronometer.ElapsedMilliseconds < timeout)
            {
                Application.DoEvents();
            }
            cronometer.Stop();

            var result = new ResponseResult();

            if (cronometer.ElapsedMilliseconds >= timeout)
            {
                result.ResponseType = ResponseType.Timeout;
            }
            else if (!_webBrowserInstances[instance.Name].AlertRaised && string.IsNullOrEmpty(_webBrowserInstances[instance.Name].WindowOpenUrl))
            {
                if (_webBrowserInstances[instance.Name] != null)
                {
                    _webBrowserInstances[instance.Name].PostbackCounts = 0;
                    result.ResponseType = ResponseType.Postback;
                }
                else
                {
                    result.ResponseType = ResponseType.Unknown;
                }
            }
            else if (_webBrowserInstances[instance.Name].AlertRaised)
            {
                result.ResponseType = ResponseType.Alert;
                result.ResponseText = _webBrowserInstances[instance.Name].AlertRaisedText;
                _webBrowserInstances[instance.Name].AlertRaisedText = string.Empty;
                _webBrowserInstances[instance.Name].AlertRaised = false;
            }
            else if (!string.IsNullOrEmpty(_webBrowserInstances[instance.Name].WindowOpenUrl))
            {
                result.ResponseType = ResponseType.WindowOpen;
                result.ResponseText = _webBrowserInstances[instance.Name].WindowOpenUrl;
                _webBrowserInstances[instance.Name].WindowOpenUrl = string.Empty;
            }
            else
            {
                result.ResponseType = ResponseType.Unknown;
            }


            return result;
        }

        /// <summary>
        /// Prepara o webbrowser para monitorar postbacks, alerts e suprimir erros de script
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="supressAlert"></param>
        /// <param name="supressWindowOpen"></param>
        /// <param name="supressByDefault"></param>
        public static void Initialize(this WebBrowser instance, bool supressAlert, bool supressWindowOpen, bool supressByDefault)
        {
            instance.StartMonitoringPostbacks();
            instance.StartMonitoringJavascript(supressAlert, supressWindowOpen);
            instance.SuppressScriptErrorsOnly(supressByDefault);
            instance.SupressCertificateErrors();
        }

        /// <summary>
        /// Descarta o objeto WebBrowser e seus dependentes
        /// </summary>
        /// <param name="instance"></param>
        public static void DisposeBrowser(this WebBrowser instance)
        {
            if (_webBrowserInstances.ContainsKey(instance.Name))
            {
                instance.StopMonitoringJavascript();
                if (_webBrowserInstances[instance.Name] != null)
                {
                    if (_webBrowserInstances[instance.Name].MonitoredAlertTexts != null)
                    {
                        _webBrowserInstances[instance.Name].MonitoredAlertTexts.Clear();
                        _webBrowserInstances[instance.Name].MonitoredAlertTexts = null;
                    }
                }
                _webBrowserInstances.Remove(instance.Name);
                _webBrowserInstances[instance.Name] = null;
            }
         
            instance.Dispose();
        }


        /// <summary>
        /// Limpa variáveis temporárias utilizadas no monitoramento da página
        /// </summary>
        public static void ClearWebBrowserCache()
        {
            foreach (var entry in _webBrowserInstances)
            {
                if (entry.Value == null) continue;
                if (entry.Value.JsWatcher != null)
                {
                    entry.Value.JsWatcher.Stop();
                    entry.Value.JsWatcher.AlertIntercepted -= JsWatcher_AlertIntercepted;
                    entry.Value.JsWatcher.WindowOpenIntercepted -= JsWatcher_WindowOpenIntercepted;
                    entry.Value.JsWatcher = null;
                }

                if (entry.Value.MonitoredAlertTexts == null) continue;
                entry.Value.MonitoredAlertTexts.Clear();
                entry.Value.MonitoredAlertTexts = null;
            }

            _webBrowserInstances.Clear();
            _webBrowserInstances = null;
            _webBrowserInstances = new Dictionary<string, WebBrowserParams>();
        }

        /// <summary>
        /// Suprime erros de certificado
        /// </summary>
        /// <param name="instance"></param>
        public static void SupressCertificateErrors(this WebBrowser instance)
        {
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
        }


        /// <summary>
        /// Interrompe o monitoramento de Javascript na página
        /// </summary>
        /// <param name="instance"></param>
        public static void StopMonitoringJavascript(this WebBrowser instance)
        {
            if (!_webBrowserInstances.ContainsKey(instance.Name)) return;
            if (_webBrowserInstances[instance.Name] == null ||
                _webBrowserInstances[instance.Name].JsWatcher == null) return;
            _webBrowserInstances[instance.Name].JsWatcher.Stop();
            _webBrowserInstances[instance.Name].JsWatcher.AlertIntercepted -= JsWatcher_AlertIntercepted;
            _webBrowserInstances[instance.Name].JsWatcher.WindowOpenIntercepted -= JsWatcher_WindowOpenIntercepted;
            _webBrowserInstances[instance.Name].JsWatcher = null;
        }

        /// <summary>
        /// Inicializa um WebBrowser para ser controlado e automatizado
        /// </summary>
        /// <param name="instance"></param>
        public static void Initialize(this WebBrowser instance)
        {
            Initialize(instance, true, true, false);
        }


        private static readonly Dictionary<string, string> MimeTypeToExtension = new Dictionary<string, string>();

        private static Dictionary<string, string> _extensionToMimeType = new Dictionary<string, string>();


        /// <summary>
        /// Converte um mimetype em uma extensão de arquivo com base no registro do windows
        /// </summary>
        /// <param name="mimeType"></param>
        /// <returns></returns>
        public static string ConvertMimeTypeToExtension(string mimeType)
        {
            if (string.IsNullOrEmpty(mimeType.Trim()))
                throw new ArgumentNullException(nameof(mimeType));

            var key = $@"MIME\Database\Content Type\{mimeType}";
            if (MimeTypeToExtension.TryGetValue(key, out var result))
                return result;

            var regKey = Registry.ClassesRoot.OpenSubKey(key, false);
            var value = regKey?.GetValue("Extension", null);
            result = value != null ? value.ToString() : string.Empty;

            MimeTypeToExtension[key] = result;
            return result;
        }

        /// <summary>
        /// Recupera uma imagem do documento renderizado pelo WebBrowser e armazena em um bitmap
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        public static Bitmap PrintScreen(this WebBrowser instance)
        {
            return PrintScreen(instance, 0);
        }

        /// <summary>
        /// Recupera uma imagem do documento renderizado pelo WebBrowser e armazena em um bitmap após um determinado período em milisegundos
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="waitTime"></param>
        /// <returns></returns>
        public static Bitmap PrintScreen(this WebBrowser instance, int waitTime)
        {
            var originalDock = DockStyle.None;

            var originalWidth = instance.Width;
            var originalHeight = instance.Height;
            Bitmap bitmap = null;

            instance.InvokeIfRequired(i =>
            {
                var wb = (i as WebBrowser);
                try
                {
                    originalDock = wb.Dock;
                    wb.Dock = DockStyle.None;

                    wb.Width = originalWidth < 1200 ? 1200 : originalWidth;

                    wb.Height = originalHeight < 900 ? 900 : originalHeight;

                    wb.Parent.Focus();

                    bitmap = new Bitmap(wb.Width, wb.Height);

                    wb.DrawToBitmap(bitmap, new Rectangle(0, 0, bitmap.Width, bitmap.Height));

                    if (waitTime > 0)
                    {
                        wb.Focus();
                        Thread.Sleep(waitTime);
                    }
                    wb.Width = originalWidth;
                    wb.Height = originalHeight;
                }
                finally
                {
                    if (wb != null) wb.Dock = originalDock;
                }
            });

            return bitmap;

        }


        /// <summary>
        /// Copia uma página através da área de transferência e a transforma em bitmap
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="htmlElementId"></param>
        /// <returns></returns>
        public static Bitmap PrintHtmlImage(this WebBrowser instance, string htmlElementId)
        {
            var doc = (IHTMLDocument2)instance.Document?.DomDocument;
            var doc3 = (IHTMLDocument3)instance.Document?.DomDocument;
            var imgRange = (IHTMLControlRange)((HTMLBody)doc?.body)?.createControlRange();

            imgRange?.add((IHTMLControlElement)doc3?.getElementById(htmlElementId));

            imgRange?.execCommand("Copy");

            var bmp = (Bitmap)Clipboard.GetDataObject()?.GetData(DataFormats.Bitmap);

            return bmp;
        }


        public static Bitmap PrintHtmlImage(this WebBrowser instance, HtmlElement element)
        {
            var doc = (IHTMLDocument2)instance.Document?.DomDocument;
            var imgRange = (IHTMLControlRange)((HTMLBody)doc?.body)?.createControlRange();

            imgRange?.add((IHTMLControlElement)element.DomElement);

            imgRange?.execCommand("Copy");

            var bmp = (Bitmap)Clipboard.GetDataObject()?.GetData(DataFormats.Bitmap);

            return bmp;
        }
    }
}

using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ap = HtmlAgilityPack;
namespace Carubbi.Web.Utils
{
    /// <summary>
    /// Extension Methods para manipular objetos relacionados ao DOM Html
    /// </summary>
    public static class HtmlExtensions
    {
        #region Atributos html
        public const string NAME_HTML_ATTRIBUTE = "name";
        public const string CLASS_HTML_ATTRIBUTE = "className";
        private const string VALUE_HTML_ATTRIBUTE = "value";
        private const string SELECTED_INDEX_HTML_ATTRIBUTE = "selectedIndex";
        private const string CHECKED_HTML_ATTRIBUTE = "checked";
        #endregion

        /// <summary>
        /// Recupera Valor de um atributo
        /// </summary>
        /// <param name="instance">elemento DOM</param>
        /// <param name="attributeName">Nome do atributo</param>
        /// <returns></returns>
        public static string GetAttributeValue(this HtmlElement instance, string attributeName)
        {

            var attributeValue = string.Empty;

            var domNode = (IHTMLDOMNode)instance.DomElement;
            var attrs = (IHTMLAttributeCollection)domNode.attributes;

            foreach (IHTMLDOMAttribute attr in attrs)
            {
                if (!attr.nodeName.Equals(attributeName)) continue;
                var attrValue = attr.nodeValue as string;
                if (string.IsNullOrEmpty(attrValue)) continue;
                attributeValue = attrValue;
                break;
            }

            return attributeValue;
        }

        /// <summary>
        /// Recupera o texto do item selecionado de um DropDown
        /// </summary>
        /// <param name="dropdown">Elemento DOM do Dropdown</param>
        /// <returns>Texto do item selecionado</returns>
        public static string GetSelectedDropDownItemText(this HtmlElement dropdown)
        {
            foreach (HtmlElement item in dropdown.Children)
            {
                if (item.OuterHtml.Trim().ToLower().Contains("selected"))
                {
                    return item.InnerText.Trim();
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Seleciona o item de um dropdown a partir do texto informado
        /// </summary>
        /// <param name="dropdown">Elemento DOM do Dropdown</param>
        /// <param name="textoSelecionar">Texto para buscar o item</param>
        public static void SelectDropdownElementByText(this HtmlElement dropdown, string textoSelecionar)
        {
            var index = 0;
            foreach(HtmlElement item in dropdown.Children)
            {
                if (item.InnerText != null && string.Equals(item.InnerText.Trim(), textoSelecionar.Trim(), StringComparison.CurrentCultureIgnoreCase))
                {
                    dropdown.SelectDropdownElement(index);
                    break;
                }
                index++;
            }
        }

        /// <summary>
        /// Retorna uma lista de nós do DOM filhos de um determinado elemento
        /// </summary>
        /// <param name="instance">Elemento DOM Pai</param>
        /// <returns>Lista de nós filhos</returns>
        public static IEnumerable<ap.HtmlNode> GetListItems(this HtmlElement instance)
        {
            var d = new HtmlAgilityPack.HtmlDocument();
            d.LoadHtml(instance.InnerHtml);
            foreach (var node in d.DocumentNode.ChildNodes.Where(item => item.Name == "#text" && !string.IsNullOrEmpty(item.InnerText.Trim())))
            {
                yield return node;
            }
        }

        /// <summary>
        /// Recupera o texto de um elemento filho a partir do nome e indice do mesmo
        /// </summary>
        /// <param name="instance">Elemento DOM Pai</param>
        /// <param name="tagName">Nome do elemento filho a ser procurado</param>
        /// <param name="tagOrder">Índice do elemento filho </param>
        /// <returns></returns>
        public static string GetTagValue(this ap.HtmlNode instance, string tagName, int tagOrder)
        {
            return instance.Descendants().Where(d => d.Name == tagName).ToList()[tagOrder - 1].InnerText;
        }


        /// <summary>
        /// Envia dados de um formulario
        /// </summary>
        /// <param name="instance">Elemento DOM do formulario</param>
        public static void Submit(this HtmlElement instance)
        {
            instance.InvokeMember("submit");
        }

        /// <summary>
        /// Marca/Desmarca um checkbox
        /// </summary>
        /// <param name="instance">Elemento DOM do checkbox</param>
        /// <param name="value">Booleano que indica se é para marcar/desmarcar o checkbox</param>
        public static void Check(this HtmlElement instance, bool value)
        {
            ((IHTMLInputElement) instance.DomElement).@checked = value;
        }

        /// <summary>
        /// Define um valor para uma caixa de texto
        /// </summary>
        /// <param name="instance">Elemento DOM do Textbox</param>
        /// <param name="value">Texto a ser definido</param>
        public static void SetValueInputBox(this HtmlElement instance, string value)
        {
            instance.SetAttribute(VALUE_HTML_ATTRIBUTE, value);
        }

        /// <summary>
        /// Chama o método javascript click de um determinado elemento
        /// </summary>
        /// <param name="instance">Elemento a ser clicado</param>
        public static void ElementClick(this HtmlElement instance)
        {
            instance.InvokeMember("click");
        }

        /// <summary>
        /// Dispara o evento javascript onchange de um determinado elemento
        /// </summary>
        /// <param name="instance">elemento a ser disparado</param>
        public static void ElementChange(this HtmlElement instance)
        {
            instance.InvokeMember("fireEvent", new object[] { "onchange" });
        }


        /// <summary>
        /// Recupera uma lista de elementos com possuem um determinado atributo
        /// </summary>
        /// <param name="instance">Elemento DOM Pai</param>
        /// <param name="attributeName">Nome do atributo</param>
        /// <param name="value">Valor do atributo</param>
        /// <returns>Lista de elementos encontrados</returns>
        public static List<HtmlElement> GetElementsByAttribute(this HtmlElement instance, string attributeName, string value)
        {
            var result = new List<HtmlElement>();
            if (instance.CanHaveChildren)
            {
                result.AddRange(GetElementsByAttribute(instance.Children, attributeName, value, 0));
            }
            return result;
        }

        /// <summary>
        /// Recupera uma lista de elementos com possuem um determinado atributo a partir do elemento raiz
        /// </summary>
        /// <param name="instance">Elemento Raiz</param>
        /// <param name="attributeName">Nome do atributo</param>
        /// <param name="value">Valor do atributo</param>
        /// <returns>Lista de elementos encontrados</returns>
        public static List<HtmlElement> GetElementsByAttribute(this HtmlDocument instance, string attributeName, string value)
        {
            var result = new List<HtmlElement>();

            var elemColl = instance.All;
            result.AddRange(GetElementsByAttribute(elemColl, attributeName, value, 0));
           
            return result;
        }

        /// <summary>
        /// Recupera uma lista de elementos com possuem um determinado atributo a partir de outra lista de elementos
        /// </summary>
        /// <param name="instance">lista de elementos origem</param>
        /// <param name="attributeName">Nome do atributo</param>
        /// <param name="value">Valor do atributo</param>
        /// <returns>Lista de elementos encontrados</returns>
        public static List<HtmlElement> GetElementsByAttribute(this HtmlElementCollection instance, string attributeName, string value)
        {
            var result = new List<HtmlElement>();

            result.AddRange(GetElementsByAttribute(instance, attributeName, value, 0));

            return result;
        }

        private static List<HtmlElement> GetElementsByAttribute(HtmlElementCollection elemColl, string attributeName, string value, Int32 depth)
        {
            var result = new List<HtmlElement>();
            foreach (HtmlElement e in elemColl)
            {
                if (e.OuterHtml != null && (e.OuterHtml.ToLower().Contains($"{attributeName.ToLower()}='{value.ToLower()}'")
                  || e.OuterHtml.ToLower().Contains($"{attributeName.ToLower()}=\"{value.ToLower()}\"")
                  || e.OuterHtml.ToLower().Contains($"{attributeName.ToLower()}={value.ToLower()}")))
                    result.Add(e);

                if (e.CanHaveChildren)
                {
                    GetElementsByAttribute(e.Children, attributeName, value, depth + 1);
                }

            }
            return result;
        }

        /// <summary>
        /// Exibe um Alert Javascript com um determinado texto em um documento HTML
        /// </summary>
        /// <param name="instance">Elemento Raiz DOM</param>
        /// <param name="text">Texto do Alert</param>
        public static void Alert(this HtmlDocument instance, string text)
        {
            instance.Window?.Alert(text);
        }

        /// <summary>
        /// Executa um bloco de código Javascript em um documento DOM
        /// </summary>
        /// <param name="instance">Documento DOM</param>
        /// <param name="javascriptCode">Block de código Javascript a ser interpretado e executado</param>
        public static void RunJavascript(this HtmlDocument instance, string javascriptCode)
        {
            var scriptEl = instance.CreateElement("script");
            var element = (IHTMLScriptElement)scriptEl?.DomElement;
            if (element != null) element.text = javascriptCode;

            if (instance.Body != null)
            {
                instance.Body.AppendChild(scriptEl ?? throw new InvalidOperationException());
            }
            else
            {
                var head = instance.GetElementsByTagName("head")[0];
                head.AppendChild(scriptEl ?? throw new InvalidOperationException());
            }
        }


        /// <summary>
        /// Seleciona um item de um dropdown a partir do índice
        /// </summary>
        /// <param name="instance">Elemento DOM Dropdown</param>
        /// <param name="index">Índice do item</param>
        public static void SelectDropdownElement(this HtmlElement instance, int index)
        {
            instance.SetAttribute(SELECTED_INDEX_HTML_ATTRIBUTE, index.ToString());
            instance.ElementChange();
        }
    }
}

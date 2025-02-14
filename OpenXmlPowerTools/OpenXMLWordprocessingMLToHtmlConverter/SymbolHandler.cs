using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Default handler that transforms every symbol into some html encoded font specific char
    /// </summary>
    public class SymbolHandler : ISymbolHandler
    {
        /// <summary>
        /// Default handler that transforms every symbol into some html encoded font specific char
        /// </summary>
        /// <param name="element"></param>
        /// <param name="fontFamily"></param>
        /// <returns></returns>
        public XElement TransformSymbol(XElement element, Dictionary<string, string> fontFamily)
        {
            var cs = (string?)element.Attribute(W._char);
            string character = !string.IsNullOrEmpty(cs)
                ? Convert.ToChar(Convert.ToUInt16(cs, 16)).ToString()
                : "\ufffd"; // Replacement character
            return new XElement(Xhtml.span, new XText(character));
        }
    }
}
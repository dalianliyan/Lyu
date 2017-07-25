// Copyright (c) 2005 J.Keuper (j.keuper@gmail.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to 
// deal in the Software without restriction, including without limitation the 
// rights to use, copy, modify, merge, publish, distribute, sublicense, and/or 
// sell copies of the Software, and to permit persons to whom the Software is 
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in 
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
// IN THE SOFTWARE.


using System;
using System.Windows.Forms;
using System.Xml;

using WixEdit.Xml;

namespace WixEdit.PropertyGridExtensions {
    public class SearchElementObject {
        private XmlNode _searchNode;
        private WixFiles _wixFiles;

        public SearchElementObject(XmlNode searchNode, WixFiles wixFiles) {
            _searchNode = searchNode;
            _wixFiles = wixFiles;
        }

        public override string ToString() {
            XmlAttribute idAtt = _searchNode.Attributes["Id"];
            if (idAtt != null) {
                return String.Format("<< {0} : '{1}'>>", _searchNode.Name, idAtt.Value);
            } else {
                return String.Format("<< {0} >>", _searchNode.Name);
            }
        }

        public XmlNode Element {
            get { return _searchNode; }
        }

        public WixFiles WixFiles {
            get { return _wixFiles; }
        }
    }
}

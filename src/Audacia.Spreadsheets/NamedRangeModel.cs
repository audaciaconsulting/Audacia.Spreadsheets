using System;
using System.Collections.Generic;
using System.Text;

namespace Audacia.Spreadsheets
{
    public class NamedRangeModel
    {
        private string _name;
        /// <summary>
        /// Name for Named Range. Must only contain Alpha Numerics
        /// </summary>
        public string Name {
            get
            {
                return _name;
            }
            set
            {
                _name = value.Replace(' ', '_');
            }
        }
        public string SheetName { get; set; }
        public string StartCell { get; set; }
        public string EndCell { get; set; }

        public void Write()
        {

        }
    }
}

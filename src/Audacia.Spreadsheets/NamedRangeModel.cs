using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

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
                Regex r = new Regex("^[a-zA-Z0-9_]*$");
                if (r.IsMatch(value))
                {
                    _name = value;
                }
                else
                {
                    throw new Exception("Named Range names must be alphanumeric");
                }
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

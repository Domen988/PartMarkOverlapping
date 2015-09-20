using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Text.RegularExpressions;

using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;

namespace PartMarkOverlapping
{
    /// <summary>
    /// contains essential information about part: id, prefix, number and a bool (needs to be changed).
    /// </summary>
    /// <remarks>
    /// This class is needed because tekla does not provide 'get position number' - you can get part position ("B/2"), part prefix ("B"), but no number.
    /// </remarks>
    class PartCustom
    {
        // variables
        private TS.Identifier identifier;
        private string prefix;
        private int number;
        private bool needsToChange;
        private bool isMainPart;

        public static Dictionary<string, List<int>> positionsDictionary = new Dictionary<string, List<int>>();

        /// <summary>
        /// Custom part constructor - PartCustom can only be instantiated with a TeklaStructuresModel.Part object.
        /// </summary>
        /// <param name="part"></param>
        public PartCustom(TSM.Part part)
        {
            string partMark = part.GetPartMark();
            string partPrefix = string.Empty;
            string assemblyPrefix = string.Empty;
            string actualPartNumber = string.Empty;
            string actualPartPrefix = string.Empty;
            int isMainPart = 0;

            part.GetReportProperty("PART_PREFIX", ref partPrefix);
            part.GetReportProperty("ASSEMBLY_PREFIX", ref assemblyPrefix);

            // next conditinals scoop out prefix and position numbers without hardcoding the 'part position separator' (it's an option in tekla)
            if (partMark.Contains(partPrefix))
            {
                actualPartNumber = Regex.Replace(partMark.Remove(0, partPrefix.Length), "[^.0-9]", "");
                actualPartPrefix = partPrefix;
            }
            else if (partMark.Contains(assemblyPrefix))
            {
                actualPartNumber = Regex.Replace(partMark.Remove(0, assemblyPrefix.Length), "[^.0-9]", "");
                actualPartPrefix = assemblyPrefix;
            }

            //MessageBox.Show(partPrefix + " + " + assemblyPrefix + " + " + partMark + "\n" + currentPartPrefix + " + " + currentPartNumber);

            this.identifier = part.Identifier;
            this.prefix = actualPartPrefix;
            this.number = Int32.Parse(actualPartNumber);
            this.needsToChange = false;
            part.GetReportProperty("MAIN_PART", ref isMainPart);
            this.isMainPart = Convert.ToBoolean(isMainPart);

            // gradually builds up positionsDictionary, which is a dictionary where keys are prefixes and values are lists of position numbers.
            if (positionsDictionary.ContainsKey(this.prefix))
            {
                if (!positionsDictionary[this.prefix].Contains(this.number))
                {
                    positionsDictionary[this.prefix].Add(this.number);
                }
            }
            else
            {
                List<int> list = new List<int>();
                list.Add(this.number);
                positionsDictionary.Add(this.prefix, list);
            }
        }

        /// <summary>
        /// property that gives access to variables.
        /// </summary>
        public bool NeedsToChange
        {
            get { return needsToChange; }
            set { needsToChange = value; }
        }
        public TS.Identifier Identifier
        {
            get { return identifier; }
            set { identifier = value; }
        }
        public string Prefix
        {
            get { return prefix; }
            set { prefix = value; }
        }
        public int Number
        {
            get { return number; }
            set { number = value; }
        }
        public bool IsMainPart
        {
            get { return isMainPart; }
        }
    }
}

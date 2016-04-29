using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;

using System.Text.RegularExpressions;

using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;
using TeklaMacroBuilder;

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

        private static string modelPath;
        public static Dictionary<string, List<int>> positionsDictionary = new Dictionary<string, List<int>>();
        public static List<PartCustom> partList = new List<PartCustom>();
        public static Dictionary<string, Tuple<string, int>> prefixChanges = new Dictionary<string, Tuple<string, int>>();

        /// <summary>
        /// Custom part constructor. Adds the part to the partList and also creates positionsDictionary.
        /// </summary>
        /// <param name="part">TeklaStructuresModel.Part object</param>
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
                actualPartNumber = Regex.Replace(partMark.Remove(0, partPrefix.Length), "[^0-9]", "");
                actualPartPrefix = partPrefix;
            }
            else if (partMark.Contains(assemblyPrefix))
            {
                actualPartNumber = Regex.Replace(partMark.Remove(0, assemblyPrefix.Length), "[^0-9]", "");
                actualPartPrefix = assemblyPrefix;
            }

            //MessageBox.Show(partPrefix + " + " + assemblyPrefix + " + " + partMark + "\n" + currentPartPrefix + " + " + currentPartNumber);

            this.identifier = part.Identifier;
            this.prefix = actualPartPrefix;
            this.number = int.Parse(actualPartNumber);
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

            // add part to part list - list of all model parts
            partList.Add(this);
        }

        /// <summary>
        /// Selects every object in the model.
        /// </summary>
        /// <param name="model">Tekla structures model</param>
        public static void SelectAll(TSM.Model model)
        {
            modelPath = model.GetInfo().ModelPath;

            TSM.ModelObjectEnumerator selectedObjects = model.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.UNKNOWN);
            System.Type[] objectTypes = new System.Type[1];
            objectTypes.SetValue(typeof(TSM.Part), 0);

            selectedObjects = model.GetModelObjectSelector().GetAllObjectsWithType(objectTypes);

            // clear lists - needed for subsequent numbering checks
            positionsDictionary.Clear();
            partList.Clear();
            prefixChanges.Clear();

            // creating instances of PartCustom class
            while (selectedObjects.MoveNext())
            {
                TSM.Part part = selectedObjects.Current as TSM.Part;
                PartCustom currentPart = new PartCustom(part);
            }
        }

        /// <summary>
        /// searches for parts that need to change part number
        /// </summary>
        /// <param name="part"></param>
        public static void FindPartsToBeRenumbered()
        {
            foreach (PartCustom part in partList)
            {
                // check if current part is secondary 
                if (!part.IsMainPart)
                {
                    string oppositePrefix = ChangeCapitalization(part.Prefix);
                    // check if main parts with same letter exist
                    if (positionsDictionary.ContainsKey(oppositePrefix))
                    {
                        // check if same number is used for main and secondaries
                        if (positionsDictionary[oppositePrefix].Contains(part.Number))
                        {
                            part.NeedsToChange = true;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// select all parts that need to be renumbered
        /// </summary>
        /// <param name="model"></param>
        /// <returns>quantity of selected parts</returns>
        public static int SelectPartsToBeRenumbered(TSM.Model model)
        {
            ArrayList partsToBeSelected = new ArrayList();

            foreach (PartCustom part in partList)
            {
                if (part.NeedsToChange)
                {
                    partsToBeSelected.Add(model.SelectModelObject(part.Identifier));
                }
            }
            TSM.UI.ModelObjectSelector mos = new TSM.UI.ModelObjectSelector();
            mos.Select(partsToBeSelected);
            return mos.GetSelectedObjects().GetSize();
        }

        /// <summary>
        /// Renumbers part that need renumbering.
        /// </summary>
        /// <param name="model">Tekla.Structures.Model model</param>
        public static void RenumberParts(TSM.Model model)
        {
            // check if numberinghistory.txt exists and change its name
            RenameNumberingHistory(model);

            foreach (PartCustom part in partList)
            {
                if (part.NeedsToChange)
                {
                    string partCurrentPosition;
                    partCurrentPosition = part.prefix.ToString() + part.number.ToString();

                    int newNum;
                    string oppositePrefix = ChangeCapitalization(part.Prefix);

                    // checks if a part with same position has already been assigned a new number. 
                    // If so, it skips it --> tekla applies the new number from the first part to all of the same parts
                    // all changes are collected in prefixChanges dictionary.
                    if (!prefixChanges.ContainsKey(partCurrentPosition))
                    {
                        bool firstGo = true;
                        string preNumber = string.Empty;
                        string postNumber = string.Empty;

                        do
                        {
                            int maxOppositeNumber = positionsDictionary[oppositePrefix].Max();
                            int maxNumber = positionsDictionary[part.Prefix].Max();

                            newNum = Math.Max(maxNumber, maxOppositeNumber) + 1;

                            // adds new number to prefixChanges dictionary
                            Tuple<string, int> tuple = new Tuple<string, int>(part.prefix, newNum);

                            if (firstGo)
                            {
                                prefixChanges.Add(partCurrentPosition, tuple);
                            }
                            else
                            {
                                prefixChanges.Remove(partCurrentPosition);
                                prefixChanges.Add(partCurrentPosition, tuple);
                            }

                            // select part - clumsy, could it be improved?
                            ArrayList aList = new ArrayList();
                            TSM.Object tPart = model.SelectModelObject(part.Identifier);
                            TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
                            aList.Add(tPart);
                            selector.Select(aList);

                            TSM.Part myPart = tPart as TSM.Part;

                            // preNumber and postNumber strings are compared in the 'while' of the do-while loop, to determine if Macrobuilder 
                            // macro was succesfully run. 
                            // (sometimes Tekla doesn't want to apply certain numbers - e.g.: if they were in use in previous model stages, ... )
                            preNumber = myPart.GetPartMark();

                            // use Macrobuilder dll to change numbering
                            MacroBuilder macroBuilder = new MacroBuilder();
                            macroBuilder.Callback("acmdAssignPositionNumber", "part", "main_frame");
                            macroBuilder.ValueChange("assign_part_number", "Position", newNum.ToString());
                            macroBuilder.PushButton("AssignPB", "assign_part_number");
                            macroBuilder.PushButton("CancelPB", "assign_part_number");
                            macroBuilder.Run();

                            postNumber = myPart.GetPartMark();

                            bool ismacrounning = true;
                            while (ismacrounning)
                            {
                                ismacrounning = TSM.Operations.Operation.IsMacroRunning();
                            }

                            // add newly created part mark to positionsDict
                            positionsDictionary[part.Prefix].Add(newNum);

                            firstGo = false;

                        }
                        //while (!AssignmentSuccesCheck(model));
                        while (preNumber == postNumber);
                    }
                }
            }
        }

        /// <summary>
        /// changes strings with one capital letter to all lower case and strings without capital letter to all upper case.
        /// </summary>
        /// <param name="str"></param>
        /// <returns>returns transformed string</returns>
        internal static string ChangeCapitalization(string str)
        {
            string result = "";
            if (string.IsNullOrEmpty(str))
            {
                result = "";
                return result;
            }
            for (int i = 0; i < str.Length; i++)
            {
                if (char.IsUpper(str[i]))
                {
                    result = str.ToLower();
                    break;
                }
                result = str.ToUpper();
            }
            return result;
        }

        /// <summary>
        /// adds datetime string to the old numberinghistory.txt file in order to speed up the overlapping correction
        /// </summary>
        internal static void RenameNumberingHistory(TSM.Model model)
        {            
            string formatedDate = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");

            string numHisOld = modelPath + "\\numberinghistory.txt";
            string numHisNew = modelPath + "\\numberinghistory until " + formatedDate + ".txt";

            try
            {
                if (File.Exists(numHisOld))
                {
                    File.Move(numHisOld, numHisNew);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// OBSOLETE - checks if assignment of the new part number succeeded
        /// </summary>
        /// <param name="model"></param>
        /// <returns>true if yes, false if no</returns>
        internal static bool AssignmentSuccesCheck(TSM.Model model)
        {
            string[] lines = File.ReadAllLines(modelPath + "\\numberinghistory.txt");
            
            if (lines[lines.Length - 2] == "")
            {
                return false;
            }
            
            return true;    
        }

        /// <summary>
        /// property that gives access to needsToChange variable.
        /// </summary>
        public bool NeedsToChange
        {
            get { return needsToChange; }
            set { needsToChange = value; }
        }

        /// <summary>
        /// property that gives access to identifier variable.
        /// </summary>
        public TS.Identifier Identifier
        {
            get { return identifier; }
            set { identifier = value; }
        }

        /// <summary>
        /// property that gives access to prefix variable.
        /// </summary>
        public string Prefix
        {
            get { return prefix; }
            set { prefix = value; }
        }

        /// <summary>
        /// property that gives access to number variable.
        /// </summary>
        public int Number
        {
            get { return number; }
            set { number = value; }
        }

        /// <summary>
        /// property that gives access to isMainPart variable.
        /// </summary>
        public bool IsMainPart
        {
            get { return isMainPart; }
        }
    }
}

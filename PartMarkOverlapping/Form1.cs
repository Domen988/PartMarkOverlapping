using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;

using TeklaMacroBuilder;

namespace PartMarkOverlapping
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TSM.Model model = new TSM.Model();
            
            TSM.ModelObjectEnumerator selectedObjects = model.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.UNKNOWN);
            
            // select object types for selector 
            System.Type[] objectTypes = new System.Type[1];
            objectTypes.SetValue(typeof(TSM.Part), 0);

            selectedObjects = model.GetModelObjectSelector().GetAllObjectsWithType(objectTypes);

            // numbering check
            if (!CheckNumberingStatus(selectedObjects))
            {
                MessageBox.Show("Numbering is not up-to-date.", "Part Mark Overlapping", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            selectedObjects.Reset();

            // populating dictionary with id's, part prefix and part position number
            Dictionary<TS.Identifier, Tuple<string, int, bool>> partDict = new Dictionary<TS.Identifier, Tuple<string, int, bool>>();
            while (selectedObjects.MoveNext())
            {
                TSM.Part part = selectedObjects.Current as TSM.Part;
                string partMark = part.GetPartMark();
                string[] splittedMark = partMark.Split('/');
                int partPosNum = Int32.Parse(splittedMark[1]);
                Tuple<string, int, bool> partTuple = new Tuple<string, int, bool>(splittedMark[0], partPosNum, false);
                partDict.Add(part.Identifier, partTuple);
            }
            
            // creates dictionary by prefixes / lists of part numbers
            Dictionary<string, List<int>> positionsDict = new Dictionary<string, List<int>>();
            foreach(KeyValuePair<TS.Identifier, Tuple<string, int, bool>> entry in partDict)
            {
                if (positionsDict.ContainsKey(entry.Value.Item1))
                {
                    List<int> list = new List<int>();
                    list = positionsDict[entry.Value.Item1];
                    if (!list.Contains(entry.Value.Item2))
                    {
                        list.Add(entry.Value.Item2);
                    }
                    positionsDict[entry.Value.Item1] = list;
                }
                else
                {
                    List<int> list = new List<int>();
                    list.Add(entry.Value.Item2);
                    positionsDict.Add(entry.Value.Item1, list);
                }
            }

            // check the dictionary by prefixes and determine the parts that need new numbers
            foreach (KeyValuePair<TS.Identifier, Tuple<string, int, bool>> entry in partDict)
            {
                if (entry.Value.Item1 == "F" && entry.Value.Item2 == 10)
                {
                    var test = "";
                }
                // check if current part is secondary - if prefix has capital letters
                if (Char.IsUpper(entry.Value.Item1.ToString()[0]))
                {
                    string testKey = entry.Value.Item1.ToString().ToLower();
                    // check if main parts with same letter exist
                    if (positionsDict.ContainsKey(testKey))
                    {
                        // check if same number is used for main and secondaries
                        if (positionsDict[testKey].Contains(entry.Value.Item2))
                        {
                            int newNum = 0;
                            // loop searches for new number and ads it to positionsDict
                            while(true)
                            {
                                newNum++;
                                if (positionsDict[testKey].Contains(newNum) || positionsDict[entry.Value.Item1].Contains(newNum))
                                {
                                    continue;
                                }
                                if (entry.Value.Item1 == "F" && entry.Value.Item2 == 10)
                                {
                                    //MessageBox.Show(positionsDict[entry.Value.Item1].Contains(newNum).ToString());
                                }
                                else
                                {
                                    // select part - clumsy, could it be improved?
                                    ArrayList aList = new ArrayList();
                                    TSM.Object part = model.SelectModelObject(entry.Key);
                                    TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
                                    aList.Add(part);
                                    selector.Select(aList);

                                    // use Macrobuilder dll to change numbering
                                    new MacroBuilder().
                                        Callback("acmdAssignPositionNumber", "part", "main_frame").
                                        ValueChange("assign_part_number", "Position", newNum.ToString()).
                                        PushButton("AssignPB", "assign_part_number").
                                        PushButton("CancelPB", "assign_part_number").
                                        Run();

                                    if (entry.Value.Item1=="F" && newNum==10)
                                    {
                                        //continue;
                                    }
                                    // add newly created part mark to positionsDict
                                    positionsDict[entry.Value.Item1].Add(newNum);

                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>Checks if numbering is up to date for all objects in a list.</summary>
        /// <param name="objectList">The object list to check.</param>
        /// <returns>true if numbering is up to date for all objects in the object list,
        /// false if even one object is not up-to date.</returns>
        internal static bool CheckNumberingStatus(TSM.ModelObjectEnumerator objectList)
        {
            objectList.SelectInstances = false;
            while (objectList.MoveNext())
            {
                if (objectList.Current is TSM.Part ||
                    objectList.Current is TSM.Assembly ||
                    objectList.Current is TSM.Reinforcement)
                {
                    try
                    {
                        if (!TSM.Operations.Operation.IsNumberingUpToDate(objectList.Current))
                        {
                            return false;
                        }
                    }
                    catch (ArgumentException)
                    {
                    }
                }
            }
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var adOpt = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_USE_ASSEMBLY_NUMBER_FOR", ref adOpt);
            if (adOpt != "MAIN_PART")
            {
                MessageBox.Show("Advanced option 'XS_USE_ASSEMBLY_NUMBER_FOR' is not set to 'MAIN_PART'.", "Part Mark Overlapping", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}



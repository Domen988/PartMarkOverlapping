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
using System.Text.RegularExpressions;

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
            List<PartCustom> partDictionary = new List<PartCustom>();
            while (selectedObjects.MoveNext())
            {
                TSM.Part part = selectedObjects.Current as TSM.Part;
                PartCustom currentPart = new PartCustom(part);
                partDictionary.Add(currentPart);
            }


            // check the dictionary by prefixes and determine the parts that need new numbers
            // select parts in model
            ArrayList partList = new ArrayList();

            foreach (PartCustom part in partDictionary)
            {
                needsToChange(part);
                if (part.NeedsToChange)
                {
                    partList.Add(model.SelectModelObject(part.Identifier));
                }
            }
            Tekla.Structures.Model.UI.ModelObjectSelector mos = new Tekla.Structures.Model.UI.ModelObjectSelector();
            mos.Select(partList);

            // change parts and select them in model
            foreach (PartCustom part in partDictionary)
            {
                if (part.Prefix=="F"||part.Number==10)
                {
                    string nothing = "";
                }
                if (part.NeedsToChange)
                {
                    int newNum = 0;
                    // loop searches for new number and ads it to positionsDict
                    while (true)
                    {
                        newNum++;
                        string oppositePrefix = changeCapitalization(part.Prefix);
                        if (PartCustom.positionsDictionary[oppositePrefix].Contains(newNum) || PartCustom.positionsDictionary[part.Prefix].Contains(newNum))
                        {
                            continue;
                        }
                        // problem je nekje tu
                        else
                        {
                            // select part - clumsy, could it be improved?
                            ArrayList aList = new ArrayList();
                            TSM.Object tPart = model.SelectModelObject(part.Identifier);
                            TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
                            aList.Add(tPart);
                            selector.Select(aList);

                            string MacrosPath = "";
                            TS.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);
                            

                            // use Macrobuilder dll to change numbering
                            MacroBuilder macroBuilder = new MacroBuilder();
                            macroBuilder.Callback("acmdAssignPositionNumber", "part", "main_frame");
                            macroBuilder.ValueChange("assign_part_number", "Position", newNum.ToString());
                            macroBuilder.PushButton("AssignPB", "assign_part_number");
                            macroBuilder.PushButton("CancelPB", "assign_part_number");
                            macroBuilder.Run();

                            /*
                            new MacroBuilder().
                                Callback("acmdAssignPositionNumber", "part", "main_frame").
                                ValueChange("assign_part_number", "Position", newNum.ToString()).
                                PushButton("AssignPB", "assign_part_number").
                                PushButton("CancelPB", "assign_part_number").
                                Run();
                                */
                            bool ismacrounning = true;
                            while (ismacrounning)
                            {
                                ismacrounning = TSM.Operations.Operation.IsMacroRunning();
                            }
                            // add newly created part mark to positionsDict
                            PartCustom.positionsDictionary[part.Prefix].Add(newNum);


                            break;
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

        /// <summary>
        /// changes strings with one capital letter to all lower case and strings without capital letter to all upper case.
        /// </summary>
        /// <param name="str"></param>
        /// <returns>returns transformed string</returns>
        static string changeCapitalization(string str)
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
        /// checks if current part needs to change part number
        /// </summary>
        /// <param name="part"></param>
        static void needsToChange(PartCustom part)
        {
            // check if current part is secondary 
            if (!part.IsMainPart)
            {
                string oppositePrefix = changeCapitalization(part.Prefix);
                // check if main parts with same letter exist
                if (PartCustom.positionsDictionary.ContainsKey(oppositePrefix))
                {
                    // check if same number is used for main and secondaries
                    if (PartCustom.positionsDictionary[oppositePrefix].Contains(part.Number))
                    {
                        part.NeedsToChange = true;
                    }
                }
            }
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



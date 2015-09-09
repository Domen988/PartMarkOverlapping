using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Tekla.Structures;
using TSM = Tekla.Structures.Model;

namespace PartMarkOverlapping
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            var adOpt = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_USE_ASSEMBLY_NUMBER_FOR", ref adOpt);
            if (adOpt != "MAIN_PART")
            {
                MessageBox.Show("Advanced option 'XS_USE_ASSEMBLY_NUMBER_FOR' is not set to 'MAIN_PART'.\n\noverlapping will exit.", "Critical warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TSM.Model model = new TSM.Model();
            
            Tekla.Structures.Model.ModelObjectEnumerator selectedObjects = model.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.UNKNOWN);
            
            // select object types for selector 
            System.Type[] objectTypes = new System.Type[1];
            objectTypes.SetValue(typeof(TSM.Part), 0);

            // select all objects with types
            selectedObjects = model.GetModelObjectSelector().GetAllObjectsWithType(objectTypes);

            if (!CheckNumberingStatus(selectedObjects))
            {
                MessageBox.Show("Numbering is not up-to-date!", "Part Mark Overlapping", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            while (selectedObjects.MoveNext())
            {

                var currentObject = selectedObjects.Current;
                var nameOfObject = "";
                var profileOfObject = "";
                var prefixAssemblyOfObject = "";
                var prefixPartOfObject = "";
                bool isFlatProfile = false;
                // get name of the object
                currentObject.GetReportProperty("NAME", ref nameOfObject);
                MessageBox.Show(nameOfObject);

                // get the profile of the object
                currentObject.GetReportProperty("PROFILE", ref profileOfObject);

                // get the prefix of the object
                currentObject.GetReportProperty("ASSEMBLY_DEFAULT_PREFIX", ref prefixAssemblyOfObject);
                currentObject.GetReportProperty("PART_PREFIX", ref prefixPartOfObject);

                // check if profile is flat profile
                if (profileOfObject.StartsWith("FL") || profileOfObject.StartsWith("PL")) isFlatProfile = true;

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
    }
}



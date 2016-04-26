using System;
using System.Windows.Forms;

using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;

namespace PartMarkOverlapping
{
    public partial class Form1 : Form
    {
        TSM.Model model = new TSM.Model();
        private string caption = "Part Mark Overlapping v1.0";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int quantityOfPartsToBeRenumbered;

            if (!checkNumbering()) return;        // check if numbering is up to date and return if it is not
            checkAdvancedOptions();  // check if advanced options are set correctly

            // cursor wait symbol
            Cursor.Current = Cursors.WaitCursor;
            
            Application.DoEvents();

            PartCustom.SelectAll(model);
            PartCustom.FindPartsToBeRenumbered();
            quantityOfPartsToBeRenumbered = PartCustom.SelectPartsToBeRenumbered(model);

            Cursor.Current = Cursors.Default;

            if (quantityOfPartsToBeRenumbered > 0)
            {
                if (MessageBox.Show("There are " + quantityOfPartsToBeRenumbered + " parts that will get renumbered\n\nConfirm renumbering?", caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    // cursor wait symbol
                    Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    PartCustom.RenumberParts(model);
                    Cursor.Current = Cursors.Default;
                }
                else
                {
                    Environment.Exit(1);
                }
            }
            else
            {
                MessageBox.Show("There are no parts that need renumbering.", caption);
                Environment.Exit(1);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!model.GetConnectionStatus())
            {
                MessageBox.Show("Tekla is not open.", caption);
                Environment.Exit(1);
            }

            checkAdvancedOptions();
            checkNumbering();
        }

        private bool checkNumbering()
        {
            if (!TSM.Operations.Operation.IsNumberingUpToDateAll())
            {
                MessageBox.Show("Numbering is not up-to-date.", caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
                //Environment.Exit(1);
            }
            return true;
        }

        private void checkAdvancedOptions()
        {
            var useAssemblyNumberFor = "";
            TS.TeklaStructuresSettings.GetAdvancedOption("XS_USE_ASSEMBLY_NUMBER_FOR", ref useAssemblyNumberFor);
            if (useAssemblyNumberFor != "MAIN_PART")
            {
                MessageBox.Show("Advanced option 'XS_USE_ASSEMBLY_NUMBER_FOR' is not set to 'MAIN_PART'.", caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
        }
    }
}



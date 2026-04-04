using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace C3dProjects25
{
    public class ChecklistItemMaster
    {
        public bool IsApproved { get; set; }
        public string Description { get; set; }
        public List<string> Categories { get; set; }
    }

    public class ChecklistItem
    {
        public bool IsApproved { get; set; }
        public string Description { get; set; }
    }

    public partial class DrawChecklistWindow : Window
    {
        public ObservableCollection<ChecklistItem> Items { get; set; }
        private List<ChecklistItemMaster> MasterItems { get; set; }

        public DrawChecklistWindow(string docName)
        {
            InitializeComponent();
            txtDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
            txtJobNo.Text = System.IO.Path.GetFileNameWithoutExtension(docName);

            MasterItems = new List<ChecklistItemMaster>();
            
            // Populate Master List from PDF
            AddMasterItem("Statement of property line sources included if shown on map", "Boundary", "Trees");
            AddMasterItem("Two Benchmarks with Vertical Datum referenced", "Boundary", "Trees");
            AddMasterItem("Tree Names & Diameters referenced", "Trees", "Site Plan", "Drainage Plan");
            AddMasterItem("Tree locations with removal table and boxed indicators for trees to remain", "Site Plan", "Drainage Plan");
            AddMasterItem("Ensure that survey stakes align with established controls.", "Construction Layout Survey");
            AddMasterItem("Adequate horizontal and vertical control points are included", "Construction Layout Survey");
            AddMasterItem("Benchmarks are maintained at required intervals", "Construction Layout Survey");
            AddMasterItem("Legal, to-scale, signed by the registered land surveyor", "As Built / Record Survey");
            AddMasterItem("Labeled 'As Built'", "As Built / Record Survey");
            AddMasterItem("Street Address of property", "As Built / Record Survey");
            AddMasterItem("Legal description and Subdivision Name, Lot and Block Number, and Plat Book", "As Built / Record Survey", "Boundary");
            AddMasterItem("FEMA flood zones, with FEMA map references and map number", "As Built / Record Survey", "Boundary", "Site Plan", "Drainage Plan");
            AddMasterItem("Indicate Coastal Construction Control Line (Atlantic Beach)", "As Built / Record Survey", "Boundary");
            AddMasterItem("Indicate FFE for each grade-level (Reference elevations in NAVD 1988)", "As Built / Record Survey");
            AddMasterItem("Show all new and existing structures and impervious surfaces, including distance to property lines and other structures", "As Built / Record Survey", "Site Plan", "Drainage Plan");
            AddMasterItem("Minimum front, sides, and rear yard setbacks and distance to structures", "As Built / Record Survey", "Site Plan");
            AddMasterItem("Existing utility or drainage easement lines and dimensions, whether public or private", "As Built / Record Survey", "Topographic", "Site Plan", "Drainage Plan");
            AddMasterItem("Indicate Benchmark elevation and location", "As Built / Record Survey");
            AddMasterItem("Finished elevations at grade for all structures, high and low points and intervals every 10 feet", "As Built / Record Survey", "Topographic", "Drainage Plan");
            AddMasterItem("Indicate Site Drainage Plans", "As Built / Record Survey");
            AddMasterItem("Lot coverage calculated as a percent of lot coverage of all impervious features", "As Built / Record Survey");
            AddMasterItem("Buildings (existing or new, garages, balconies, decks, lanai etc.)", "As Built / Record Survey");
            AddMasterItem("Accessory structures (walkways, stoops, concrete/pave patios, pool decks/coping, sheds, a/c and pool equipment pads, etc.)", "As Built / Record Survey");
            AddMasterItem("Parking areas, including driveway length and width", "As Built / Record Survey");
            AddMasterItem("North Arrow required.", "Site Plan", "Drainage Plan");
            AddMasterItem("Elevations before and after at corners (no finished floor elevations required).", "Site Plan");
            AddMasterItem("Corner Elevations", "Drainage Plan");
            AddMasterItem("Table of impervious surface calculations (house, driveway, sidewalk, patio).", "Site Plan", "Drainage Plan");
            AddMasterItem("Driveway, sidewalk, and back patio shown and dimensioned.", "Site Plan", "Drainage Plan");
            AddMasterItem("Boundary survey or a sketch with lot lines and improvements.", "Site Plan", "Drainage Plan", "Boundary");
            AddMasterItem("Cross-sections showing proposed grading, centerline, and swales if applicable.", "Drainage Plan");
            AddMasterItem("Design flow calculations (before and after flow conditions: Q = CiA).", "Drainage Plan");
            AddMasterItem("Finished Floor Elevations (FFE) for structures to confirm positive drainage.", "Drainage Plan");

            Items = new ObservableCollection<ChecklistItem>();
            dgChecklist.ItemsSource = Items;
            
            cmbSurveyType.SelectedIndex = 0; // Select "All" by default
            RefreshList();
        }

        private void AddMasterItem(string desc, params string[] categories)
        {
            MasterItems.Add(new ChecklistItemMaster 
            { 
                Description = desc, 
                IsApproved = false, 
                Categories = new List<string>(categories) 
            });
        }

        private void CmbSurveyType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshList();
        }

        private void RefreshList()
        {
            if (Items == null || cmbSurveyType == null || MasterItems == null) return;
            
            // Save state of current items before clearing so checkboxes persist
            foreach (var currentItem in Items)
            {
                var match = MasterItems.FirstOrDefault(m => m.Description == currentItem.Description);
                if (match != null) match.IsApproved = currentItem.IsApproved;
            }

            Items.Clear();

            var selectedType = (cmbSurveyType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "All";

            foreach (var masterItem in MasterItems)
            {
                if (selectedType == "All" || masterItem.Categories.Contains(selectedType))
                {
                    Items.Add(new ChecklistItem 
                    { 
                        Description = masterItem.Description, 
                        IsApproved = masterItem.IsApproved 
                    });
                }
            }
        }

        private void BtnInsert_Click(object sender, RoutedEventArgs e)
        {
            // Sync final state before closing
            foreach (var currentItem in Items)
            {
                var match = MasterItems.FirstOrDefault(m => m.Description == currentItem.Description);
                if (match != null) match.IsApproved = currentItem.IsApproved;
            }
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}

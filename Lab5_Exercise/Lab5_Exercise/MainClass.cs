using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using wf = System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.DocumentParts;


namespace Lab5_Exercise
{
    [PluginAttribute("Lab5_Exec", "TwentyTwo",DisplayName = "Lab5_Exec", ToolTip = "Lab5_exercise project")]

    public class MainClass : AddInPlugin
    {
        // implement execute method 
        public override int Execute(params string[] parameters)
        {         
            // current document
            Document doc = Application.ActiveDocument;
            // display message
            StringBuilder message = new StringBuilder();

            // SearchBox form
            SearchBox sb = new SearchBox();
            // show form
            sb.ShowDialog();
            // assign Id value
            int Id = 0;

            try
            {
                // convert sb return value string to int
                Id = Convert.ToInt32(sb.ReturnValue);
            }
            catch { }

            // create search object
            Search search = new Search();
            // selection to search
            search.Selection.SelectAll();

             // create SearchCondition by InternalName of Category & Property
             // to find the specific item by its ID
            SearchCondition condition = SearchCondition.HasPropertyByName("LcRevitData_Element",
                 "LcRevitPropertyElementId").EqualValue(new VariantData(Id));
 
            /*
            // create SearchCondition by DisplayName of Category & Property
            // to find the specific item by its ID
            SearchCondition condition = SearchCondition.HasPropertyByDisplayName("Element",
                "Id").EqualValue(new VariantData(Id));
           */

            // SearchCondition to applied during search
            search.SearchConditions.Add(condition);
            // collect model item (if found)
            ModelItem item = search.FindFirst(doc,false);

            // item found
            if (item != null)
            {
                // make selection
                doc.CurrentSelection.Add(item);
                // get modelitem's Element category by display name method
                PropertyCategory elementCategory = item.PropertyCategories.
                    FindCategoryByDisplayName("Element");
                // all properties of Element category
                DataPropertyCollection dataProperties = elementCategory.Properties;               
                // display properties count 
                message.Append(String.Format("[{0}] ModelItem's Element Category has {1} Properties.\n",
                    Id.ToString(), dataProperties.Count));
                // index
                int index = 1;
                // iterate properties
                foreach (DataProperty dp in elementCategory.Properties)
                {
                    // append to display "Property Display Name & Property Value(includes DataType)"
                    message.Append(String.Format("{0}. {1} => {2}\n", index, dp.DisplayName, dp.Value));
                    // index increment
                    index += 1;
                    
                }
                // display message
                wf.MessageBox.Show(message.ToString(), "Element Category");
            }
            // return value
            return 0;          
        }

    }
}

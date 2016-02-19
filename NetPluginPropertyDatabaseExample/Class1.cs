using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;


using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
 
namespace NetPluginPropertyDatabaseExample
{
    [Plugin("NetPluginPropertyDatabaseExample", "COINS", DisplayName = "Net Plugin Property Database Example")]
    public class WorkingClass : EventWatcherPlugin

    {
        private CADOLink m_dblink;

        public override void OnLoaded()
        {
            //The plugin will be loaded as soon as is possible in the GUI
            //Autodesk.Navisworks.Api.Application.ActiveDocumentChanged += Application_ActiveDocumentChanged;

            // On startup the application will ask for 
            // a suitable access database.
            m_dblink = new CADOLink();
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "MDB files|*.mdb";
            fileDialog.Title = "Access Database";
            DialogResult result = fileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                string filename = fileDialog.FileName;
                m_dblink.connect(filename);
            }

            Autodesk.Navisworks.Api.Application.ActiveDocumentChanged += Application_ActiveDocumentChanged;

        }
        public override void OnUnloading()
        {
            //The plugin is unloaded at the end of the Navisworks session
            //Autodesk.Navisworks.Api.Application.ActiveDocumentChanged -= Application_ActiveDocumentChanged; 

            Autodesk.Navisworks.Api.Application.ActiveDocumentChanged -= Application_ActiveDocumentChanged;      
       


        }

        /// simple event handler for Application.GuiCreated    
        void Application_ActiveDocumentChanged(object sender,
                                        System.EventArgs e)
        {
            //MessageBox.Show("ActiveDocumentChanged");
            Autodesk.Navisworks.Api.Application.ActiveDocument.Models.
                CollectionChanged += ActiveDocument_Models_CollectionChanged;
        }
         
       
         void ActiveDocument_Models_CollectionChanged(object sender,
                                                 System.EventArgs e)
        {
            //MessageBox.Show("ActiveDocument_Models_CollectionChanged");

            if(sender != null)
            {
                Document oDoc = sender as Document;
                if (oDoc.Title == "gatehouse.nwd")
                {
                    //make sure to subscribe the event one time only
                    oDoc.CurrentSelection.Changed -= CurrentSelection_Changed;
                    oDoc.CurrentSelection.Changed += CurrentSelection_Changed;

                    oDoc.CurrentSelection.Changing -= CurrentSelection_Changing;
                    oDoc.CurrentSelection.Changing += CurrentSelection_Changing;

                    oDoc.FileSaving -= oDoc_FileSaving;
                    oDoc.FileSaving += oDoc_FileSaving;
                }
            }

        }

          string mytabname = "MyCustomTabUserName";
          void oDoc_FileSaving(object sender,
                          System.EventArgs e)
         {
             
             //in theory, since we have removed the tab timely in selection changing, 
             // there should be only one node that still contains the custom tab. this node is the last selected node.
             //so get it out and remove the properties. 

             //in case there are more nodes, this code can also remove their custom tabs.

             if (sender != null)
             {
                 Document oDoc = sender as Document;

                 if (oDoc.Title == "gatehouse.nwd")
                 {
                     try
                     {
                         //firstly use .NET API to get the items with custom tab 
                         Search search = new Search();
                         search.Selection.SelectAll();
                         search.SearchConditions.Add(SearchCondition.HasCategoryByDisplayName(mytabname));
                         ModelItemCollection items = search.FindAll(oDoc, false);

                         ComApi.InwOpState9 oState = ComApiBridge.State;

                         foreach (ModelItem oitem in items)
                         {

                             //convert .NET items to COM items
                             ComApi.InwOaPath3 oPath = ComApiBridge.ToInwOaPath(oitem) as ComApi.InwOaPath3;

                             if ((oPath.Nodes().Last() as ComApi.InwOaNode).IsLayer)
                             {
                                 //check whether the custom property tab has been added. 
                                 int customProTabIndex = 1;
                                 ComApi.InwGUIPropertyNode2 nodePropertiesOwner = oState.GetGUIPropertyNode(oPath, true) as ComApi.InwGUIPropertyNode2;
                                 ComApi.InwGUIAttribute2 customTab = null;
                                 foreach (ComApi.InwGUIAttribute2 nwAtt in nodePropertiesOwner.GUIAttributes())
                                 {
                                     if (!nwAtt.UserDefined) continue;

                                     if (nwAtt.ClassUserName == mytabname)
                                     {
                                         //remove the custom tab
                                         nodePropertiesOwner.RemoveUserDefined(customProTabIndex);
                                         customTab = nwAtt;
                                         break;
                                     }
                                     customProTabIndex += 1;
                                 }
                             }
                         }
                     }
                     catch (Exception ex)
                     {
                         MessageBox.Show(ex.ToString());
                     }

                 }
             }
         }

         void CurrentSelection_Changing(object sender,
                                      System.EventArgs e)
         {
             ////it looks there is issue with this event. It does not fire for the higher levels of modelitems. 
             // reserve this workflow. I am checking with engineer team

             if (sender != null)
             {
                 Document oDoc = sender as Document;

                 if (oDoc.Title == "gatehouse.nwd")
                 {
                     //this is the old selection before selecting
                     if (oDoc.CurrentSelection.SelectedItems.Count > 0)
                     {
                         ComApi.InwOpState9 oState = ComApiBridge.State;

                         try
                         {
                             ComApi.InwOaPath3 oPath = oState.CurrentSelection.Paths()[1];
                             if ((oPath.Nodes().Last() as ComApi.InwOaNode).IsLayer)
                             {
                                 //check whether the custom property tab has been added.

                               
                                 int customProTabIndex = 1;
                                 ComApi.InwGUIPropertyNode2 nodePropertiesOwner = oState.GetGUIPropertyNode(oPath, true) as ComApi.InwGUIPropertyNode2;
                                 ComApi.InwGUIAttribute2 customTab = null;
                                 foreach (ComApi.InwGUIAttribute2 nwAtt in nodePropertiesOwner.GUIAttributes())
                                 {
                                     if (!nwAtt.UserDefined) continue;

                                     if (nwAtt.ClassUserName == mytabname)
                                     {
                                         //remove the custom tab
                                         nodePropertiesOwner.RemoveUserDefined(customProTabIndex);
                                         customTab = nwAtt;
                                         break;
                                     }
                                     customProTabIndex += 1;
                                 }
                             } 
                         }
                         catch (Exception ex)
                         {
                             MessageBox.Show(ex.ToString());
                         }
                     }
                 }
             }
         }


        void CurrentSelection_Changed(object sender,
                                      System.EventArgs e)
        {
            if (sender != null)
            {
                Document oDoc = sender as Document;


                if (oDoc.Title == "gatehouse.nwd")
                {
                    //this is the new selection after selecting
                    if (oDoc.CurrentSelection.SelectedItems.Count > 0)
                    {
                        ComApi.InwOpState9 oState = ComApiBridge.State;

                        try
                        {
                            ComApi.InwOaPath3 oPath = oState.CurrentSelection.Paths()[1];
                            if ((oPath.Nodes().Last() as ComApi.InwOaNode).IsLayer)
                            {
                                //check whether the custom property tab has been added.

                                string mytabname = "MyCustomTabUserName";
                                int customProTabIndex = 1;
                                ComApi.InwGUIPropertyNode2 nodePropertiesOwner = oState.GetGUIPropertyNode(oPath, true) as ComApi.InwGUIPropertyNode2;
                                ComApi.InwGUIAttribute2 customTab = null;
                                foreach (ComApi.InwGUIAttribute2 nwAtt in nodePropertiesOwner.GUIAttributes())
                                {
                                    if (!nwAtt.UserDefined) continue;

                                    if (nwAtt.ClassUserName == mytabname)
                                    {
                                        customTab = nwAtt;
                                        break;
                                    }
                                    customProTabIndex += 1;
                                }


                                if (customTab == null)
                                {
                                    ////create the tab if it does not exist
                                    ComApi.InwOaPropertyVec newPvec =
                                         (ComApi.InwOaPropertyVec)oState.ObjectFactory(
                                               ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                                    ComApi.InwOaProperty prop1 = oState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty) as ComApi.InwOaProperty;
                                    ComApi.InwOaProperty prop2 = oState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty) as ComApi.InwOaProperty;
                                    prop1.name = "Finish";
                                    prop1.UserName = "Finish Date";
                                    object linkVal = m_dblink.read("Finish", (oPath.Nodes().Last() as ComApi.InwOaNode).UserName);
                                    if (linkVal != null)
                                    {
                                        prop1.value = linkVal;
                                        newPvec.Properties().Add(prop1);
                                    }

                                    prop2.name = "Notes";
                                    prop2.UserName = "Notes";
                                    linkVal = m_dblink.read("Notes", (oPath.Nodes().Last() as ComApi.InwOaNode).UserName);
                                    if (linkVal != null)
                                    {
                                        prop2.value = linkVal;
                                        newPvec.Properties().Add(prop2);
                                    }

                                    //the first argument is always 0 if adding a new tab
                                    nodePropertiesOwner.SetUserDefined(0, mytabname, mytabname, newPvec);
                                }
                                else
                                {
                                    ////update the properties in the tab with the new values from database
                                    ComApi.InwOaPropertyVec newPvec =
                                       (ComApi.InwOaPropertyVec)oState.ObjectFactory(
                                             ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                                    foreach (ComApi.InwOaProperty nwProp in customTab.Properties())
                                    {
                                        ComApi.InwOaProperty prop = oState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty) as ComApi.InwOaProperty;
                                         prop.name = nwProp.name;
                                        prop.UserName = nwProp.UserName;
                                        object linkVal = m_dblink.read(prop.name, (oPath.Nodes().Last() as ComApi.InwOaNode).UserName);
                                        if (linkVal != null)
                                        {
                                            prop.value = linkVal;
                                            newPvec.Properties().Add(prop);
                                        } 
                                    }
                                    nodePropertiesOwner.SetUserDefined(customProTabIndex, mytabname, mytabname, newPvec);

                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                    
                }
            }
        }
         

    }
}

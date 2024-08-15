from pyrevit import script, forms
from pyrevit import revit, DB, UI
from pyrevit import HOST_APP, EXEC_PARAMS
import Autodesk.Revit.DB as db
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('IronPython.wpf')
import sys

from System.Windows.Controls import ComboBox, ComboBoxItem

xamlfile_family_selection = script.get_bundle_file('SelectMEPFamilyType.xaml')

xamlfile_parameter_pairs_selection = script.get_bundle_file('SelectParams.xaml')

import wpf
from System import Windows


family_instance_parameters = []
family_instance_parameters_names = []
params_names_vs_params_objs_dict = {}
family_instances_of_selected_family = []

def isMEPInstanceInSpace(MEPInstance_elem, space_elem):

    mep_instance_space_id = MEPInstance_elem.Space.Id
    space_id = space_elem.Id

    return mep_instance_space_id == space_id

def composeSheetNumberName(sheet_item):

    return sheet_item.SheetNumber + "_" + sheet_item.Name



class FamilySelectionWindow(Windows.Window):

    def __init__(self):
        wpf.LoadComponent(self, xamlfile_family_selection)
        self.selected_value_first = None
        self.selected_value_second = None
        self.reversed_panel_schedule_template_dict = {}
        self.selected_sheet1 = None
        self.selected_sheet2 = None
        self.all_views_option = "All views"

    def btn_ok_clicked(self, sender, e):
        self.Close()
        self.changePanelScheduleTemplate()
        # ParameterPairsSelectionWindow(self.selected_value).ShowDialog()

    def populate_combo_box_family_types(self, doc):
        all_MEP_family_types = filtered_collector.OfClass(db.FamilySymbol).ToElements()

        all_sheets = DB.FilteredElementCollector(doc).OfClass(db.ViewSheet).WhereElementIsNotElementType().ToElements()

        # print(len(all_sheets), " sheets found")
        # print("SAMPLE SHEET", all_sheets[0])


        elem_class_filter_panel_schedule_views = db.ElementClassFilter(db.Electrical.PanelScheduleView)

        elem_class_filter_panel_schedule_templates = db.ElementClassFilter(db.Electrical.PanelScheduleTemplate)


        all_panel_schedule_views = DB.FilteredElementCollector(doc).WherePasses(elem_class_filter_panel_schedule_views).ToElements()

        all_panel_schedule_templates = DB.FilteredElementCollector(doc).WherePasses(elem_class_filter_panel_schedule_templates).ToElements()
        # all_panel_schedule_views = filtered_collector.OfClass(db.Electrical.PanelScheduleView)
        # print(all_panel_schedule_views)

        # initilaize list of panel schedule templates list

        panel_schedule_template_list = []

        sheets_dict = {sheet_item.Id:composeSheetNumberName(sheet_item) for sheet_item in all_sheets}

        self.reversed_sheets_dict = {value:key for (key,value) in sheets_dict.items()}

        

        # iterate over panel schedule views and extract panel schedule templates
        for panel_schedule_view in all_panel_schedule_views:
            panel_schedule_template = panel_schedule_view.GetTemplate()
            if panel_schedule_template is not None:
                panel_schedule_template_list.append(panel_schedule_template)
        
        panel_schedule_template_list = list(set(panel_schedule_template_list))

        panel_schedule_template_dict = {key:value for (key, value) in [(templ.Id, templ.Name) for templ in all_panel_schedule_templates if templ.IsBranchPanelSchedule]}

        self.reversed_panel_schedule_template_dict = {value:key for (key, value) in panel_schedule_template_dict.items()}

        # print("REVERSED DICT:", self.reversed_panel_schedule_template_dict)

        quantity_MEP_Elements = len(all_MEP_family_types)

        all_family_names = list(set([elem.Family.Name for elem in all_MEP_family_types]))

        # for family_name in all_family_names:
        #     combobox_item = ComboBoxItem()
        #     combobox_item.Content = family_name
        #     self.combo_family_type.Items.Add(combobox_item)

        for template_item_id in panel_schedule_template_dict.keys():
            combobox_item_first = ComboBoxItem()
            combobox_item_first.Content = panel_schedule_template_dict[template_item_id]
            self.combo_first_panel_schedule_template.Items.Add(combobox_item_first)
            combobox_item_second = ComboBoxItem()
            combobox_item_second.Content = panel_schedule_template_dict[template_item_id]
            self.combo_second_panel_schedule_template.Items.Add(combobox_item_second)


        # add option to ignore sheets
        
        combobox_item_all = ComboBoxItem()
        combobox_item_all.Content = self.all_views_option
        self.combo_sheet_selection_1.Items.Add(combobox_item_all)
        for sheet_item in all_sheets:
            combobox_item_third = ComboBoxItem()
            combobox_item_third.Content = composeSheetNumberName(sheet_item=sheet_item)
            self.combo_sheet_selection_1.Items.Add(combobox_item_third)
        self.reversed_sheets_dict.update({self.all_views_option: None})
        
    def comboBoxFirstPanelScheduleTemplate_SelectionChanged(self, sender, e):
        selected_item = self.combo_first_panel_schedule_template.SelectedItem
        if selected_item:
            self.selected_value_first = selected_item.Content
            # print(self.selected_value)
        else:
            print("No Item Selected")
    
    def comboBoxSecondPanelScheduleTemplate_SelectionChanged(self, sender, e):
        selected_item = self.combo_second_panel_schedule_template.SelectedItem
        if selected_item:
            self.selected_value_second = selected_item.Content
            # print(self.selected_value)
        else:
            print("No Item Selected")

    def comboBoxSheetSelection1_SelectionChanged(self, sender, e):
        selected_item = self.combo_sheet_selection_1.SelectedItem
        if selected_item:
            self.selected_sheet1 = selected_item.Content
            # print(self.selected_value)
        else:
            print("No Item Selected")

    def changePanelScheduleTemplate(self):

        # select all views of the selected PanelScheduleTemplate
        elem_class_filter_panel_schedule_views = db.ElementClassFilter(db.Electrical.PanelScheduleView)
        all_panel_schedule_views = DB.FilteredElementCollector(doc).WherePasses(elem_class_filter_panel_schedule_views).ToElements()
        panel_schedule_views_list_of_template_1 = []
        panel_schedule_views_list_of_template_2 = []

        if self.selected_sheet1 == self.all_views_option:
            placed_views_ids = None
        else:
            placed_views_ids = self.get_schedules([self.reversed_sheets_dict[self.selected_sheet1]])[1]

        if placed_views_ids is not None:
            all_panel_schedule_views = [view_item for view_item in all_panel_schedule_views if view_item.Id in placed_views_ids]



        if True:

            t = DB.Transaction(doc, "ChangePanelScheduleTemplatesLoadCalcsShowHide")
            t.Start()

            closed = False

            try: 

                for schedule_view in all_panel_schedule_views:

                    
                    template_of_schedule_view = doc.GetElement(schedule_view.GetTemplate())
                    if template_of_schedule_view.Name == self.selected_value_first:
                        panel_schedule_views_list_of_template_1.append(template_of_schedule_view)
                        schedule_view.GenerateInstanceFromTemplate(self.reversed_panel_schedule_template_dict[self.selected_value_second])

                        ## commented out to switch from first template to the second only and not from the second to the first - AE
                    #     continue
                    # if template_of_schedule_view.Name == self.selected_value_second:
                    #     panel_schedule_views_list_of_template_2.append(template_of_schedule_view)
                    #     schedule_view.GenerateInstanceFromTemplate(self.reversed_panel_schedule_template_dict[self.selected_value_first])
                
                print("Template 1", panel_schedule_views_list_of_template_1)
                print("Template 2", panel_schedule_views_list_of_template_2)

            except Exception as error:
                print("ERROR: ", error)
                t.Commit()
                self.Close()
                closed = True
                        # print("Match")
            if not closed:
                t.Commit()
                self.Close()

    def get_schedules(self, view_sheet_ids_list):

        # idea from Jeremy Tammik
        # https://thebuildingcoder.typepad.com/blog/2013/02/retrieving-schedules-on-a-sheet.html

        all_placed_views = []
        all_placed_views_ids = []

        for sheet_id in view_sheet_ids_list:

            view_sheet = doc.GetElement(sheet_id)
            view_doc = view_sheet.Document

            collector = DB.FilteredElementCollector(view_doc, view_sheet.Id)

            schedule_sheet_instances = collector.OfClass(db.Electrical.PanelScheduleSheetInstance).ToElements()

            print(len(schedule_sheet_instances), " SHEET INSTANCES")

            for schedule_sheet_instance in schedule_sheet_instances:
                schedule_id = schedule_sheet_instance.ScheduleId

                if schedule_id == db.ElementId.InvalidElementId:
                    continue

                view_schedule = doc.GetElement(schedule_id)

                if isinstance(view_schedule, db.Electrical.PanelScheduleView):
                    all_placed_views.append(view_schedule)
                    all_placed_views_ids.append(view_schedule.Id)
        print(all_placed_views)



        return all_placed_views, all_placed_views_ids





class ParameterPairsSelectionWindow(Windows.Window):

    def __init__(self, selected_value):
        wpf.LoadComponent(self, xamlfile_parameter_pairs_selection)

        self.selected_value = selected_value

        self.selected_family_instances_list = []

        self.selected_value_space_1 = None
        self.selected_value_MEP_Instance_1 = None

        self.selected_value_space_2 = None
        self.selected_value_MEP_Instance_2 = None

        self.selected_value_space_3 = None
        self.selected_value_MEP_Instance_3 = None

        self.selected_value_space_4 = None
        self.selected_value_MEP_Instance_4 = None

        MEP_family_name = self.selected_value
        # Collect all Families in the Document
        family_collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)
        # Collect all Spaces and Rooms in the Active View
        spaces_collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(DB.SpatialElement)

        # # DeBugging (commented)
        # for f in family_collector:
        #     print("Family_Name: ".format(f.Name))
        # print("Family Collector", list(family_collector.ToElements()))

        # Collect only Families which were selected from Combo Box in the First Popup Window
        MEP_families = [f for f in family_collector if f.Name == MEP_family_name]

        # Transform Filtered Collector to Elements
        spatial_elements = spaces_collector.ToElements()
       

        # # DeBugging (commented)
        # for spatial_family in spatial_elements:
        #     print("Spatial Element Type {}".format(spatial_family.SpatialElementType))

        # separate Spaces from Rooms
        self.spaces_elements = [spatial_elem for spatial_elem in spatial_elements if spatial_elem.SpatialElementType == DB.SpatialElementType.Space]
        print("So many spaces: {}".format(len(self.spaces_elements)))

        # populate Combo Boxes for Spaces
        if self.spaces_elements:
            self.spaces_params = list(self.spaces_elements[0].GetOrderedParameters())
            for space_param in self.spaces_params:
                space_param_name = space_param.Definition.Name
                combobox_item_spaces_1 = ComboBoxItem()
                combobox_item_spaces_2 = ComboBoxItem()
                combobox_item_spaces_3 = ComboBoxItem()
                combobox_item_spaces_4 = ComboBoxItem()
                combobox_item_spaces_1.Content = space_param_name
                combobox_item_spaces_2.Content = space_param_name
                combobox_item_spaces_3.Content = space_param_name
                combobox_item_spaces_4.Content = space_param_name
                self.combo_parameter_from_space_1.Items.Add(combobox_item_spaces_1)
                self.combo_parameter_from_space_2.Items.Add(combobox_item_spaces_2)
                self.combo_parameter_from_space_3.Items.Add(combobox_item_spaces_3)
                self.combo_parameter_from_space_4.Items.Add(combobox_item_spaces_4)
        else:
            print("No spaces elements. Something went wrong. EXIT PROGRAM")
            sys.exit()
            self.Close()


        if MEP_families:
            for MEP_family in MEP_families:
                # Get the ElementId of the selected family
                family_ids = list(MEP_family.GetFamilySymbolIds())

                # # DeBugging (commented)
                # print(type(family_ids))
                # print(family_ids)

                

                # Define the filter for FamilyInstances of the selected family
                for family_id in family_ids:

                    # setup family instance filter, it uses FamilySymbolIds
                    family_instance_filter = DB.FamilyInstanceFilter(doc, family_id)

                    # Get all family instances of the selected family
                    family_instance_collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id).WherePasses(family_instance_filter)
                    mep_family_Instances = family_instance_collector.ToElements()

                    self.selected_family_instances_list += mep_family_Instances

                    # family_instances_of_selected_family += mep_family_Instances

                    print("Len Family Instancies MEP: {}".format(len(mep_family_Instances)))

                    for family_instance in mep_family_Instances:
                        for param in family_instance.Parameters:
                            if param not in family_instance_parameters:
                                param_name = param.Definition.Name


                                # Populate param names dictionary, to find each param on its Definition.Name
                                if param_name not in params_names_vs_params_objs_dict:
                                    params_names_vs_params_objs_dict[param_name] = [param]
                                else:
                                    temp_list_params = params_names_vs_params_objs_dict[param_name]
                                    temp_list_params.append(param)
                                    params_names_vs_params_objs_dict[param_name] = temp_list_params
                                

                                # Populate Combo Boxes
                                if param_name not in family_instance_parameters_names:
                                    family_instance_parameters_names.append(param_name)
                                    combobox_item_1 = ComboBoxItem()
                                    combobox_item_2 = ComboBoxItem()
                                    combobox_item_3 = ComboBoxItem()
                                    combobox_item_4 = ComboBoxItem()
                                    combobox_item_1.Content = param_name
                                    combobox_item_2.Content = param_name
                                    combobox_item_3.Content = param_name
                                    combobox_item_4.Content = param_name
                                    self.combo_param_to_1.Items.Add(combobox_item_1)
                                    self.combo_param_to_2.Items.Add(combobox_item_2)
                                    self.combo_param_to_3.Items.Add(combobox_item_3)
                                    self.combo_param_to_4.Items.Add(combobox_item_4)

                    # print("So many Instances of {}: {} ".format(family_id, len(my_family_instances)))

                # Now you can process the family instances as needed
                for my_family_instance in mep_family_Instances:
                    # Do something with each family instance
                    pass
        else:
            print("Family {} not found.".format(MEP_family_name))

        

    def transfer_parameters(self, space_elem, MEP_instance, space_parameter_name, MEP_instance_parameter_name):

        if space_parameter_name is not None and MEP_instance_parameter_name is not None:

            space_parameter = space_elem.LookupParameter(space_parameter_name)
            MEP_parameter = MEP_instance.LookupParameter(MEP_instance_parameter_name)

            if (MEP_parameter is not None) and (space_parameter is not None):
                
                print("Transaction fire up: MEP instance param: {}; space param: {}".format(MEP_instance_parameter_name, space_parameter_name))
                MEP_parameter.Set(space_parameter.AsValueString())
                
                print("Transaction committed")
            else:
                print("MEP and/or space parameter do not exist")
        else:
            print("Parameter not filled")



    def btn_ok_clicked(self, sender, e):

        counter_matches = 0

        t = DB.Transaction(doc, "Param transfer CUSTOM")
        t.Start()

        closed = False

        try: 

            for space_elem in self.spaces_elements:
                for MEP_instance in self.selected_family_instances_list:
                    
                    phase = list(doc.Phases)[-1]  # retrieve the last phase of the project
                    # print("MEP Instance Space type: {}".format(type(MEP_instance.Space[phase])))
                    
                    # if MEP_instance.Space[phase] is not None:
                    #     mep_location_space_id = MEP_instance.Space[phase].Id
                    # else:
                    #     mep_location_space_id = MEP_instance.Space.Id
                    mep_location_point = MEP_instance.Location.Point
                    print("SPACE: ", space_elem)
                    if space_elem.IsPointInSpace(mep_location_point):
                        counter_matches += 1

                        self.transfer_parameters(space_elem=space_elem, MEP_instance=MEP_instance,
                                                space_parameter_name=self.selected_value_space_1, 
                                                MEP_instance_parameter_name=self.selected_value_MEP_Instance_1)
                        
                        self.transfer_parameters(space_elem=space_elem, MEP_instance=MEP_instance,
                                                space_parameter_name=self.selected_value_space_2, 
                                                MEP_instance_parameter_name=self.selected_value_MEP_Instance_2)
                        
                        self.transfer_parameters(space_elem=space_elem, MEP_instance=MEP_instance,
                                                space_parameter_name=self.selected_value_space_3, 
                                                MEP_instance_parameter_name=self.selected_value_MEP_Instance_3)
                        
                        self.transfer_parameters(space_elem=space_elem, MEP_instance=MEP_instance,
                                                space_parameter_name=self.selected_value_space_4, 
                                                MEP_instance_parameter_name=self.selected_value_MEP_Instance_4)
        except Exception as error:
            print("ERROR: ", error)
            t.Commit()
            self.Close()
            closed = True
                    # print("Match")
        if not closed:
            t.Commit()
            self.Close()
                    
        print("Number of matches: {}". format(counter_matches))
        
        
        self.Close()

    
    def combobox_item_spaces_1_SelectionChanged(self, sender, e):
        selected_item = self.combo_parameter_from_space_1.SelectedItem
        if selected_item:
            self.selected_value_space_1 = selected_item.Content
            
        else:
            print("No Item Selected")
    
    def combobox_item_spaces_2_SelectionChanged(self, sender, e):
        selected_item = self.combo_parameter_from_space_2.SelectedItem
        if selected_item:
            self.selected_value_space_2 = selected_item.Content
            print(self.selected_value)
        else:
            print("No Item Selected")

    def combobox_item_spaces_3_SelectionChanged(self, sender, e):
        selected_item = self.combo_parameter_from_space_3.SelectedItem
        if selected_item:
            self.selected_value_space_3 = selected_item.Content
            print(self.selected_value)
        else:
            print("No Item Selected")
    
    def combobox_item_spaces_4_SelectionChanged(self, sender, e):
        selected_item = self.combo_parameter_from_space_4.SelectedItem
        if selected_item:
            self.selected_value_space_4 = selected_item.Content
            print(self.selected_value)
        else:
            print("No Item Selected")


    def combobox_item_MEP_1_SelectionChanged(self, sender, e):
        selected_item = self.combo_param_to_1.SelectedItem
        if selected_item:
            self.selected_value_MEP_Instance_1 = selected_item.Content
            print(self.selected_value)
        else:
            print("No Item Selected")

    def combobox_item_MEP_2_SelectionChanged(self, sender, e):
        selected_item = self.combo_param_to_2.SelectedItem
        if selected_item:
            self.selected_value_MEP_Instance_2 = selected_item.Content
            print(self.selected_value)
        else:
            print("No Item Selected")

    def combobox_item_MEP_3_SelectionChanged(self, sender, e):
        selected_item = self.combo_param_to_3.SelectedItem
        if selected_item:
            self.selected_value_MEP_Instance_3 = selected_item.Content
            print(self.selected_value)
        else:
            print("No Item Selected")

    def combobox_item_MEP_4_SelectionChanged(self, sender, e):
        selected_item = self.combo_param_to_4.SelectedItem
        if selected_item:
            self.selected_value_MEP_Instance_4 = selected_item.Content
            print(self.selected_value)
        else:
            print("No Item Selected")
    
    

        



# MyWindow().ShowDialog()

if __name__ == "__main__":
    doc = revit.doc
    uidoc = HOST_APP.uidoc
    filtered_collector = DB.FilteredElementCollector(doc)

    # all_MEPFamilyTypes = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(db.BuiltInCategory.OST_DuctAccessory).ToElements()

    # all_MEP_family_types = filtered_collector.OfClass(db.FamilySymbol).ToElements()

    # quantity_MEP_Elements = len(all_MEP_family_types)

    # all_family_names = list(set([elem.Family.Name for elem in all_MEP_family_types]))

    # for family_name in all_family_names:
    #     combobox_item = ComboBoxItem()
    #     combobox_item.Content = family_name
    #     self.combo_family_type.Items.Add(combobox_item)

    # all_family_types = list(set([elem.GetTypeId() for elem in all_MEP_family_types]))

    my_window_obj = FamilySelectionWindow()
    my_window_obj.populate_combo_box_family_types(doc)
    my_window_obj.ShowDialog()

    # MyWindow().ShowDialog()

    # print("all MEP: ", all_family_types)


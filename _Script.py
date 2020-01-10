"""
Calculates the Net area of the Room, Minus the Casework and
places the value in the parameter 'Net Room Area' of a room
object.  It will skip rooms where Floor Area in Sq. Ft. per
Occupant is set to Gross
"""
__Title__ = 'Net Area Calculator'
__author__ = 'Nick Rosati'

# Net area Calculator
# Written by Nick Rosati
# 3.12.2019

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Analysis import *
from Autodesk.Revit.UI import *
import math
import Autodesk.Revit.DB as DB

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document


def find_phase(obj_phase):
  '''
  DOCSTRING:  function will determin phase of object
  INPUT: Object
  OUTPUT: Phase of Object
  '''
  pha = doc.Phases
  count = 0
  for p in pha:
    p_int = p.Id.IntegerValue
    if p_int == obj_phase:
      break
    else:
      count += 1
  return pha[count]
# Creating Collector instance and collecting all the casework from the model.

# Get Element Ids for all Families in project
family_name_id = DB.ElementId(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)

# Get the Name of the Family
family_name_provider = DB.ParameterValueProvider(family_name_id)
# Create a Filter to look for specific text
filter_string_begin_with = DB.FilterStringBeginsWith()
# Filter for specific string
string_begins_with_12_Base = DB.FilterStringRule(family_name_provider, filter_string_begin_with, '12 BASE', 1)
#Filter for specific string
string_begins_with_12_Tall = DB.FilterStringRule(family_name_provider, filter_string_begin_with, '12 Tall', 1)

string_begins_with_12_Case = DB.FilterStringRule(family_name_provider, filter_string_begin_with, '12  CASE', 1)



# Filter strings
p_fil_a = DB.ElementParameterFilter(string_begins_with_12_Base)
p_fil_b = DB.ElementParameterFilter(string_begins_with_12_Tall)
p_fil_c = DB.ElementParameterFiletr(string_begins_with_12_Case)

filter_list = [p_fil_a, p_fil_b, p_fil_c]

# Determine if one of the filters has items that pass
param_filter = DB.LogicalOrFilter(filter_list)
# Collect items with the parameters filetered above
cwk_coll = DB.FilteredElementCollector(doc)\
        .WherePasses(param_filter)\
        .WhereElementIsNotElementType()
# Collect items with the parameters filtered above all the way to Element IDs\
cwk_sel = DB.FilteredElementCollector(doc)\
        .WherePasses(param_filter)\
        .WhereElementIsNotElementType()\
        .ToElementIds()


# Collece all the Rooms in a project
all_rms = FilteredElementCollector(doc)\
          .OfCategory(BuiltInCategory.OST_Rooms)\
          .WhereElementIsNotElementType()

# Set all rooms initial Net Room Area Parameters equal to room area.
transaction = Transaction(doc, "Set NRA Start Value")
transaction.Start()

for rm in all_rms:
	rm_net = rm.LookupParameter('Net') #.AsInteger()
	rm_net.Set(0)
	rm_gross = rm.LookupParameter('Gross')
	rm_gross.Set(0)
	rm_area = rm.LookupParameter('Area').AsValueString()
	rm_area = rm_area[:-3]
	rm_area = float(rm_area)
	rm_area = int(math.ceil(rm_area))
	rm_net_start = rm.LookupParameter('Net Room Area')
	if rm_net_start:
		rm_net_start.Set(rm_area)

transaction.Commit()

transaction = Transaction(doc, "Update Net Area")
transaction.Start()
counter = 0
for cwk in cwk_coll:
   cwk_fi = cwk
   cwk = cwk.Id
# Get Casework Element
   cwk_element = doc.GetElement(cwk)
# Get depth of Casework element
   try: 
     de = cwk_element.Symbol.LookupParameter('Depth').AsDouble()
   except:
     de = cwk_element.LookupParameter('Depth').AsDouble()
# Get Width of Casework element
   try:
     wi = cwk_element.Symbol.LookupParameter('Width').AsDouble()
   except:
     wi = cwk_element.LookupParameter('Width').AsDouble()
# Get Area value for Casework element
   cwk_area = de * wi
   cwk_area = int(cwk_area)
   cwk_area = math.ceil(cwk_area)
# Get Casework Base Point
   cwk_bp = cwk_element.Location.Point
# Get Phase Caswork was created.
   cwk_ph = cwk_element.CreatedPhaseId.IntegerValue
# Get Casework Phase Created Parameter
   ph_set = find_phase(cwk_ph)
# Get Room at Casework Base Point
   cwk_rm = doc.GetRoomAtPoint(cwk_bp, ph_set)
# Get Room Number associated with Casework
   if cwk_rm:
      cwk_rm_param = cwk_rm.LookupParameter('Number').AsString()
      # cwk_rm_param = int(cwk_rm_param.AsString())

   # Check to see if caswork is in room, if it is update NRA Parameter.
   for rms in all_rms:
       rm_num = rms.LookupParameter('Number').AsString()
       # rm_num = int(rm_num.AsString())
       nra_param = rms.LookupParameter('Net Room Area').AsValueString()
       nra_param = nra_param[:-3]
       nra_param = float(nra_param)
       nra_param = int(math.ceil(nra_param))
       nra_param_wk = rms.LookupParameter('Net Room Area')
       if cwk_rm:
         if rm_num == cwk_rm_param:
           red_area = nra_param - cwk_area
           nra_param_wk.Set(red_area)
           counter += 1
transaction.Commit()
print('{} Casework objects Detected'.format(counter))

transaction = Transaction(doc, "Add Net Reduction to Room")
transaction.Start()

# After all casework has been checked find difference between
# Actual Area and Net Area and add Recuction value to 
# the Net SF Recuction Instance Parameter field
rm_no_occ_set = []
rm_num_nocc = []
rmcnt = 0
ocount = 0
for rm in all_rms:
    try:
      if rm:
      	rm_net_ckbox = rm.LookupParameter('Net')
      	rm_net_ckbox.Set(0)
      	rm_gross_ckbox = rm.LookupParameter('Gross')
      	rm_gross_ckbox.Set(0)
        rm_area = rm.LookupParameter('Area').AsValueString()
        rm_area = rm_area[:-3]
        rm_area = float(rm_area)
        rm_na = rm.LookupParameter('Name').AsString()
        rm_nm = rm.LookupParameter('Number').AsString()
        if rm_area == 0:
          rm_name = rm.LookupParameter('Name').AsString()
          if rm_name == 'Room':
            rmcnt += 1
          else:
            print('Room with name "{}" exists in the project but has Area = 0 sf'.format(rm_name))

    except:
      print('There are no Rooms in the model')
      break
    if rm_area > 0:
      rm_area = rm.LookupParameter('Area').AsValueString() # Get actual Area as a string
      rm_area = rm_area[:-3] # Remove the space and SF notation to get only numbers
      rm_area = float(rm_area) # convert string to a floating point number
      rm_net = rm.LookupParameter('Net Room Area').AsValueString() # Get Net Room Area Param as String
      rm_net = rm_net[:-3] # Remove the space and SF notation to get only numbers
      rm_net = float(rm_net) # convert string to a floating point number
      rm_net_none = rm.LookupParameter('Net Room Area') # Get Net Room Area Parameter variable to set
      reduction = rm_area - rm_net  # Find the difference in area
      reduction = math.ceil(reduction) # Round the number up to the next integer
      reduction = int(reduction) # Convert the float to an integer
      rm_red = rm.LookupParameter('Net SF Reduction') # set variable to point at intended destination
      g_or_n = rm.LookupParameter('OCCUPANCY SQ. FT. TYPE').AsString()
      # Set Occupant Load parameter initial value
      occ = rm.LookupParameter('OCCUPANCY SQ. FT. PER PERSON').AsDouble()
      occlf = rm.LookupParameter('Occupant Load Factor')
      occlf.Set(occ) # set 'Occupant Load' parameter used in tag
      
      if g_or_n == 'Net':
        rm_red.Set(reduction) # set 'Net SF Reduction' parameter
        rm_net_ckbox.Set(1) # Set Net Checkbox Parameter State to checked
        rm_net = math.floor(rm_net)
        rm_net = int(rm_net)        
        occ_ld = rm_net / occ
        ol = rm.LookupParameter('Occupant Load').AsValueString()
        ols = rm.LookupParameter('Occupant Load')
        occ_ld = math.floor(occ_ld)
        occ_ld = int(occ_ld)
        ols.Set(occ_ld)

      elif g_or_n == 'Gross':
      	rm_gross_ckbox.Set(1)
      	rm_red.Set(0)
        rm_net_none.Set(rm_area)
        occ_ld = rm_area / occ
        ol = rm.LookupParameter('Occupant Load').AsValueString()
        ols = rm.LookupParameter('Occupant Load')
        occ_ld = math.floor(occ_ld)
        occ_ld = int(occ_ld)
        if occ_ld < 1:
        	occ_ld = 1
        ols.Set(occ_ld)        

      else:
        rm_red.Set(0)
        rm_net_none.Set(rm_area)
      rm_occ_type_set = rm.LookupParameter('Room Occupancy').AsValueString()

      if rm_occ_type_set == '(none)':
      	ocount += 1
      	rm_no_occ_set.append(rm_na)
      	rm_num_nocc.append(rm_nm)
       
transaction.Commit()

print("\nNet SF Reduction Calculations are complete; please check the "\
      "'Analysys Results' section of the room properties\nand verify for proper results.")

if ocount > 0:
	print("\n{} Room(s) have no Occupany set please check the "\
	      "following room(s) in your project:".format(ocount))
	no_occ = zip(rm_num_nocc,rm_no_occ_set)
	for r in no_occ:
		print(r)
		
print("\nThe {} Casework objects deteted will stay selected after window is closed. Use 'Save Selection'"\
	   "\nto have this selection set avalible for future review".format(counter))


uidoc.Selection.SetElementIds(cwk_sel)

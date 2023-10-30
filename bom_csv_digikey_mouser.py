#
# Example python script to generate a BOM from a KiCad generic netlist
#
# Example: Sorted and Grouped CSV BOM
#

"""
    @package
    Output: CSV (comma-separated)
    Grouped By: Value, Footprint, DNP, specified extra fields
    Sorted By: Reference
    Fields: #, Reference, Qty, Value, Footprint, DNP, specified extra fields

    Outputs components grouped by Value, Footprint, and specified extra fields.
    Extra fields can be passed as command line arguments at the end, one field per argument.

    Command line:
    python "pathToFile/bom_csv_grouped_extra.py" "%I" "%O.csv" "Extra_Field1" "Extra_Field2"
"""

import sys
sys.path.append("C:\\Program Files\\KiCad\\7.0\\bin\\scripting\\plugins")
sys.path.append("C:/Git/mouser-api")
sys.path.append("C:/Git/mouser-api/mouser")
from mouser.api import MouserPartSearchRequest
# Import the KiCad python helper module and the csv formatter
import kicad_netlist_reader
import kicad_utils
import csv
import pandas as pd
import os
import re

# Get extra fields from the command line
extra_fields = sys.argv[3:]

comp_fields = ['Value', 'Footprint', 'Voltage', 'Type'] + extra_fields
#header_names = ['#', 'Reference', 'Qty'] + comp_fields + ["Mnf PN", "Mouser PN", "Digikey PN", "Mouser Price", "Digikey Price"]

def getComponentString(comp, field_name):
    if field_name == "Value":
        return comp.getValue()
    elif field_name == "Footprint":
        return comp.getFootprint()
    elif field_name == "DNP":
        return comp.getDNPString()
    elif field_name == "Datasheet":
        return comp.getDatasheet()
    else:
        return comp.getField( field_name )

def myEqu(self, other):
    """myEqu is a more advanced equivalence function for components which is
    used by component grouping. Normal operation is to group components based
    on their Value and Footprint.

    In this example of a more advanced equivalency operator we also compare the
    Footprint, Value, DNP and all extra fields passed from the command line. If
    these fields are not used in some parts they will simply be ignored (they
    will match as both will be empty strings).

    """
    result = True
    for field_name in comp_fields:
        if getComponentString(self, field_name) != getComponentString(other, field_name):
            result = False

    return result

def add_purchase_info(pn_df, row):
    # Filter pn data frame according to current row in BOM
    for field in comp_fields:
        if row[field] == "":
            continue

        pn_df = pn_df[pn_df[field] == row[field]]

    if len(pn_df) >= 1:
        row['Mnf PN']     = pn_df.iloc[0]['Mnf PN']
        row['Mouser PN']  = pn_df.iloc[0]['Mouser PN']
        row['Digikey PN'] = pn_df.iloc[0]['Digikey PN']

        if str(row['Mouser PN']) not in ["nan", "", "DNP"]:
            request = MouserPartSearchRequest('partnumber')
            request.part_search(row['Mouser PN'])
            part = request.get_clean_response()
            row['Mouser Manufacturer'] = part['Manufacturer']
            row['Mouser Description'] = part['Description']
            price_breaks = ""
            for price_break in part['PriceBreaks']:
                price_breaks += f"{price_break['Quantity']}x:{price_break['Price']}; "
            price_break = price_breaks[:-2]

            row['Mouser PriceBreaks']  = price_breaks
            row['Mouser Availability'] = part['Availability'].replace(" In Stock", "")

    return row


# Override the component equivalence operator - it is important to do this
# before loading the netlist, otherwise all components will have the original
# equivalency operator.
kicad_netlist_reader.comp.__eq__ = myEqu

# Generate an instance of a generic netlist, and load the netlist tree from
# the command line option. If the file doesn't exist, execution will stop
net = kicad_netlist_reader.netlist(sys.argv[1])

df = pd.DataFrame()

# Get all of the components in groups of matching parts + values
# (see kicad_netlist_reader.py)
grouped = net.groupComponents()
pn_df = pd.read_csv(os.path.dirname(sys.argv[1]) + "/PartNumbers.csv")

# Output all of the component information
for index, group in enumerate(grouped):
    refs = ""

    # Add the reference of every component in the group and keep a reference
    # to the component so that the other data can be filled in once per group
    for component in group:
        refs += component.getRef() + ", "
        c = component

    # Remove trailing comma
    refs = refs[:-2]

    # Fill in the component groups common data
    row = {}
    row['#'] = index + 1
    row['Reference'] = refs
    row['Qty'] = len(group)

    # Add the values of component-specific data
    for field_name in comp_fields:
        row[field_name] = getComponentString( c, field_name )

    row = add_purchase_info(pn_df, row)

    # concat row to end of final data frame
    df_row =  pd.DataFrame([row])
    df = pd.concat([df, df_row], ignore_index=True)


try:
    df.to_csv(sys.argv[2], index=False)
except:
    print(f"{sys.argv[2]} Open, close it and re-run BOM tool")

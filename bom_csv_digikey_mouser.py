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
#sys.path.append("C:/Git/mouser-api")
#sys.path.append("C:/Git/mouser-api/mouser")
from mouser.api import MouserPartSearchRequest

import digikey
from digikey.v3.productinformation import KeywordSearchRequest
from digikey.v3.batchproductdetails import BatchProductDetailsRequest

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

def get_mouser_part_info(part_number):
    # if part number starts with a quantity, then remove that for part number lookup
    if str(part_number).lower().startswith("qty:x"):
        part_number = part_number.split(" ")[1:]
        part_number = " ".join(part_number)
    part_data = {}
    part_data['Mouser PN'] = part_number
    if str(part_number) in ["", "nan"]:
        return part_data
    request = MouserPartSearchRequest('partnumber')
    request.part_search(part_number)
    part = request.get_clean_response()

    part_data['Mouser Manufacturer'] = part['Manufacturer']
    part_data['Mouser Description'] = part['Description']
    part_data['Mouser Availability'] = part['Availability'].replace(" In Stock", "")
    part_data['Mouser Category'] = part['Category']
    part_data['Mouser ManufacturerPartNumber'] = part['ManufacturerPartNumber']

    #store price breaks in form of qty{x}:price; (semi colin seperated list
    price_breaks = ""
    for price_break in part['PriceBreaks']:
        price_breaks += f"{price_break['Quantity']}x:{price_break['Price']}; "
    price_breaks = price_breaks[:-2]
    part_data['Mouser PriceBreaks'] = price_breaks

    return part_data

def get_digikey_part_info(part_number):
    # if part number starts with a quantity, then remove that for part number lookup
    if str(part_number).lower().startswith("qty:x"):
        part_number = part_number.split(" ")[1:]
        part_number = " ".join(part_number)
    
    part_data = {}
    part_data['Digikey PN']  = part_number
    if str(part_number) in ["", "nan"]:
        return part_data

    print(f"'{part_number}'")
    part = digikey.product_details(part_number)
    print(part)
    part_data['Digikey Manufacturer'] = part.manufacturer.value
    part_data['Digikey Description'] = part.detailed_description
    part_data['Digikey Availability'] = part.quantity_available
    part_data['Digikey Category'] = part.family.value
    part_data['Digikey ManufacturerPartNumber'] = part.manufacturer_part_number

    price_breaks = ""
    for price_break in part.my_pricing:
        price_breaks += f"{price_break.break_quantity}x:{price_break.unit_price}; "
    price_breaks = price_breaks[:-2]
    part_data['Digikey PriceBreaks'] = price_breaks 

    part_data['Digikey URL'] = part.product_url
    return part_data


def add_purchase_info(pn_df, row):
    # Filter pn data frame according to current row in BOM
    for field in comp_fields:
        if row[field] == "":
            continue

        pn_df = pn_df[pn_df[field] == row[field]]

    if len(pn_df) >= 1:

        if "DNP" not in str(row['Value']):
            row = row | \
                  get_mouser_part_info( pn_df.iloc[0]['Mouser PN']) | \
                  get_digikey_part_info(pn_df.iloc[0]['Digikey PN'])
            row['Mnf PN']     = pn_df.iloc[0]['Mnf PN']

            ##########################################################################
            # Collect price info for different quantities for both digikey and Mouser
            ##########################################################################
            for pcb_qty in [1, 10, 100, 1000]:
                min_price = 1000000000000
                for vendor in ["Mouser", "Digikey"]:
                    if str(row[f'{vendor} PN']).lower().startswith("qty:x"):
                        qty = row[f'{vendor} PN'].split(":")[0].split(" ")[0]
                        # number of parts on board times qty of pn per footprint times number of pcbs to purchase
                        qty = int(qty) * row['Qty'] * pcb_qty
                    else:
                        qty = row['Qty'] * pcb_qty
                    if f'{vendor} PriceBreaks' not in row or str(row[f'{vendor} PriceBreaks']) in ["nan", ""]:
                        continue

                    # Loop through price breaks from largest quantity to smallest
                    print(vendor)
                    print(row)
                    print(row[f'{vendor} PN'])
                    print(row[f'{vendor} PriceBreaks'])
                    for price_break in row[f'{vendor} PriceBreaks'].split(";")[::-1]:
                        print(f"'{price_break}'")
                        price_break_qty = price_break.split("x")[0]
                        print(price_break_qty)
                        price_break_qty = int(price_break_qty)

                        price_break_price = price_break.split(":")[1].replace("$", "")
                        price_break_price = float(price_break_price)

                        if qty >= price_break_qty:
                            break

                    row[f'{vendor} {pcb_qty} PCB Part Price total']  = round(price_break_price * qty / pcb_qty, 2)
                    min_price = min(round(price_break_price * qty / pcb_qty, 2), min_price)

                row[f'{vendor} 1 PCB Order Qty'] = int(qty / 1000)
                row[f'Cheapest Vendor {pcb_qty} PCB Part Price total']  = min_price

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


df_summary = pd.DataFrame()
row = {}
df_no_dnp = df[df['Mouser PN'] != "DNP"]

for pcb_qty in [1, 10, 100, 1000]:
    row['PCB Quantity'] = pcb_qty
    row['Mouser Component Cost']       = df_no_dnp[f'Mouser {pcb_qty} PCB Part Price total'].sum()
    row['Mouser Missing Part Numbers'] = df_no_dnp[f'Mouser {pcb_qty} PCB Part Price total'].isna().sum()
    row['Digikey Component Cost']       = df_no_dnp[f'Digikey {pcb_qty} PCB Part Price total'].sum()
    row['Digikey Missing Part Numbers'] = df_no_dnp[f'Digikey {pcb_qty} PCB Part Price total'].isna().sum()
    row['Cheapest Vendor Component Cost']       = df_no_dnp[f'Cheapest Vendor {pcb_qty} PCB Part Price total'].sum()
    row['Cheapest Vendor Missing Part Numbers'] = df_no_dnp[f'Cheapest Vendor {pcb_qty} PCB Part Price total'].isna().sum()

    df_row = pd.DataFrame([row])
    df_summary = pd.concat([df_summary, df_row], ignore_index=True)

try:
    df.to_csv(sys.argv[2], index=False)
except:
    print(f"{sys.argv[2]} Open, close it and re-run BOM tool")

try:
    df_summary.to_csv(os.path.dirname(sys.argv[2]) + "/BOM_Summary.csv", index=False)
except:
    print(f"{sys.dirname(os.path.argv[2]) + '/BOM_Summary.csv'} Open, close it and re-run BOM tool")

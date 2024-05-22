import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import re

from matplotlib.colors import Normalize
from matplotlib.colorbar import ColorbarBase
from wafermap.wafermap import WaferMap
from PIL import Image


def add_colorbar_to_image(image_path, vmin, vmax, cmap='coolwarm', bar_width=0.05, label=''):
    # Create the colorbar with adjusted sizes
    fig, ax = plt.subplots(figsize=(0.5, 9))  # Adjust the figsize to match the aspect ratio of the color bar
    norm = Normalize(vmin=vmin, vmax=vmax)
    cb = ColorbarBase(ax, cmap=cmap, norm=norm, orientation='vertical')
    cb.set_label(label, fontsize = 18)
    # Adjusting the font size of the tick labels
    
    for t in cb.ax.get_yticklabels():
        t.set_fontsize(14)  # Set to your desired font size


    # Save just the colorbar without margin
    plt.tight_layout()  # This will reduce excess white space
    plt.savefig('colorbar.png', dpi=300, bbox_inches='tight')
    plt.close()

    # Open the original image and the colorbar
    original_img = Image.open(image_path)
    colorbar_img = Image.open('colorbar.png')
    
    # Create a new image with a combined width that's slightly larger than the sum of the original image and color bar to reduce separation
    total_width = original_img.width + colorbar_img.width
    combined_img = Image.new('RGB', (total_width, original_img.height))

    # Paste the color bar closer to the original image
    gap = 750  # Number of pixels to reduce the gap by
    combined_img.paste(original_img, (0, 0))
    combined_img.paste(colorbar_img, (original_img.width - gap, 0))
    
    # Save the combined image
    combined_image_path = image_path.replace('.png', '_with_colorbar.png')
    combined_img.save(combined_image_path)
    
    return combined_image_path

# Initialize WaferMap
wafer_map = WaferMap(wafer_radius=100, cell_size=(25, 25), cell_margin=(0, 0), cell_origin=(0, 0), grid_offset=(0, 0), edge_exclusion=2.2, coverage='inner', notch_orientation=270)


def initialize_wafer_map():
    wafer_map = WaferMap(wafer_radius=100, cell_size=(25, 25), cell_margin=(0, 0), cell_origin=(0, 0), grid_offset=(0, 0), edge_exclusion=2.2, coverage='inner', notch_orientation=270)

    return wafer_map

# Retrieve the cell map
mapped_cells = wafer_map.cell_map

# Function to map value to color
def value_to_color(value, global_min, global_max):
    norm = plt.Normalize(global_min, global_max)
    cmap = cm.ScalarMappable(norm=norm, cmap='coolwarm')
    color = cmap.to_rgba(value)
    return cm.colors.rgb2hex(color)

# Define your allowed list of sheet names
allowed_sheets = [
    "L500W700",
    "L200W700",
    "L150W700",
    "L100W700",
    "L60W700"
    ]

# Ask the user for the file path and the physical property
file_path = input("Enter the path to your .xlsx file: ")
physical_property = input("What is the physical property of the input data set? (Vt, SS, or DIBL): ")

# Determine the unit based on the physical property
if physical_property.lower() in "vt":  # Check if the input is 'Vt' or 'DIBL'
    unit = "V"
elif physical_property.lower() == "ss":  # Check if the input is 'SS'
    unit = "V/dec"
elif physical_property.lower() == "dibl":  # Check if the input is 'dibl'
    unit = "V/V"
else:
    unit = None
    print("Invalid physical property. Please enter 'Vt', 'SS', or 'DIBL'.")

# If the unit is valid, proceed to load the Excel file
if unit:
    xls = pd.ExcelFile(file_path)
    print(f"Loaded data for {physical_property} with units {unit}.")
else:
    print("Data not loaded due to invalid physical property.")


xls = pd.ExcelFile(file_path)

# Initialize variables to find the global min and max
global_min = float('inf')
global_max = float('-inf')

# Scan through each sheet to find the global min and max
for sheet_name in xls.sheet_names:
    if sheet_name in allowed_sheets:  # Check if sheet name matches the pattern
        data = pd.read_excel(xls, sheet_name=sheet_name)
        current_min = data['50%'].min()
        current_max = data['50%'].max()
        global_min = min(global_min, current_min)
        global_max = max(global_max, current_max)

print("Global min:", global_min)
print("Global max:", global_max)

for sheet_name in allowed_sheets:
    data = pd.read_excel(xls, sheet_name=sheet_name)
    wafer_map = initialize_wafer_map()  
    for index, row in data.iterrows():
        if pd.isna(row['DieX']) or pd.isna(row['DieY']):
            continue
        die_x = int(row['DieX']) + 1
        die_y = int(row['DieY'])
        die_num = int(row['Die'])
        median_value = row['50%']
        # print(f"DieX: {die_x}, DieY: {die_y}, Median Value: {median_value}")
            
        try:
            color = value_to_color(median_value, global_min, global_max)
        except ValueError as e:
            print(f"An error occurred with median value {median_value}: {e}")
        # Optionally, skip this iteration if there's an error
            continue

        formatted_value = "{:.2f}".format(median_value)
        color = value_to_color(median_value, global_min, global_max)
        cell_style = {'fillColor': color, 'fillOpacity': 1.0, 'fill': True}

        # Use the mapped cell coordinates to apply styles
        # Ensure that this section is inside the for loop
        if (die_x, die_y) in mapped_cells:
            wafer_map.style_cell(cell=(die_x, die_y), cell_style=cell_style)
            wafer_map.add_label(cell=(die_x, die_y), label_text=(str(die_num)), offset=(1.5, 1.5))
            wafer_map.add_label(cell=(die_x, die_y), label_text=(str(formatted_value)), label_html_style=('font-size: 20pt; color: black; text-align: center'), offset=(12.5, 12.5))

    # The save and colorbar addition should remain outside the inner loop
    output_png = f'wafer_{physical_property}_visualization_{sheet_name}.png'
    wafer_map.save_png(output_png)
    label = f'Median {physical_property} Distribution {sheet_name} (Unit: {unit})'
    add_colorbar_to_image(output_png, global_min, global_max, cmap='coolwarm', bar_width=0.05, label=label)


"""
# The path to your wafer visualization image
image_path = 'wafer_visualization_L200W70.png'

# Add a color bar to the image
output_with_colorbar_path = add_colorbar_to_image(image_path, global_min, global_max, cmap='coolwarm', Vt, bar_width=0.05)

# Now `output_with_colorbar_path` holds the path to the new image with the color bar
"""

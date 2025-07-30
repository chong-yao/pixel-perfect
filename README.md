# Pixel-Perfect: Image-to-Excel and Excel-to-Image Converter

map each pixel into a cell

# A program to split an image down to its individual pixels no matter the resolution.

#### Takes a standard image file and transforms it into an Excel file where each cell is coloured to represent a pixel.

#### It can also perform the reverse operationm taking a specially formatted Excel file and converting back into an image.

### Useful for:
- ***Upscaling & enhancing low-resolution images**
- **Detailed Image Editing**: Manually edit images pixel-by-pixel directly within Excel by changing cell colors.
- **Creative Projects**: Generate unique visual effects or data art
- **Creating Pixel Art**: Convert any image into a pixelated grid, perfect for artists and designers.

## Features

-   **Image to Excel:** Converts standard image formats (PNG, JPG, etc.) into an `.xlsx` file.
-   **Excel to Image:** Rebuilds an image from a color-formatted Excel sheet.
-   **Resolution Independent:** The script automatically detects the image resolution for a 1:1 pixel-to-cell conversion. You can also define a custom grid size to downscale images.
-   **Optimized for Performance:** Uses efficient libraries to handle large images and minimize processing time.



---

# 

![Excel to Image Demo](https://i.imgur.com/a/5O1n94a)

A Python script that splits any image into its individual pixels, transforming it into an Excel spreadsheet where each cell's background color represents a single pixel. It can also reverse the process, converting a colored Excel sheet back into an image file.

This tool gives you the power to see, edit, and manipulate images at the most fundamental level—the pixel.

### Useful for:
-   **Creating Pixel Art:** Convert any image into a pixelated grid, perfect for artists and designers.
-   **Detailed Image Editing:** Manually edit images pixel-by-pixel directly within Excel by changing cell colors.
-   **Image Analysis:** Examine the color composition of an image in a structured, grid-based format.
-   **Creative Projects:** Generate unique visual effects or data art from your favorite pictures.



## How to Use

### 1. Prerequisites

Before you begin, ensure you have Python 3 installed on your system. You will also need to install the required libraries, `Pillow` and `openpyxl`.

You can install them using pip:
```bash
pip install Pillow openpyxl
```

### 2. Configuration

Open the Python script (`your_script_name.py`) in a text editor and modify the variables in the `--- USER SETTINGS ---` section to match your needs.

```python
# ==============================================================================
# --- USER SETTINGS ---
# ==============================================================================

# -- Choose the operation:
# 'to_excel': Converts the input image to an Excel file.
# 'to_image': Converts the input Excel file back to an image.
# 'both':     Runs both conversions sequentially.
OPERATION_MODE = 'to_excel'

# -- Define file paths:
# The source image to convert.
INPUT_IMAGE_PATH = Path("your_input_image.png")
# The path for the generated Excel file.
EXCEL_PATH = Path("pixel-art.xlsx")
# The path for the final image restored from Excel.
FINAL_IMAGE_PATH = Path("output_image.png")

# -- Set processing parameters:
# The resolution for the pixel art. The script automatically uses the image's
# original resolution, but you can override it here (e.g., 128 for a 128x128 grid).
# For a 1:1 conversion, set these to your image's width and height.
GRID_SIZE_X = 1920
GRID_SIZE_Y = 1080

# The multiplier for scaling up the final image. A higher value results in a larger image.
UPSCALE_FACTOR = 5
```

#### Configuration Details:
*   `OPERATION_MODE`:
    *   `'to_excel'`: Use this to create the Excel file from your image.
    *   `'to_image'`: Use this after you have a generated Excel file that you want to turn back into an image.
    *   `'both'`: A convenient option to run the entire cycle in one go.
*   `INPUT_IMAGE_PATH`: The path to the source image you want to process.
*   `EXCEL_PATH`: The name of the Excel file that will be created or read.
*   `GRID_SIZE_X` & `GRID_SIZE_Y`: **This is a key setting.**
    *   For a **perfect 1:1 pixel representation**, set these values to the exact width and height of your input image.
    *   To create **pixel art from a high-resolution image**, use smaller values (e.g., `128` and `128`) to downscale it.
    *   **Warning:** Very high values (e.g., `4096`x`4096`) will create extremely large Excel files and may be very slow or fail to run on most computers.
*   `UPSCALE_FACTOR`: When converting from Excel back to an image, this determines the final image size. An `UPSCALE_FACTOR` of `5` on a `128x128` grid will produce a `640x640` pixel image.

### 3. Execution

Once you have configured the settings, save the script and run it from your terminal:

```bash
python your_script_name.py
```

The script will print its progress and notify you upon completion. The output files will be saved in the locations you specified in the configuration.

## How It Works

### Image to Excel

1.  **Load Image**: The script opens the input image using the Pillow library.
2.  **Resize**: It resizes the image to the specified `GRID_SIZE_X` and `GRID_SIZE_Y`. For creating pixel art, this step condenses the image detail. For a 1:1 copy, this step ensures the dimensions are correct.
3.  **Create Excel Sheet**: A new Excel workbook is created using `openpyxl`.
4.  **Pixel Mapping**: The script iterates through every pixel of the resized image. For each pixel, it gets its RGB color value.
5.  **Cell Coloring**: It sets the background fill of the corresponding cell in the Excel worksheet to the pixel's color. To improve performance, color styles are cached so they are not recreated unnecessarily.
6.  **Formatting**: The cell widths and heights are adjusted to be square, providing a visually accurate pixel grid.

### Excel to Image

1.  **Load Workbook**: The script loads the Excel file in a highly-optimized, read-only mode.
2.  **Read Dimensions**: It automatically detects the number of rows and columns that contain colored cells.
3.  **Create New Image**: A new, blank image is created with dimensions matching the Excel grid.
4.  **Color Reading**: The script iterates through each cell in the Excel sheet and reads its background color.
5.  **Pixel Creation**: The color from each cell is used to draw a pixel at the corresponding coordinate in the new image.
6.  **Upscaling**: The final image is resized (scaled up) without blurring to ensure the pixelated aesthetic is preserved and the image is easily viewable.

## Notes and Limitations

-   **Performance**: Converting large images (e.g., resolutions higher than 512x512) can be time-consuming and generate very large Excel files. It is recommended to start with smaller grid sizes.
-   **Dependencies**: This script requires an active Python environment and is not a standalone executable.
-   **Excel Limitations**: Microsoft Excel has its own limitations on the maximum number of rows and columns. This script is best suited for resolutions within those bounds.
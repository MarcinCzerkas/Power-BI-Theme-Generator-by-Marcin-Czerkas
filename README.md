# Power BI Theme Generator

A simple utility to generate custom Power BI themes from Excel inputs and JSON templates using Python. 🐍

[![Watch the video](https://img.youtube.com/vi/lRWNpOoFIYA/0.jpg)](https://youtu.be/lRWNpOoFIYA)
*Watch the video to find out more.*

## ✅ Features

- Define theme parameters in **Excel**:
  - Theme type (Light / Dark)
  - Font
  - Visual padding
  - Color palette (8 colors)
- Automatically computes **lighter and darker variants** of the primary color.
- Generates a ready-to-use **Power BI theme JSON** file.
- Minimal dependencies, easy to run locally.
- Optional popup confirmation when the theme is generated.

## ⬇️ Installation

1. Clone the repository or download the files manually

2. Install Python dependencies:

```
pip install openpyxl
```

## 👨‍💻 Usage

1. Fill in your desired theme parameters in *Theme_Generator.xlsx*

    - The Excel file contains data validation rules helping you choose values from dropdown lists
    - You can choose colors by selecting the cell background (fill)

2. Run the Python script (double-click only): *PBI_Theme_Generator.pyw*

3. A popup will confirm successful generation, and the theme JSON will be saved in the same folder

4. After loading to Power BI, you can choose any of the background images stored in */Background_images/PNG* or create a new one using the PPT file *Background_generator.pptx*

## 🎨 Customization

Default grey color is used if any cell color is empty.

You can adjust default settings directly in the Python script if needed or head over to the generated JSON file.

Feel free to use the PPT template to edit your own background images.

## 💡 Future expansion

This project is planned to be expanded in the future.

Possible extensions:

- adding advanced settings to the UI (e.g. line chart smoothness)
- automation of background image generation
- online hosting instead of Excel UI

## 📃 License

This project is licensed under the MIT License.

## 🖥️ Preview

![Dashboard-light](/Assets/Sample_light.png)
Sample light theme

![Dashboard-dark](/Assets/Sample_dark.png)
Sample dark theme

![Excel generator](/Assets/Excel_generator.png)
UI in Excel

![Popup](/Assets/Popup.png)
Popup after successful JSON generation
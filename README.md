# ExPlot
A modern Python/Tkinter app to plot data from Excel sheets using *seaborn* and *matplotlib*, featuring beautiful theming with ttkbootstrap. Fully developed with Claude 3.7 Sonnet, Cascade Base, SWE-1, and GPT-4.1 using Windsurf.

![ExPlot with ttkbootstrap Theme](img/ExPlot_ttkbootstrap.png)

*Screenshot: ExPlot with ttkbootstrap's 'darkly' theme*

# Data structure
Provide an Excel file with one or several sheets for different data sets to be plotted. The data should be in the following format:

| x_category | y_value |  group  |
|------------|---------|---------|
| A          | 1       | Treated |
| A          | 2       | Treated |
| A          | 3       | Treated |
| A          | 8       | Control |
| A          | 6       | Control |
| A          | 7       | Control |
| B          | 14      | Treated |
| B          | 15      | Treated |
| B          | 13      | Treated |
| B          | 21      | Control |
| B          | 22      | Control |
| B          | 21      | Control |

Rows with identical x_categories will be averaged and and used to derive error estimates, or they might be plotted as individual data points. The *group* column can be used to plot data as grouped elements.

The *x_category* column contains the categories to be plotted on the x-axis, the *y_value* column contains the values to be plotted on the y-axis, and the *group* column contains the categories to be used for grouping data.
Alternatively, provide several *y_value* columns, each with a different name.

# Features
- Bar graphs
- Box plots
- Strip plots
- XY plots
- XY Fitting with predefined or cutom models
- t-tests, ANOVA, and post-hoc tests
- Save and load data
- App themes (`View > Themes`)

# Examples
File --> Load Example Data

# Installation

## Using conda
```bash
# Create and activate a conda environment
conda create --name explot python=3.10
conda activate explot
conda config --add channels conda-forge
conda config --set channel_priority strict
conda install --file requirements.txt
pip install PyMuPDF ttkbootstrap
```

## Using pip
```bash
# Create and activate a virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install from requirements.txt
pip install -r requirements.txt
pip install PyMuPDF ttkbootstrap
```

# Running the Application

## With ttkbootstrap Theme (recommended)
```bash
python launch.py
```

## With Default Theme
```bash
python explot.py
```

# Packaging
```bash
./build.sh
```




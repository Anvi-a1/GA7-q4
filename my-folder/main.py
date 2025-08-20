# Import necessary libraries
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# --- Step 1: Load the Dataset ---
# Load the supply chain data from the Excel file into a pandas DataFrame.
# Note: Reading Excel files requires the 'openpyxl' library.
# You can install it by running: pip install openpyxl
#
# Make sure the 'q-excel-correlation-heatmap.xlsx' file is in the same directory as this script.
try:
    df = pd.read_excel('q-excel-correlation-heatmap.xlsx')
    print("Successfully loaded the dataset.")
except FileNotFoundError:
    print("Error: 'q-excel-correlation-heatmap.xlsx' not found.")
    print("Please make sure the data file is in the same folder as the script.")
    exit()
except Exception as e:
    print(f"An error occurred: {e}")
    print("Please ensure you have the 'openpyxl' library installed (`pip install openpyxl`).")
    exit()

# --- Step 2: Calculate the Correlation Matrix ---
# Use the .corr() method on the DataFrame to compute the pairwise correlation of columns.
correlation_matrix = df.corr()
print("\nCorrelation Matrix:")
print(correlation_matrix)

# --- Step 3: Export the Correlation Matrix to CSV ---
# Save the calculated correlation matrix to a file named 'correlation.csv'.
# The 'index=True' argument ensures that the row labels (the variable names) are included.
correlation_matrix.to_csv('correlation.csv', index=True)
print("\nSuccessfully saved correlation matrix to 'correlation.csv'.")

# --- Step 4: Generate and Save the Heatmap ---
# Create a heatmap to visualize the correlation matrix.
# The image size in pixels is determined by figsize * dpi.
# To meet the < 512x512 pixel requirement, we'll create a 5x5 inch figure at 100 DPI,
# resulting in a 500x500 pixel image.
plt.figure(figsize=(5, 5)) # Set the figure size in inches

# Use seaborn's heatmap function.
# - 'correlation_matrix': The data to plot.
# - 'annot=True': Display the correlation values on the heatmap.
# - 'cmap='coolwarm'': Use a color map where low values are blue (cool) and high values are red (warm).
#   This is similar to the Red-White-Green scale, showing negative and positive correlations clearly.
# - 'fmt=".2f"': Format the annotations to two decimal places.
sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', fmt=".2f", linewidths=.5)

# Add a title to the plot.
plt.title('Supply Chain Metrics Correlation Matrix')

# Ensure the plot layout is tight and clean.
plt.tight_layout()

# Save the generated heatmap as a PNG image file.
# We set the DPI (dots per inch) to 100.
plt.savefig('heatmap.png', dpi=100)
print("Successfully generated and saved heatmap to 'heatmap.png' (500x500 pixels).")

# --- (Optional) Step 5: Display the Plot ---
# If you are running this script in an environment that supports GUI,
# this will show the plot in a new window.
# plt.show()

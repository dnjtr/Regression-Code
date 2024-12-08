import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import statsmodels.api as sm
from PIL import ImageGrab

class RegressionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Regression Analysis Tool")
        self.root.geometry("1024x768")
        
        # Buttons and input fields
        self.load_button = tk.Button(root, text="Load File (Excel/CSV)", command=self.load_file)
        self.load_button.grid(row=0, column=0, padx=10, pady=10)

        # Dataset display
        self.tree_frame = tk.Frame(root)
        self.tree_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10)
        
        self.tree = ttk.Treeview(self.tree_frame, show='headings')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.config(yscrollcommand=self.scrollbar.set)

        # Input fields for regression
        tk.Label(root, text="Input Y Range (column number):").grid(row=2, column=0, sticky=tk.W)
        self.y_entry = tk.Entry(root)
        self.y_entry.grid(row=2, column=1, pady=5)
        
        tk.Label(root, text="Input X Range (comma-separated column numbers):").grid(row=3, column=0, sticky=tk.W)
        self.x_entry = tk.Entry(root)
        self.x_entry.grid(row=3, column=1, pady=5)
        
        tk.Label(root, text="Confidence Level (default 95%):").grid(row=4, column=0, sticky=tk.W)
        self.confidence_entry = tk.Entry(root)
        self.confidence_entry.insert(0, "95")
        self.confidence_entry.grid(row=4, column=1, pady=5)
        
        self.run_button = tk.Button(root, text="Perform Regression", command=self.perform_regression)
        self.run_button.grid(row=5, column=0, columnspan=2, pady=10)

        # Results Text Box with Scrollbar
        self.results_frame = tk.Frame(root)
        self.results_frame.grid(row=6, column=0, columnspan=4, padx=10, pady=10)
        
        self.results_text = tk.Text(self.results_frame, wrap=tk.WORD, width=100, height=20)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar for Results Text Box
        self.results_scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_text.config(yscrollcommand=self.results_scrollbar.set)
        
        self.results_text.config(state=tk.DISABLED)  # Make text box read-only

        # Add buttons for image and dataset screenshot
        self.image_button = tk.Button(root, text="Add Image", command=self.add_image)
        self.image_button.grid(row=7, column=0, pady=10)
        
        self.screenshot_button = tk.Button(root, text="Capture Dataset Screenshot", command=self.capture_screenshot)
        self.screenshot_button.grid(row=7, column=1, pady=10)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        try:
            # Load dataset
            if file_path.endswith(".csv"):
                self.data = pd.read_csv(file_path)
            elif file_path.endswith(".xlsx"):
                self.data = pd.read_excel(file_path)
            
            self.populate_treeview(self.data)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file: {e}")

    def populate_treeview(self, data):
        # Clear existing data in the Treeview
        self.tree.delete(*self.tree.get_children())
        
        # Set column headers
        self.tree["column"] = list(data.columns)
        self.tree["show"] = "headings"
        
        for col in data.columns:
            self.tree.heading(col, text=col)
        
        # Insert rows into Treeview
        for _, row in data.iterrows():
            self.tree.insert("", tk.END, values=list(row))
        
        # Resize columns to fit content
        for col in data.columns:
            self.tree.column(col, width=100, anchor="center")

    def perform_regression(self):
        try:
            y_col = int(self.y_entry.get().strip()) - 1
            x_cols = [int(col.strip()) - 1 for col in self.x_entry.get().split(",")]
            
            y = pd.to_numeric(self.data.iloc[:, y_col], errors='coerce')
            X = self.data.iloc[:, x_cols].apply(pd.to_numeric, errors='coerce')

            # Drop rows with missing values
            data = pd.concat([y, X], axis=1).dropna()
            y = data.iloc[:, 0]
            X = data.iloc[:, 1:]

            X = sm.add_constant(X)  # Add intercept
            model = sm.OLS(y, X).fit()
            
            self.display_results(model)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during regression: {e}")

    def display_results(self, model):
        # Enable the text box to add content
        self.results_text.config(state=tk.NORMAL)
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        
        # Display Regression Equation
        eq = "Regression Equation:\n\nY = "
        params = model.params
        eq += f"{params['const']:.4f} "
        for var in params.index[1:]:
            eq += f"+ {params[var]:.4f} * {var} "
        
        # Add Regression Equation to the text box
        self.results_text.insert(tk.END, eq + "\n\n")
        
        # Display Regression Statistics
        stats = f"Regression Statistics:\n\nMultiple R: {model.rsquared ** 0.5:.4f}\n"
        stats += f"R Square: {model.rsquared:.4f}\n"
        stats += f"Adjusted R Square: {model.rsquared_adj:.4f}\n"
        
        # Use the model scale for the standard error
        stats += f"Standard Error: {model.scale**0.5:.4f}\n"  # Standard error from model scale
        stats += f"Observations: {len(model.model.endog)}"
        
        # Add Statistics to the text box
        self.results_text.insert(tk.END, stats + "\n\n")
        
        # Display Regression Coefficients and other results
        summary_df = model.summary2().tables[1]
        self.results_text.insert(tk.END, "Variable   Coefficient   Std Error   t Stat   P-value   Lower 95%   Upper 95%\n")
        for index, row in summary_df.iterrows():
            self.results_text.insert(tk.END, f"{index:<12} {row['Coef.']:<12.4f} {row['Std.Err.']:<12.4f} {row['t']:<8.4f} "
                                            f"{row['P>|t|']:<8.4f} {row['[0.025']:<12.4f} {row['0.975]']:<12.4f}\n")
        
        # Display Recommendations
        recommendation = "\nRecommendation: Variables to Purge\n\n"
        p_values = model.pvalues
        for var, p_value in p_values.items():
            if var != "const" and p_value > 0.05:  # If p-value > 0.05, recommend purging
                recommendation += f"Variable {var} has a high p-value ({p_value:.4f}). Consider purging it.\n"
        
        # Add Recommendations to the text box
        if recommendation != "\nRecommendation: Variables to Purge\n\n":
            self.results_text.insert(tk.END, recommendation)
        else:
            self.results_text.insert(tk.END, "No variables have high p-values for purging.\n")
        
        # Disable the text box to make it read-only
        self.results_text.config(state=tk.DISABLED)

    def add_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png")])
        if file_path:
            messagebox.showinfo("Image Added", "Image successfully added!")

    def capture_screenshot(self):
        try:
            x = self.root.winfo_rootx()
            y = self.root.winfo_rooty()
            x1 = x + self.root.winfo_width()
            y1 = y + self.root.winfo_height()
            
            screenshot = ImageGrab.grab(bbox=(x, y, x1, y1))
            save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
            if save_path:
                screenshot.save(save_path)
                messagebox.showinfo("Screenshot Captured", f"Screenshot saved to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not capture screenshot: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RegressionApp(root)
    root.mainloop()

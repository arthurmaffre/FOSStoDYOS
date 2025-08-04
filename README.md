# FOSStoDYOS

**Streamline Your Analysis Pipeline**

FOSStoDYOS is a simple, elegant tool designed to automate the export of Moûts data from FOSS to Dyostem. Built for Clochers et Terroirs, it simplifies the process of extracting and formatting analysis data from Excel files, saving time and reducing manual errors. With a clean, intuitive interface inspired by modern design principles, it lets you focus on what matters: insightful analysis.

## Key Features
- **Effortless Data Processing**: Upload your FOSS Excel file, select a date, and generate a Dyostem-ready export in seconds.
- **Date Detection**: Automatically scans and lists available dates from your uploaded file.
- **Custom Formatting**: Maps columns precisely, with hardcoded values like "Sauvignon blanc" for consistency.
- **Modern UI**: A sleek, Apple-inspired interface that's easy on the eyes and simple to use.

## Requirements
To get started, ensure you have Python installed (version 3.8 or later recommended). The project depends on a few lightweight libraries:

- Streamlit
- Openpyxl
- Pandas

Install them easily with:
```
pip install -r requirements.txt
```

## Installation
1. Clone the repository:
   ```
   git clone https://github.com/your-username/FOSStoDYOS.git
   ```
   (Replace with your actual repo URL if hosted.)

2. Navigate to the project directory:
   ```
   cd FOSStoDYOS
   ```

3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage
Running the app is as straightforward as it gets:

1. Launch the Streamlit application:
   ```
   streamlit run main.py
   ```

2. Open your web browser to the provided local URL (usually http://localhost:8501).

3. **Step-by-Step in the App**:
   - Upload your source `.xlsx` file from FOSS.
   - Select a date from the automatically detected list.
   - Click "Générer le fichier" to process.
   - Download the generated `export_dyostem.xlsx` file, ready for Dyostem.

That's it—clean, quick, and reliable.

## About the Project
This tool evolved from a command-line script (`old_main.py`) into a user-friendly web app, making data export accessible to everyone on the team. It's perfect for handling Moûts analyses, ensuring data like sugar quantity, acidity, pH, and more is perfectly formatted.

If you encounter any issues or have suggestions, feel free to open an issue. Let's make analysis pipelines even smoother together.

*Built with simplicity in mind.*
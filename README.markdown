# Fishbone Diagram Generator

A Streamlit web application to create customizable fishbone (Ishikawa) diagrams from an Excel file. Users can upload an Excel file, adjust diagram parameters, preview the diagram, and download it as a PNG or PowerPoint (PPTX) file. This app is ideal for visualizing root cause analyses in quality control, project management, or process improvement.

## Features
- **Excel Input**: Upload an Excel file with a "Main Problem" column and category columns containing causes.
- **Customizable Parameters**: Adjust the fish head radius, main problem font size, and bone length via input fields.
- **Interactive Preview**: View the generated fishbone diagram in the web interface.
- **Download Options**: Save the diagram as a PNG image or a PowerPoint slide.
- **Error Handling**: User-friendly error messages for invalid Excel files or data formats.

## Installation

### Prerequisites
- Python 3.8 or higher
- Git (to clone the repository)
- A web browser (e.g., Chrome, Firefox)

### Steps
1. **Clone the Repository**
   ```bash
   git clone https://github.com/your-username/your-repo-name.git
   cd your-repo-name
   ```

2. **Create a Virtual Environment** (recommended)
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies**
   Ensure you have a `requirements.txt` file with the following:
   ```plaintext
   streamlit
   pandas
   matplotlib
   python-pptx
   openpyxl
   ```
   Install them using:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the App Locally**
   ```bash
   streamlit run app.py
   ```
   Open the provided URL (usually `http://localhost:8501`) in your browser.

## Usage
1. **Launch the App**: After running `streamlit run app.py`, the app opens in your browser.
2. **Upload an Excel File**: Click the file uploader to select a `fishbone.xlsx` file (see [Excel File Format](#excel-file-format) below).
3. **Adjust Parameters**:
   - **Fish Head Radius**: Set the size of the fish head (default: 1.5).
   - **Main Problem Font Size**: Set the font size for the main problem text (default: 9.0).
   - **Bone Length**: Set the length of category bones (default: 180.0).
4. **Generate Preview**: Click "Generate Preview" to display the fishbone diagram.
5. **Download**: Use the "Download as PNG" or "Download as PPTX" buttons to save the diagram.

## Excel File Format
The app expects an Excel file (`.xlsx`) with the following structure:

| Main Problem          | Category 1 | Category 2 | Category 3 |
|-----------------------|------------|------------|------------|
| Production Delays     | Machine Failure | Staff Shortage | Supply Issues |
|                       | Maintenance Issues | Training Gaps | Late Deliveries |

- **Main Problem**: A single, non-empty string in the first row of the "Main Problem" column.
- **Categories**: Other columns represent categories, with causes listed as strings in subsequent rows.
- **Notes**:
  - Empty cells are ignored.
  - At least one category must have valid causes.
  - Non-string values (e.g., numbers) are converted to strings.

## Deployment on Streamlit Cloud
1. **Push to GitHub**:
   - Ensure `app.py` and `requirements.txt` are in your repository’s root.
   - Push to a public GitHub repository.
2. **Create a Streamlit Cloud App**:
   - Log in to [Streamlit Cloud](https://streamlit.io/cloud).
   - Click "New app" and connect your GitHub repository.
   - Specify `app.py` as the main file and select the appropriate branch.
3. **Deploy**: Streamlit Cloud will install dependencies and deploy the app. Check the logs if issues arise.

## Troubleshooting
- **Error: "Column 'Main Problem' is not found"**: Ensure your Excel file has a column named exactly "Main Problem".
- **Error: "No valid categories found"**: Verify that at least one category column contains non-empty causes.
- **Diagram Clipping**: Adjust the `Fish Head Radius` or `Bone Length` to fit the diagram within the preview.
- **Deployment Issues**: Check Streamlit Cloud logs for dependency or memory errors. Ensure `requirements.txt` is correct.

## Example
1. Prepare an Excel file (`fishbone.xlsx`) with the format shown above.
2. Upload it via the app’s file uploader.
3. Set parameters (e.g., Fish Head Radius: 2.0, Font Size: 10.0, Bone Length: 200.0).
4. Click "Generate Preview" to view the diagram.
5. Download as PNG or PPTX for presentations or reports.

## Contributing
Contributions are welcome! Please:
1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/your-feature`).
3. Commit changes (`git commit -m "Add your feature"`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Open a pull request.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments
- Built with [Streamlit](https://streamlit.io/), [Matplotlib](https://matplotlib.org/), and [python-pptx](https://python-pptx.readthedocs.io/).
- Inspired by the need for user-friendly root cause analysis tools.
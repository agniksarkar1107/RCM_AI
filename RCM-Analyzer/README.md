# Risk Control Matrix Analyzer

An AI-powered tool that analyzes Risk Control Matrix (RCM) documents to identify control objectives, risks, and provide detailed risk analysis across departments.

## Features

- **Document Analysis**: Upload and analyze Risk Control Matrices in Excel, CSV, PDF, or DOCX formats
- **AI-Powered Risk Analysis**: Utilizes Google's Gemini AI model to assess risks and controls
- **Interactive Dashboard**: Visual representation of departmental risks and control gaps
- **Downloadable Reports**: Export analysis results as Excel or CSV for offline use
- **Risk Visualization**: Color-coded risk classification (High, Medium, Low) across departments
- **Control Gap Identification**: Highlights control gaps and provides recommendations

## Setup

### Prerequisites
- Python 3.10+
- Required Python packages (see requirements.txt)

### Installation

1. Clone the repository
```bash
git clone https://github.com/Arittra-Bag/RCM-Analyzer.git
cd RCM-Analyzer
```

2. Install dependencies
```bash
pip install -r requirements.txt
```

3. Set up environment variables
Create a `.env` file in the root directory with your Gemini API key:
```
GEMINI_API_KEY=your_api_key_here
```

### Running the Application

```bash
streamlit run app.py
```

## Usage

1. **Upload a Risk Control Matrix document** (Excel, CSV, PDF, or DOCX)
2. Click **Analyze Document** to process the file
3. View the analysis results in the interactive dashboard
4. Download the complete analysis as Excel or CSV using the download buttons
5. Explore risks by department using the tabbed interface
6. Use the **Clear Analysis** button to reset and start with a new document

## Streamlit Cloud Deployment

This application is designed to be compatible with Streamlit Cloud deployment. To deploy on Streamlit Cloud:

1. Fork this repository to your GitHub account
2. On Streamlit Cloud, create a new app and connect it to your fork
3. Set the required environment variables (GEMINI_API_KEY)
4. Deploy the application

### SQLite Version Compatibility

ChromaDB requires SQLite version 3.35.0 or higher. The application will automatically handle environments with older SQLite versions:

- If the SQLite version is compatible, the app will use persistent ChromaDB storage
- If the SQLite version is too old (like on some Streamlit Cloud instances), the app will use in-memory ChromaDB storage
- This ensures the application works correctly regardless of the SQLite version available

## Example Data

An example RCM file is included in the `examples` directory for testing purposes.

## Project Structure

```
├── app.py                    # Main Streamlit application
├── requirements.txt          # Python dependencies
├── .env                      # Environment variables (create this file)
├── README.md                 # Project documentation
├── utils/                    # Utility modules
│   ├── __init__.py           # Package initialization
│   ├── document_processor.py # Document processing utilities
│   ├── db.py                 # ChromaDB vector database integration
│   └── gemini.py             # Gemini API integration for AI analysis
└── chroma_db/                # ChromaDB persistent storage (created at runtime)
```

## Development

To contribute or modify the project:

1. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install development dependencies:
```bash
pip install -r requirements.txt
```

3. Make your changes and test them with Streamlit:
```bash
streamlit run app.py
```

## License

MIT License

## Author

Created by Arittra Bag and Agnik Sarkar

## Acknowledgements

- [Streamlit](https://streamlit.io/) for the web framework
- [Google Gemini AI](https://ai.google.dev/) for the AI analysis
- [ChromaDB](https://www.trychroma.com/) for vector database functionality
- [Plotly](https://plotly.com/) for interactive visualizations 
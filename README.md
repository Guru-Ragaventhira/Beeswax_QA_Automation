# Beeswax QA Automation Suite

![Python](https://img.shields.io/badge/Python-3.6%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.24.0%2B-red)
![pandas](https://img.shields.io/badge/pandas-1.5.3%2B-green)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1.2%2B-yellow)

A comprehensive QA automation solution for Beeswax digital advertising campaigns, featuring a Streamlit web application and multiple specialized QA modules for campaign validation.

## 🚀 Features

- **Interactive Web Application**: Streamlit-based interface for uploading campaign briefs and generating QA reports
- **Brief QA Validation**: Automatic detection of issues in campaign briefs before processing
- **Campaign Data Extraction**: Automated extraction of campaign data from Beeswax APIs
- **Comprehensive QA Checks**: 
  - Flight date verification
  - Naming taxonomy validation
  - Targeting settings validation
  - Creative settings verification
  - Geo-targeting validation
  - Viewability requirements checking
- **Consolidated Reporting**: All QA checks combined into a single multi-tabbed Excel report

## 📊 Project Structure

```
Beeswax_QA_Automation/
├── .streamlit/                  # Streamlit configuration
├── Brief/                       # Example campaign brief files
├── input_folder/                # Credentials and configuration files
├── output_folder/               # Generated QA reports (automated output)
├── output_raw/                  # Intermediate processing files
│
├── qa_automation.py             # Main Streamlit web application
├── run_qa.py                    # Core QA coordination script
├── brief.py                     # Brief validation module
├── beeswax_api.py               # API communication module
│
├── targeting.py                 # Targeting validation module
├── targeting_general.py         # General targeting checks
├── creative.py                  # Creative validation module
├── qa_flight_v3.py              # Flight date verification module
├── name_assign.py               # Naming taxonomy validation module
├── brief_extractor.py           # Brief data extraction utilities
│
├── requirements.txt             # Python dependencies
├── .gitignore                   # Git exclusion patterns
└── README.md                    # Project documentation
```

## 🔄 Workflow

The QA process follows this workflow:

1. **Brief Upload**: Upload campaign brief through the Streamlit interface or specify path in environment variables
2. **Brief Validation**: The `brief.py` module validates the brief for errors and inconsistencies
3. **Data Extraction**: The `beeswax_api.py` module connects to Beeswax API and fetches campaign data
4. **QA Processing**: The `run_qa.py` orchestrates specialized validation modules:
   - `qa_flight_v3.py`: Validates flight dates between brief and API data
   - `name_assign.py`: Verifies naming conventions for campaigns, line items, and creatives
   - `targeting.py`: Validates targeting settings like geo, operating systems, and content categories
   - `creative.py`: Checks creative settings against specifications
5. **Report Generation**: All results are compiled into a comprehensive Excel report with multiple tabs

## 🔧 Installation

### Prerequisites

- Python 3.6 or higher
- Git (optional, for cloning)

### Setup Steps

1. Clone or download this repository:
   ```bash
   git clone https://github.com/Guru-Ragaventhira/Beeswax_QA_Automation.git
   cd Beeswax_QA_Automation
   ```

2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure credentials by creating a `.env` file in the `input_folder` directory:
```
# API credentials
LOGIN_EMAIL=your_email@example.com
PASSWORD=your_password

# API endpoints
LOGIN_URL=https://catalina.api.beeswax.com/rest/v2/authenticate
CAMPAIGN_URL=https://catalina.api.beeswax.com/rest/v2/campaigns
LINEITEM_URL=https://catalina.api.beeswax.com/rest/v2/line-items
CREATIVE_URL=https://catalina.api.beeswax.com/rest/v2/creatives
CREATIVE_LINEITEM_URL=https://catalina.api.beeswax.com/rest/v2/line-items/{line_item_id}/creatives
ADVERTISER_URL=https://catalina.api.beeswax.com/rest/v2/advertisers/{id}
LINEITEM_EXPORT_URL=https://catalina.api.beeswax.com/rest/v2/line-items/export
SEGMENT_URL=https://catalina.api.beeswax.com/rest/v2/ref/segment-tree

# File and directory paths
BRIEF_PATH=./Brief/Campaign_Brief.xlsx
OUTPUT_DIR=./output_folder
   ```

## 💻 Usage

### Running the Web Application

Launch the Streamlit web application:

```bash
streamlit run qa_automation.py
```

The application will be available at `http://localhost:7777` in your web browser.

### Using the Web Interface

1. Connect to your company VPN (if accessing Beeswax API)
2. Navigate to `http://localhost:7777` in your browser
3. Upload a campaign brief file (Excel or CSV format)
4. Review brief validation results
5. Click "Generate QA Report" to initiate the full QA process
6. Download the resulting QA report

### Running from Command Line

For automated workflows, you can also run the QA process from the command line:

```bash
python run_qa.py --brief path/to/brief.xlsx --output path/to/output
```

## 🔒 Security

- Credentials are managed through environment variables and never committed to the repository
- The `.gitignore` file is configured to exclude sensitive files such as `.env` files
- Streamlit secrets can be used for secure credential management in deployed environments

## 📋 Detailed Module Descriptions

### Brief Validation (`brief.py`)

Performs preliminary checks on campaign briefs, including:
- Date format validation
- Budget calculation verification
- Geo-targeting specification checks
- Creative placement validation
- Viewability requirement consistency

### Beeswax API Integration (`beeswax_api.py`)

Handles all communication with the Beeswax API:
- Authentication and session management
- Campaign data retrieval by alternative ID
- Line item and creative data extraction
- Advertiser information gathering
- Targeting settings extraction

### Flight Date Verification (`qa_flight_v3.py`)

Ensures consistency between brief specifications and implemented flight dates:
- Validates campaign start/end dates
- Verifies line item flight dates match brief specifications
- Identifies discrepancies in date configurations
- Generates a detailed flight date comparison report

### Naming Taxonomy Validation (`name_assign.py`)

Checks adherence to naming conventions:
- Campaign, line item, and creative name format verification
- Quarter and year designation validation
- Product type and platform prefix checking
- Media type code verification
- Geo-targeting indicator validation
- Viewability requirement annotation checking

### Targeting Validation (`targeting.py` & `targeting_general.py`)

Comprehensive validation of targeting settings:
- Country validation
- App bundle and domain list verification
- Environment type checking
- Operating system validation
- Device type verification
- Content category validation
- Segment usage verification
- Geo-targeting implementation checking

### Creative Validation (`creative.py`)

Verification of creative settings against brief requirements:
- Creative format validation
- Size and dimension verification
- Third-party tracking validation
- Creative-to-line-item assignment checking
- Video creative duration validation


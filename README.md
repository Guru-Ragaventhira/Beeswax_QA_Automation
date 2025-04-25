# Beeswax QA Automation Tool

This tool automates the QA process for Beeswax campaigns by extracting data from campaign briefs and API endpoints.

## Setup

1. Clone or download this repository
2. Create or modify the `.env` file in the `input_folder` directory
3. Make sure you have all the required Python packages installed:
   ```
   pip install pandas requests python-dotenv openpyxl
   ```

## Environment Variables

The script uses a `.env` file to store API credentials and file paths. Here's an example of what should be in your `.env` file:

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
ENV_PATH=./input_folder/beeswax_input_qa.env
```

## Usage

### Basic Usage

Run the script without any arguments to use the paths specified in the `.env` file:

```
python test_api.py
```

The script automatically looks for the `.env` file in the following locations (in order):

1. `./beeswax_input_qa.env` (current directory)
2. `./input_folder/beeswax_input_qa.env`

You don't need to specify paths each time - just set them up once in your `.env` file and then run the script directly.

### Command Line Arguments (Optional)

Command line arguments are available for one-off runs with different settings, but they're completely optional:

```
python test_api.py --env /path/to/custom.env --brief /path/to/brief.xlsx --output /path/to/output
```

Options:
- `--env`: Override the default .env file location
- `--brief`: Override the brief path specified in the .env file
- `--output`: Override the output directory specified in the .env file

### Relative Paths

The script supports relative paths, which makes it more portable. For example:

```
BRIEF_PATH=./Brief/Campaign_Brief.xlsx
```

This path is relative to where you run the script from, not where the script is located.

## Output

The script generates an Excel file with three sheets:
1. Consolidated Report - Contains all campaign, line item, and creative data
2. Targeting Data - Contains targeting information for each line item
3. Segment Data - Contains segment information for each line item

The output file will be named with a timestamp: `qa_report_YYYYMMDD_HHMMSS.xlsx`

## Running the App Locally

1. Connect to the company VPN
2. Double-click `run_qa_app.bat` to start the application
3. The app will be available at: http://localhost:8501

## Accessing the App from Another Computer

If someone has the app running on their machine, you can access it from your computer:

1. Connect to the company VPN
2. Ask the person hosting the app for their IP address (displayed when they start the app)
3. In your browser, navigate to: `http://<their-ip-address>:8501`

## Troubleshooting

If you can't connect to someone else's app:

1. Ensure both you and the host are connected to the company VPN
2. Check that the host's firewall allows inbound connections on port 8501
3. Verify you're using the correct IP address

## Hosting Instructions

If you're hosting the app for others:

1. Run `run_qa_app.bat`
2. When the app starts, it will display your IP address
3. Share this IP address with your team members
4. Ensure your firewall allows incoming connections on port 8501
   - In Windows Defender Firewall, add an inbound rule for port 8501
   - Alternatively, temporarily disable your firewall while sharing the app (not recommended)

# Campaign QA Automation Tools

This repository contains tools for automating QA checks on digital advertising campaigns.

## Tools

### 1. Flight Date Verification Tool (`qa_flight_v3.py`)

This tool automatically verifies flight dates between Campaign Brief and QA Reports to ensure consistency in campaign configurations.

#### Purpose

The Flight Date Verification tool helps QA analysts to:

1. Cross-check campaign start/end dates between QA reports and campaign briefs
2. Verify line item flight dates against brief specifications
3. Identify any discrepancies in date configurations
4. Generate a comprehensive report highlighting matching and mismatching dates

#### How It Works

The tool performs the following steps:

1. Loads the QA report and Campaign Brief files
2. Extracts campaign-level flight dates from the brief (cells C15 and C16)
3. Compares campaign start/end dates between both sources
4. Locates target-level data section in the brief (using BV ID, BVP, and BVT headers)
5. Maps BVT IDs to BVP IDs from the target data
6. Finds placement data section in the brief (typically starting around row 27)
7. Extracts start/end dates for each BVP ID from the placement data
8. Maps line item alternative IDs (BVT) in the QA report to their corresponding BVP dates
9. Compares line item flight dates between QA report and brief
10. Generates an Excel report with all comparisons and a summary of matches/mismatches

#### Usage

Run the script with Python:

```bash
python qa_flight_v3.py
```

### 2. Naming Taxonomy Verification Tool (`name_assign.py`)

This tool checks the naming taxonomy of campaigns, line items, and creatives against standard conventions and brief requirements.

#### Purpose

The Naming Taxonomy Verification tool helps QA analysts to:

1. Ensure campaign, line item, and creative names follow required format conventions
2. Verify names contain proper quarter and year designations
3. Check for product type, platform, and media type indicators in names
4. Validate geo-targeting indicators in names match brief specifications
5. Verify viewability requirements are reflected in naming
6. Ensure consistency between line item and creative naming

#### How It Works

The tool performs the following steps:

1. Loads the QA report and Campaign Brief files
2. Extracts key information from the brief (product type, viewability requirements, etc.)
3. Identifies campaign year from the flight dates in the brief
4. Maps BVT IDs to BVP IDs from the brief's target data
5. Extracts platform, media type, and geo-targeting information for each BVP
6. Analyzes campaign names against naming conventions and brief requirements
7. Analyzes line item names, checking for correct platform prefixes and media type codes
8. Analyzes creative names, ensuring they match the corresponding line item conventions
9. Generates a comprehensive Excel report with all naming issues identified
10. Provides a summary of compliance across all entities

#### Naming Rules Checked

- No spaces or special characters (except underscores)
- Quarter designation (e.g., _Q1_, _Q2_, etc.)
- Year designation (_2025 or _25)
- Product type short form (e.g., _CTV_, _SBV_, _VMR_)
- INFMT designation for HUB campaigns
- Viewability requirements (e.g., _65_Viewability_)
- Geo-targeting designation (_Geo_ or _GEO_)
- Platform prefixes (MO_, DE_, CTV_)
- Media type codes (_BA_, _RM_, _VI_)

#### Usage

Run the script with Python:

```bash
python name_assign.py
```

## File Structure

- `qa_flight_v3.py`: Flight date verification script
- `name_assign.py`: Naming taxonomy verification script
- `qa_report.xlsx`: QA report containing campaign and line item information
- `Campaign_Brief_Bimbo.xlsx`: Campaign brief with target and placement data
- `qa_flight_v3_output.xlsx`: Output file for flight verification results
- `name_assign_output.xlsx`: Output file for naming verification results

## Requirements

- Python 3.6+
- pandas
- numpy
- openpyxl (for Excel file handling)

## Notes

- Both tools include robust error handling and fallback mechanisms for different brief formats
- Dates are compared ignoring time components (only the date part is considered)

# Environment-Based Script Configuration

All QA scripts now use a common `.env` file for configuration, making it easier to manage paths and settings in one place. The scripts include:

1. **test_api.py** - For fetching data from the Beeswax API
2. **qa_flight_v3.py** - For validating flight dates 
3. **name_assign.py** - For checking naming conventions
4. **creative.py** - For validating creative attributes

## Integrated Features

### Common Configuration

All scripts now:
- Read from the same `.env` file
- Support the same folder structure
- Use the same input files when possible
- Generate output to a consistent location

### Automatic Latest QA Report Detection

The scripts automatically detect and use the most recently generated QA report if one isn't explicitly specified. This workflow allows you to:

1. Run `test_api.py` to generate a new QA report
2. Run any of the other scripts without manually specifying the path to the QA report

### Environment File Structure

The `.env` file contains all necessary paths:

```
# API credentials
LOGIN_EMAIL=your_email@example.com
PASSWORD=your_password

# API endpoints
LOGIN_URL=...
CAMPAIGN_URL=...
# ...other API endpoints...

# File and directory paths
BRIEF_PATH=./Brief/Campaign_Brief.xlsx
OUTPUT_DIR=./output_folder
QA_REPORT_PATH=./qa_report.xlsx

# Script-specific output paths
QA_FLIGHT_OUTPUT_PATH=./qa_flight_v3_output.xlsx
NAME_ASSIGN_OUTPUT_PATH=./name_assign_output_v2.xlsx
CREATIVE_OUTPUT_PATH=./creative_qa_output.xlsx
```

## Recommended Workflow

1. Place the campaign brief in the Brief directory
2. Run `test_api.py` to generate a QA report
3. Run the validation scripts (`qa_flight_v3.py`, `name_assign.py`, `creative.py`) in any order
4. Review results in the output files

You can repeat steps 3-4 as needed for further analysis without having to update file paths.

# Beeswax QA Automation

A Streamlit application for automating QA processes for Beeswax campaigns.

## Docker Deployment

This application can be deployed using Docker, making it easy to share with your team without worrying about dependencies or environment setup.

### Prerequisites

- Docker and Docker Compose installed on your machine
- The `.env` file with your Beeswax credentials

### Quick Start

1. Place your `.env` file in the `input_folder` directory with the name `beeswax_input_qa.env`

2. Build and run the Docker container:

```bash
docker-compose up -d
```

3. Access the application:
   - If running locally: http://localhost:8501
   - If running on a server: http://[server-ip]:8501

### Stopping the Application

```bash
docker-compose down
```

## Deployment on a Shared Server

For team access, you can deploy this Docker container on a shared server:

1. Clone this repository on the server
2. Set up the environment file
3. Run with Docker Compose
4. Share the URL with your team

## Volume Mounts

The Docker setup includes two volume mounts:
- `./input_folder`: For the environment file and input data
- `./output_folder`: Where generated reports will be saved

This means files placed in these folders on your host machine will be accessible to the application, and reports generated by the application will be saved back to your host machine. 
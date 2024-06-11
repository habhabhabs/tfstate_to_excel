Alex KM 2024
MIT License

# Terraform State to Excel Extractor

## Aim
The aim of this project is to extract data from a Terraform state file (`.tfstate`) and convert it into an Excel workbook with resources organized by sheets. It allows users to include or exclude data sources (excluded by default).

## Configuration
- **Terraform State File**: The path to the `.tfstate` file can be specified.
  
- **Include Data Sources**: Option to include data sources in the extraction (disabled by default).

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/tfstate_to_excel.git
    cd tfstate_to_excel
    ```

2. Run the installation script:
    ```sh
    ./install.sh
    ```

## Usage

```sh
python tfstate_excel_extractor.py --help

```

## Example

```sh
python tfstate_excel_extractor.py --tfstate terraform.tfstate --include-data-sources --output terraform_state.xlsx
```

## Options
- `--tfstate` : The path to the `.tfstate` file (required).

- `--include-data-sources` : Include data sources in the extraction (optional).

- `--output` : Path to the output Excel file (optional, defaults to terraform_state.xlsx in the current directory).

## Project Structure
```
tfstate_to_excel/
├── README.md
├── install.sh
├── requirements.txt
├── tfstate_excel_extractor.py
└── setup.py
```

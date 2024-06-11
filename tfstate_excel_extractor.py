import json
import openpyxl
import argparse
from openpyxl.utils import get_column_letter

def safe_str(value):
    """Convert a value to a string, handling non-string types gracefully."""
    if isinstance(value, (dict, list)):
        return json.dumps(value)
    return str(value)

def safe_sheet_title(title):
    """Ensure the sheet title is within Excel's 31-character limit."""
    return title[:31]

def main(tfstate_path, include_data_sources):
    try:
        # Load the Terraform state file
        with open(tfstate_path, 'r') as file:
            state = json.load(file)

        # Create an Excel workbook
        wb = openpyxl.Workbook()
        default_ws = wb.active
        default_ws.title = 'Summary'

        # Dictionary to hold worksheets by resource type
        sheets = {}

        # Dictionary to track all unique attribute keys per resource type
        attribute_keys = {}

        # First pass: Collect all unique attribute keys
        for resource in state['resources']:
            try:
                # Filter to include only resources (exclude data sources)
                if resource.get('mode') != 'managed' and not include_data_sources:
                    continue

                resource_type = resource['type']
                
                if resource_type not in attribute_keys:
                    attribute_keys[resource_type] = set()
                
                for instance in resource.get('instances', []):
                    attributes = instance.get('attributes', {})
                    attribute_keys[resource_type].update(attributes.keys())
            except KeyError as e:
                print(f"Key error while processing resource {resource}: {e}")
            except Exception as e:
                print(f"Unexpected error while processing resource {resource}: {e}")

        # Second pass: Create sheets and write data
        for resource in state['resources']:
            try:
                # Filter to include only resources (exclude data sources)
                if resource.get('mode') != 'managed' and not include_data_sources:
                    continue

                resource_type = resource['type']
                resource_name = resource['name']
                
                # Get or create a worksheet for the resource type
                sheet_title = safe_sheet_title(resource_type)
                if sheet_title not in sheets:
                    ws = wb.create_sheet(title=sheet_title)
                    sheets[sheet_title] = ws
                    # Define the headers for the new sheet
                    headers = ['Resource Name', 'Resource Address'] + sorted(attribute_keys[resource_type])
                    ws.append(headers)
                else:
                    ws = sheets[sheet_title]
                
                # Iterate through instances within the resource
                for instance in resource.get('instances', []):
                    resource_address = instance.get('address', 'N/A')
                    attributes = instance.get('attributes', {})
                    
                    row = [resource_name, resource_address]
                    for key in sorted(attribute_keys[resource_type]):
                        row.append(safe_str(attributes.get(key, '')))
                    
                    ws.append(row)
            except KeyError as e:
                print(f"Key error while processing resource {resource}: {e}")
            except Exception as e:
                print(f"Unexpected error while processing resource {resource}: {e}")

        # Adjust column widths for each sheet
        for ws in sheets.values():
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter  # Get the column name
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = adjusted_width

        # Save the workbook
        wb.save('terraform_state.xlsx')

    except FileNotFoundError:
        print(f"The '{tfstate_path}' file was not found. Please ensure the file is present in the specified directory.")
    except json.JSONDecodeError:
        print("Failed to decode JSON. Please ensure the 'terraform.tfstate' file is a valid JSON.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='(2024) Alex KM MIT License Extract Terraform state to an Excel workbook.')
    parser.add_argument('--tfstate', required=True, help='Path to the Terraform state (.tfstate) file.')
    parser.add_argument('--include-data-sources', action='store_true', help='Include data sources in the extraction.')

    args = parser.parse_args()
    main(args.tfstate, args.include_data_sources)


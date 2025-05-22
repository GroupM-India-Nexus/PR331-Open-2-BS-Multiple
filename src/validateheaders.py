from openpyxl import load_workbook
import logging

logger = logging.getLogger(__name__)

def validate_headers(input_file, required_headers):
 """
 Validates that the Excel file has the required headers in the correct order up to 'Lang'.
 Returns (True, []) if valid, or (False, list_of_issues) if invalid.
 """
 try:
     # Load the workbook and select the active sheet
     workbook = load_workbook(input_file, data_only=True)
     sheet = workbook.active

     # Extract actual headers from the first row
     actual_headers = [str(cell.value).strip().lower() for cell in sheet[1] if cell.value]
     required_clean = [h.strip().lower() for h in required_headers]

     # Check if 'Lang' exists in both actual and required headers
     try:
         lang_index_required = required_clean.index("lang")
         lang_index_actual = actual_headers.index("lang")
     except ValueError:
         logger.error(f"File '{input_file}': Missing critical header: 'Lang'")
         return False, [f"Missing critical header: 'Lang'"]

     # Validate headers up to and including 'Lang'
     expected_before_lang = required_clean[:lang_index_required + 1]
     actual_before_lang = actual_headers[:lang_index_actual + 1]

     mismatches = []
     for i, (actual, expected) in enumerate(zip(actual_before_lang, expected_before_lang)):
         if actual != expected:
             mismatches.append(f"Position {i+1}: expected '{expected}', got '{actual}'")

     # Check for any unexpected headers before or at 'Lang'
     if len(actual_before_lang) > len(expected_before_lang):
         unexpected = actual_before_lang[len(expected_before_lang):]
         mismatches.append(f"Unexpected headers before 'Lang': {unexpected}")

     if mismatches:
         logger.error(f"\u274c File '{input_file}': Header validation failed:\n" + "\n".join(mismatches))
         return False, mismatches

     logger.info(f"\u2705 File '{input_file}': Header check passed.")
     return True, []

 except Exception as e:
     logger.exception(f"An error occurred while validating headers for file '{input_file}'.")
     return False, [str(e)]
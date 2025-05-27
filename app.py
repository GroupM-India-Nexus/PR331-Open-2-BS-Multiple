from flask import Flask, request, jsonify
import os
import base64
import logging
from src.mbs import process_tv_commercial_data
from src.validateheaders import validate_headers

app = Flask(__name__)

# Define the relative data folder
data_folder = os.path.join(os.path.dirname(__file__), "data")

# Configure logging
VERBOSE_LOGGING = True  # Toggle this flag to control verbose logging

log_level = logging.DEBUG if VERBOSE_LOGGING else logging.INFO
logging.basicConfig(level=log_level, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def is_base64_encoded(data):
  try:
      base64.b64decode(data)
      return True
  except Exception:
      return False

@app.route('/process_excelfile', methods=['POST'])
def process_excelfile():
  try:
      # Parse and log the incoming JSON body
      body = request.get_json()
      if VERBOSE_LOGGING:
          logger.debug("Request JSON body: %s", body)

      # Check if the request contains a single file or multiple files
      if "files" in body and isinstance(body["files"], list):
          # Multiple files format
          files = body["files"]
      else:
          # Single file format - convert to list format for consistent processing
          files = [body]

      if not files:
          msg = "No files to process"
          logger.error(msg)
          return jsonify({"error": msg}), 400

      # Define required headers
      REQUIRED_HEADERS = [
          "Channels", "Dur", "COMMERCIAL_FROM_DT", "COMMERCIAL_TO_DT", "Program",
          "RO Start Date", "RO End Date", "CAMPAIGN_NAME", "COMMERCIAL_NAME",
          "BRAND_NAME", "IB_NO", "Net Outlay", "GRP", "TVR", "Lang"
      ]

      results = []  # To store the results for each file

      # Process each file in the request
      for file_info in files:
          filename = file_info.get("xlsx-name")
          attach_body = file_info.get("attach-body")
          logger.info("Processing file: %s", filename)

          if not attach_body or not filename:
              msg = f"File '{filename}': Missing 'attach-body' or 'xlsx-name'"
              logger.error(msg)
              results.append({"filename": filename, "status": "error", "message": msg})
              continue

          content = attach_body.get("contentBytes")
          if not content:
              msg = f"File '{filename}': Missing 'contentBytes' in attach-body"
              logger.error(msg)
              results.append({"filename": filename, "status": "error", "message": msg})
              continue

          if not is_base64_encoded(content):
              msg = f"File '{filename}': Invalid base64 content"
              logger.error(msg)
              results.append({"filename": filename, "status": "error", "message": msg})
              continue

          # Decode and save the input Excel file
          decoded_bytes = base64.b64decode(content)
          input_path = os.path.join(data_folder, filename)
          os.makedirs(data_folder, exist_ok=True)

          with open(input_path, "wb") as f:
              f.write(decoded_bytes)
          logger.info("File saved successfully to: %s", input_path)

          # Validate headers
          is_valid, issues = validate_headers(input_path, REQUIRED_HEADERS)
          if not is_valid:
              logger.error(f"File '{filename}': Header validation failed.")
              results.append({
                  "filename": filename,
                  "status": "error",
                  "message": "Header validation failed",
                  "details": issues
              })
              continue

          # Process the file
          try:
              # Create the output file path
              output_file = os.path.join(data_folder, f"Output_{filename}")
              
              # Process the file - make sure mbs.py is updated to use this path
              process_tv_commercial_data(input_path, output_file)
              
              # Verify the output file exists
              if not os.path.exists(output_file):
                  raise FileNotFoundError(f"Output file not created: {output_file}")
              
              logger.info(f"File '{filename}' processed successfully. Output saved to {output_file}")

              # Read the processed file and encode it to Base64
              with open(output_file, "rb") as f:
                  output_data = f.read()
              encoded_output = base64.b64encode(output_data).decode('utf-8')

              # Add success result
              results.append({
                  "filename": filename,
                  "status": "success",
                  "data": encoded_output,
                  "output_filename": os.path.basename(output_file)
              })

          except Exception as e:
              logger.exception(f"An error occurred while processing file '{filename}'.")
              results.append({
                  "filename": filename,
                  "status": "error",
                  "message": str(e)
              })

      # Return the results for all files
      return jsonify(results)

  except Exception as e:
      logger.exception("An error occurred while processing the request.")
      return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
  app.run(debug=True)
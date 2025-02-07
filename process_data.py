import pandas as pd
import openpyxl
import re
import sys

def process_data(input_file):
  try:
    # 엑셀 파일 읽기
    df = pd.read_excel(input_file)

    # 필수 컬럼 체크
    required_cols = {"Year", "Make", "Model", "Trim", "Engine", "Notes"}
    if not required_cols.issubset(df.columns):
      raise ValueError("Missing required columns in the file.")

    # Trim과 Engine 데이터 처리 함수
    def clean_trim(trim):
      """Extract submodel, body type, and body number from Trim."""
      
      if not isinstance(trim, str):
        return None, None, None  # case where trim isn't a string

      # Split the trim string
      parts = trim.split()

      submodel = " ".join(parts[:-2])
      body_type = parts[-2]

      # Search for a number before the hyphen
      match = re.match(r'(\d)-', parts[-1])
      if match:
        body_number = int(match.group(1))  # Extract the number before the hyphen
      
      return submodel, body_type, body_number

    def clean_engine(engine):
      """Extracts engine specifications into separate attributes without units."""
      if not isinstance(engine, str):
          return None, None, None, None, None, None, None
      
      # Extract liters (e.g., '2.0L' → '2.0')
      liter_match = re.search(r'(\d+\.\d+)L', engine)
      liters = liter_match.group(1) if liter_match else None

      # Extract CC (e.g., '1998CC' → '1998')
      cc_match = re.search(r'(\d{3,5})CC', engine)
      cc = cc_match.group(1) if cc_match else None

      # Extract CID (e.g., '122Cu. In.' → '122')
      cid_match = re.search(r'(\d+)Cu\. In\.', engine)
      cid = cid_match.group(1) if cid_match else None

      # Extract Cylinders (e.g., 'l4', 'V6' → '4', '6')
      cyl_match = re.search(r'(?:l|V)(\d+)', engine, re.IGNORECASE)
      cylinders = cyl_match.group(1) if cyl_match else None  # Only the number

      # Extract Fuel Type (Assuming it’s always 'GAS' for now)
      fuel_type = "GAS" if "GAS" in engine else None

      # Extract Cylinder Head Type (e.g., 'DOHC', 'SOHC', 'OHV')
      head_match = re.search(r'(DOHC|SOHC|OHV|OHC)', engine, re.IGNORECASE)
      cylinder_head_type = head_match.group(0) if head_match else None

      # Extract Aspiration (e.g., 'Turbocharged', 'Naturally Aspirated')
      aspiration_match = re.search(r'(Turbocharged|Naturally Aspirated)', engine, re.IGNORECASE)
      aspiration = aspiration_match.group(0) if aspiration_match else None

      return liters, cc, cid, cylinders, fuel_type, cylinder_head_type, aspiration


    # Apply functions
    df[['Submodel', 'Body Type', 'Body Number']] = df['Trim'].apply(lambda x: pd.Series(clean_trim(x)))
    df[['Liters', 'CC', 'CID', 'Cylinders', 'Fuel Type', 'Cylinder Head Type', 'Aspiration']] = df['Engine'].apply(lambda x: pd.Series(clean_engine(x)))

    # 수정된 데이터를 새로운 시트에 저장
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a') as writer:
      df.to_excel(writer, sheet_name='ModifiedData', index=False)
      print(f"Data processed and saved to a new sheet 'ModifiedData' in {input_file}")
    
  except Exception as e:
    print(f"An error occurred while processing the data: {e}")
    sys.exit()

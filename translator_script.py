import pandas as pd
from googletrans import Translator
import zhon
import openpyxl
import time
import os

def is_chinese_char(char):
    return char >= '\u4e00' and char <= '\u9fff'

def translate_chinese_to_english(text):
    max_retries = 3
    retry_delay = 5  # seconds

    try:
        # Check if the text contains Chinese characters
        if any(is_chinese_char(char) for char in text):
            translator = Translator()

            for _ in range(max_retries):
                try:
                    translation = translator.translate(text, src='zh-CN', dest='en')
                    return translation.text
                except Exception as e:
                    print(f"Error translating text: {text}\n{e}")

                    if "The read operation timed out" in str(e):
                        print(f"Retrying in {retry_delay} seconds...")
                        time.sleep(retry_delay)
                        continue
                    else:
                        return None
            else:
                print(f"Max retries reached. Unable to translate text: {text}")
                return None
        else:
            return text  # Return original text if it doesn't contain Chinese characters
    except Exception as e:
        print(f"Error translating text: {text}\n{e}")
        return None

def xls_to_csv(input_file, output_file, delimiter='|'):
    df = pd.read_excel(input_file)
    df.replace('', pd.NA, inplace=True)
    df.to_csv(output_file, sep=delimiter, index=False, na_rep='NaN')

def chineseCSV_to_englishCSV(input_file, output_file, delimiter='|'):
    df = pd.read_csv(input_file, sep=delimiter)

    # Translate headers
    df.columns = df.columns.map(lambda x: translate_chinese_to_english(x) if pd.notna(x) else x)

    # Translate values in each cell
    for col in df.columns:
        df[col] = df[col].map(lambda x: translate_chinese_to_english(str(x)) if pd.notna(x) else x)

    df.to_csv(output_file, sep=delimiter, index=False)

def englishCSV_to_xlsx(input_file, output_file, delimiter='|'):
    df = pd.read_csv(input_file, sep=delimiter)

    # Write to Excel with original headers and handling NaN values
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, na_rep='', header=True)

        # Access the XlsxWriter workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Set column widths
        for col_num, value in enumerate(df.columns.values):
            max_len = max(df[value].astype(str).apply(len).max(), len(value)) + 2
            worksheet.set_column(col_num, col_num, max_len)

def delete_intermediate_files(csv_file, eng_csv_file):
    try:
        os.remove(csv_file)
        os.remove(eng_csv_file)
        print(f"Intermediate files '{csv_file}' and '{eng_csv_file}' deleted successfully.")
    except FileNotFoundError:
        return
    except Exception as e:   
        return
    
def main():
    xls_input_file = 'Order_Export.xls'
    csv_output_file = 'CSV_Order_Export.csv'
    english_csv_output_file = 'Eng_Order_Export.csv'
    xlsx_output_file = 'English_Order_Export.xlsx'

    # Convert Excel to CSV
    xls_to_csv(xls_input_file, csv_output_file)

    # Convert Chinese CSV to English CSV
    chineseCSV_to_englishCSV(csv_output_file, english_csv_output_file)

    # Convert English CSV to Excel
    englishCSV_to_xlsx(english_csv_output_file, xlsx_output_file)

    # Delete intermediate files
    delete_intermediate_files(csv_output_file, english_csv_output_file)
            
    print(f"Translation of chinese files Order_Export.xls to English_Order_Export.xlsx has been done successfully.")

if __name__ == "__main__":
    main()

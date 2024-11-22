# mmd_diagnostic
The Parser function reads any PDFs located within a read_files folder, converts them into text using OCR, and parses them according to the criteria below to produce an Excel file. Once all the files have been parsed, the function merges all the Excel files to produce a file called 'combined_sheets_with_hyperlinks.xlsx'.

These are the details the program goes and collects: 
'etf_number', 'claim_id', 'patient_name', 'provider_id', 'date_of_service', 'payer_name', 'payer_id', 'billed_amount', 'amount_allowed', 'paid_amount', 'patient_responsibility', 'denial_code, denial_reason_description', 'adjustment_code', 'adjustment_reason_desc', 'appeal_status', 'appeal_date', 'appeal_decision_date', 'appeal_outcome', 'remittance_date', 'service_line_details'.

# ANw_spendenbescheinigungen
Generator for Spendenbescheinigungen for the non-profit [Aktion Not wenden](https://www.aktionnotwenden.de)

![app_v0.1.0](https://github.com/vbruder/ANw_spendenbescheinigungen/blob/main/doc/Screenshot_0_1.png)

## How to use

### Read and edit data

- Select the path for the address file (must be an .xlsx Excel file)
  - *Optional:* Enter the passwrd if the xlsx is encrypted
- Select the path for the bank statement file (must be a .csv)
- **"Load Data"** Loads the data from the bank statement and tries to match them with the data contained in the address file.
  - The *Match Score* indicates the certainty of the matching, low scores get highlighted.
  - The list can be updated by clicking on single entries and editing the fields or by adding/removing entire rows
  - After editing, the address list can be updated with the added information by clicking **"Update Address File"**

### Generate receips

- Select the template file (mut be a .docx Word document)
- Select the output directory for the generated word documents
- Select the output directory for the log files
- **"Generate Receipts"** generates the receipts for all entries in the table
  - Documents in the output directory get overwritten
  - On successful generation, an entry is added to the logfile of the respective year in the log output directory if it does not exist already

- Select the output directory for the pdf files
- **"Convert to PDFs"** converts all .docx files in the document output directory into pdfs, the pdfs get saved to the pdf output directory



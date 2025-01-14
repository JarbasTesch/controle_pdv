# Email PDV Tracking Automation

This Python script automates the process of tracking stages and sending emails for documents in different stages of the PDV (Point of Sale) process. It reads data from an Excel sheet, verifies the existence of documents in specified directories, and updates the status for each entry in the tracking sheet.

## Features

- **Stage 1 (Registration)**: Adds a new record for each "matricula" (employee ID) found in a specified directory, which is not yet registered in the tracking sheet.
- **Stage 2 (TCGC)**: Tracks documents in the "TCGC" stage, updates the status of documents that are found, and checks for existing records.
- **Stage 3 (Adesão)**: Similar to Stage 2, but for documents in the "Adesão" stage.
- **Information Update**: Updates names and email addresses from another data file based on the matricula for entries that are still in the "Registro criado" stage.

## Technologies Used

- **Python**: Programming language for the automation.
- **pandas**: Used for reading and updating Excel files.
- **os**: Used for interacting with directories and files.
- **Excel**: Used to track the stages of the documents.

## Script Functions

### `cadastrar_etapa1()`
Registers employees in the first stage of the process if they are not already listed in the tracking sheet. It scans a specified directory for document files, extracts the matricula (employee ID) from the filenames, and adds a new record to the Excel sheet for any new matriculas.

### `incluir_informacoes()`
Updates the tracking sheet with names and email addresses for employees based on a second data file (consulta_email.xlsx). If a matricula is found in the second file, it updates the corresponding entry in the tracking sheet.

### `cadastrar_etapa2()`
Registers employees in the second stage ("TCGC") of the process, checking if documents for a specific matricula are found in a specified directory. Updates the tracking sheet accordingly based on the status of each document.

### `cadastrar_etapa3()`
Registers employees in the third stage ("Adesão") of the process, similar to Stage 2, but for documents in a different directory. The status is updated based on the presence of documents.

### `enviar_email_etapaX()`
This function (and its variants) is responsible for sending an email notifying about the progress of the respective step. In other words, "send_email_step1" will handle sending emails related to the "step1" column.

## How to Use

1. Ensure the required directories and Excel files are in place:
   - `acompanhamento_email_PDV.xlsx`: This file tracks the progress of each matricula through the stages.
   - `consulta_email.xlsx`: This file contains the name and email addresses for each matricula.

2. Ensure the directories for each stage (Stage 1, Stage 2, Stage 3) contain the necessary documents:
   - Stage 1 documents are stored in a directory path defined as `etapa_um_assinados`.
   - Stage 2 documents are stored in `etapa_dois_assinados`.
   - Stage 3 documents are stored in `etapa_tres_assinados`.

3. Call the functions in the script to update the status and register the employees at each stage:
   - `cadastrar_etapa1()` to register employees in Stage 1.
   - `incluir_informacoes()` to update names and email addresses.
   - `cadastrar_etapa2()` to register employees in Stage 2.
   - `cadastrar_etapa3()` to register employees in Stage 3.

4. The script will update the `acompanhamento_email_PDV.xlsx` file with the current status for each matricula.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

If you have any questions or suggestions, feel free to contact me via [LinkedIn](https://www.linkedin.com/in/jarbastesch/) or [GitHub](https://github.com/jarbastesch).

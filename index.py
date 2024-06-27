from docx import Document 
from pandas import DataFrame
import os
# The module "openpyxl" has to be installed to generate the Excel file


##------- MANAGEMENT OF THE DATA TO EXTRACT -------##
# String found juste before the value to extract 
str_before_value = {
    "Date": "Last modification:",
    "RO": "Repair order:",
    "Reason for return": "Reasons for return",
    "MSN": "MSN Aircraft:",
}


##------- MANAGEMENT OF THE FILES -------##
# List all the files in the directory
os.chdir('CRE')
list_of_files_names = os.listdir()


# These lists will stock the final table once values are extracted from the documents
list_date = [] # Date of last modification of the document
list_RO = [] # Repair Order (Unic value, key of the database)
list_reason_for_return = [] # Reason of return given by the customer 
list_MSN = [] # MSN of the aircraft

# List of files which rised an error 

##------- MAIN LOOP -------##
def main():
    """
    Main function of the program which will execute the other ones. 
    Arg: 
        None
    Return: 
        None
    """
    i=1; extraction_error = False # Is changed to True if an error has occured during data extraction in the Word documents
    list_error_files, list_error = [], []

    ##----- FILE OPENING AND DATA STORED IN PYTHON LISTS-----##
    # Loop over all the files present in the folder to extract data and store it into the lists
    for file_name in list_of_files_names:
        try:
            document = Document(file_name)
            extract_data(document)
            print(f'Extraction des données du document {i} / {len(list_of_files_names)}    {round((i)/len(list_of_files_names)*100, 1)} %')
            i+=1

        # If an error occurs during the extraction on a document it will be ignored. 
        # The user will be informed of files which led to an error at the end of the program execution
        except Exception as error:
            list_error_files.append(list_of_files_names[i-1])
            list_error.append(error)

            extraction_error = True

            
    ##----- MANAGEMENT OF THE DUPLICATES -----##
    print(f'\n\n{" GESTION DES DOUBLONS ":-^50}')
    find_duplicates()    
    f_data = delete_duplicates() # Filtered data without duplicates
    

    ##----- CONVERT PYTHON LIST TO .TXT and EXCEL files -----##
    print(f'{" GESTION DES FICHIERS DE SORTIE ":-^50}')
    # Generation of the .txt file
    try: 
        store_data_as_txt(f_data["RO"], f_data["date"], f_data["RFR"], f_data["MSN"])
    except Exception as error:
        print('<!> Erreur lors de la génération du fichier .txt :')
        print(error)


    # Generation of the Excel file
    try:
        store_data_as_Excel(f_data["RO"], f_data["date"], f_data["RFR"], f_data["MSN"])
    except Exception as error:
        print("<!> Erreur lors de la génération du fichier Excel. Vérifier que le fichier Excel n'est pas ouvert. Message d\'erreur : ")
        print(error)


    ##----- OUTPUT OF THE ERRORS -----##
    print(f'\n\n{" GESTION DES ERREURS ":-^50}')
    # Error occured during the extraction
    if extraction_error:
        print('<!> Les documents suivants n\'ont pas pu être ouverts :   <!>')
        for i in range(len(list_error_files)):
            print(list_error_files[i])
            print(f'Cause de l\'erreur : {list_error[i]}')

        print('\nLes données de ces document n\'ont donc pas été extraites.')

    print('\n \nProgramme réalisé par Mathéo TROUILLE au service support (S9).')
    input('Appuyer sur Entrée pour quitter le programme > ')


##------- DATA EXTRACTION -------##
def extract_data(document):
    """
    Extract the required data from a Word document
    Arg: 
        document: Object of the Document class
    Return:
        None
    """
    RO_found = False 
    MSN_found = False
    RFR_found = False

    #--- DATE----#
    try:
        date = document.core_properties.modified.date()
        list_date.append(date)
    except:
        list_date.append("N.A." + " ")


    for p in document.paragraphs:
        #--- RO----#
        if str_before_value["RO"] in p.text: 
            RO = p.text
            begin = RO.find(str_before_value["RO"])
            RO = (RO[begin+14::]).strip()
            
            list_RO.append("Z" + RO)
            RO_found = True

        #--- RFR ----#
        if str_before_value["Reason for return"] in p.text: 
            RFR = p.text
            begin = RFR.find(str_before_value["Reason for return"])
            RFR = (RFR[begin+20::]).strip()
            
            list_reason_for_return.append(RFR + " ")
            RFR_found = True


        #--- MSN ----#
        if str_before_value["MSN"] in p.text:
            MSN = p.text
            
            begin = MSN.find(str_before_value["MSN"])
            end = MSN.find("Received date")

            MSN = MSN[begin+13: end-2].strip()

            list_MSN.append(MSN + " ")
            MSN_found = True


    #--- VALUES NOT FOUND ----#
    if not RO_found:
        list_RO.append("NA" + " ")
    if not MSN_found:
        list_MSN.append("NA" + " ")
    if not RFR_found:
        list_reason_for_return.append("NA" + " ")



def find_duplicates():
    """
    Find the duplicated R.O. Indeed one R.O. can be linked to several Word documents. We need to know which R.O. are duplicated to keep only one at the end
    Arg: 
        None
    Return: 
        None 
    """
    occurences = {}

    # For each RO,
    for RO in list_RO:
        # If the RO has already been found once
        if RO in occurences:
            occurences[RO] += 1
        # If it is the first time we find the occurences
        else:
            occurences[RO] = 1

    duplicates = [RO for RO, count in occurences.items() if count > 1]

    if len(duplicates) > 0:
        print(f'<!> ATTENTION <!> Les RO suivants ({len(duplicates)} éléments) sont en doublons, seul seront gardés ceux avec la date de modification la plus récente :')
        for RO in duplicates: print(RO)
    else:
        print('Aucun doublon trouvé')    
        print(duplicates)

    print('\n')



def delete_duplicates():
    """
    In the case where the R.O. refers to several Word document, only the last modificated one is kept
    Arg:
        None
    Return: 
        (Dictionnary) filtered_data : Extracted data with no duplicated R.O. 
    """
    ids = list_RO
    dates = list_date
    data_1 = list_reason_for_return
    data_2 = list_MSN

    info_dict = {}

    for i in range(len(ids)):
        current_id = ids[i]
        current_date = dates[i]
        current_data_1 = data_1[i]
        current_data_2 = data_2[i]

        # If the ID is already in the dictionnary 
        if current_id in info_dict:
            stored_date = info_dict[current_id]['date']

            # Update only if the current date is more recent than the stored one. In the other hand we keep the previous one which is more recent.
            if current_date > stored_date:
                info_dict[current_id] = {
                    'date': current_date, 
                    'data_1': current_data_1, 
                    'data_2': current_data_2
                }

        # If the ID is not already in the dictionnary
        else:
            info_dict[current_id] = {
                    'date': current_date, 
                    'data_1': current_data_1, 
                    'data_2': current_data_2
                }


    # Storing the data into the lists used to extract 
    final_list_RO, final_date_list, final_RFR_list, final_MSN_list = [], [], [], []
    for id_, info in info_dict.items():
        final_list_RO.append(id_)
        final_date_list.append(info['date'])
        final_RFR_list.append(info['data_1'])
        final_MSN_list.append(info['data_2'])

    filtered_data = {
        "RO": final_list_RO,
        "date": final_date_list,
        "RFR": final_RFR_list,
        "MSN": final_MSN_list
    }

    return filtered_data

##------- STORING DATA -------##
def store_data_as_Excel(filtered_RO, filtered_date, filtered_RFR, filtered_MSN):
    
    # Excel file creation
    print('\nCréation du fichier Excel ...')


    # Creation of the dataframe
    df = DataFrame({
        'RO': filtered_RO,
        'Date création doc.': filtered_date,
        'Reason for return': filtered_RFR,
        'MSN': filtered_MSN 
    })

    # Creation of the Excel file
    df.to_excel('extraction.xlsx', index=False)
    print('Fichier Excel généré avec succès !')



def store_data_as_txt(filtered_RO, filtered_date, filtered_RFR, filtered_MSN):
    os.chdir('../')

    # Text file creation
    print('Création du fichier .txt...')
    with open('extraction CRE.txt', 'w', encoding='utf-8') as file:
        for RO, date, RFR, MSN in zip(filtered_RO, filtered_date, filtered_RFR, filtered_MSN):
            file.write(f'{RO}|{date}|{RFR}|{MSN}\n')
        print('Fichier txt généré avec succès !')

main()
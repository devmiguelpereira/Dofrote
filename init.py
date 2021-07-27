# external modules
from pip._vendor.distlib.compat import raw_input
from docxtpl import DocxTemplate
from docx2pdf import convert
import csv
import os

# Global Variables
dirNames = ["Generated/", "ConvertedToPDF/", "Templates/", "FileImport/"]


# Functions
def checkDirectoriesExistence():
    for dirName in dirNames:
        if not os.path.exists(dirName):
            os.mkdir(dirName)
            print("Directory ", dirName, " Created ")
        # else:
        #    print("Directory ", dirName, " already exists")

def displayAvailableTemplates(path=dirNames[2]):
    for i, file in enumerate(os.listdir(path)):
        if file[-5:] == '.docx' and file[:2] != "~$":
            print(file)


def getAvailableTemplates(path=dirNames[2]):
    filenames = list()
    for i, file in enumerate(os.listdir(path)):
        if file[-5:] == '.docx' and file[:2] != "~$":
            file_split = file.split(".")
            filenames.append(file_split[0])

    return filenames

# imports data from a csv file
def importFromFile(filename):
    set = []

    with open(filename, 'r', encoding='utf-8') as csvfile:
        valueLines = csv.reader(csvfile, delimiter=',')
        for line in valueLines:
            set.append(line)

    return set

def sortDataContext(dataset):
    dataContext = dict()
    headers = dataset[0]  # saving headers aside
    dataset.pop(0)  # deleting the first line with headers

    for index in range(len(dataset)):
        dataContext[index] = {}
        for item, header in enumerate(headers):
            dataContext[index][header] = dataset[index][item]

    return dataContext

def generateDoc(outputPath=dirNames[0], setTemplate=dirNames[2]+"template.docx", dataContext=dict()):
    try:
        for count, item in enumerate(dataContext):
            doc = DocxTemplate(setTemplate)
            doc.render(dataContext[item])
            doc.save(outputPath + "/generated_{}.docx".format(count))
            print("Documento Generated Successfully!")
    except:
        print("Erro while trying to generate de document!")

    main()

def convert2Pdf(inputPath=dirNames[0], outputPath=dirNames[1]):
    for i, file in enumerate(os.listdir(inputPath)):
        if file[-5:] == '.docx' and file[:2] != "~$":
            convert(f"{inputPath}{file}", f"{outputPath}")

def print_menu():
    print(30 * "-", "MENU", 30 * "-")
    print("1. List of Templates")
    print("2. Generate Document from csv")
    print("3. Convert Generated Documents to PDF")
    print("4. Exit")
    print(67 * "-")

# End of function declaration


def main(): # Main function that is responsible to execute and join all functions
    checkDirectoriesExistence()

    loop = True
    while loop:  ## While loop which will keep going until loop = False
        print_menu()  ## Displays menu
        choice = int(input("Enter your choice [1-5]: "))

        if choice == 1:
            print("Displaying available template")
            displayAvailableTemplates()
        elif choice == 2:
            print("Generating Document from csv")
            print("Select document template and csv data to import")

            templateName = getAvailableTemplates(dirNames[2])
            for count, template in enumerate(templateName):
                print(count, template)
            select = int(input("chose a template number option: \n"))
            loop = True
            while loop:
                if select < len(templateName):
                    dataset = importFromFile(dirNames[3] + templateName[select] +".csv")
                    generateDoc(setTemplate=dirNames[2] + templateName[select] +".docx",
                                dataContext=sortDataContext(dataset))
                    loop = False
                else:
                    raw_input("Wrong option selection. Enter any key to try again..")

        elif choice == 3:
            print("Convert Generated Documents to PDF")
            convert2Pdf()
        elif choice == 4:
            print("Terminating Application")
            loop = False  # This will make the while loop to end as not value of loop is set to False
        else:
            # Any integer inputs other than values 1-5 we print an error message
            raw_input("Wrong option selection. Enter any key to try again..")


main()

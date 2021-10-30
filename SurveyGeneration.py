from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import time
import os
import shutil
import win32com.client as win32
import tkinter as tk
from PIL import Image, ImageTk
import getpass
import webbrowser


# Function to start the entire process - takes in user input
def startSurvey(root):
    # set up the user input functionality
    surveyFrame = tk.Frame(root, height = 600, width = 600, bg='navy')
    surveyFrame.place(relx=0.5, rely=0.5, anchor="center")
    courseName_var = tk.StringVar()
    courseDate_var = tk.StringVar()
    instructorName_var = tk.StringVar()
    # input the course name
    cName_label = tk.Label(surveyFrame, text = ' Course name (intro, debug, nios, etc.):', bg='white', fg='navy', font=('calibre',10, 'bold'))
    cName_entry = tk.Entry(surveyFrame, textvariable = courseName_var, font=('calibre',10,'normal'))
    # input the course date
    cDate_label = tk.Label(surveyFrame, text = '         Course date (MM/DD/YYYY):           ',bg='white', fg='navy', font=('calibre',10, 'bold'))
    cDate_entry = tk.Entry(surveyFrame, textvariable = courseDate_var, font=('calibre',10,'normal'))
    # input the instructor name
    instName_label = tk.Label(surveyFrame, text = '                 Instructor name:                   ', bg='white', fg='navy', font=('calibre',10, 'bold'))
    instName_entry = tk.Entry(surveyFrame, textvariable = instructorName_var, font=('calibre',10,'normal'))
     
    # place the input boxes on the window
    cName_label.grid(row=0,column=0)
    cName_entry.grid(row=0,column=1)
    cDate_label.grid(row=1,column=0)
    cDate_entry.grid(row=1,column=1)
    instName_label.grid(row = 2, column = 0)
    instName_entry.grid(row = 2, column = 1)

    # button to submit user input - calls createSurvey function
    submit = tk.Button(surveyFrame, text = 'Get survey link', bg='white', fg='navy', command = lambda: createSurvey(instructorName_var, courseName_var, courseDate_var, root, surveyFrame))
    submit.grid(row=3,column=1)

    manualBypasser = tk.Button(surveyFrame, text = 'Bypass automation step', bg='white', fg='navy', command = lambda: manualBypass(instructorName_var, courseName_var, courseDate_var, root, surveyFrame))
    manualBypasser.grid(row=4, column=1)

def manualBypass(instructorName_var, courseName_var, courseDate_var, root, surveyFrame):
    courseName = courseName_var.get()
    courseDate = courseDate_var.get()
    instructorName = instructorName_var.get()

    # cleaning up user input
    if "intro" in courseName.lower():
        courseAbbrev = "IUWINTRO"
        courseName = "Introduction to FPGAs and the Intel Quartus Prime Software"
    elif "debug" in courseName.lower():
        courseAbbrev = "IUWSDBUG"
        courseName = "Introduction to Simulation and Debug of FPGAs"
    elif "nios" in courseName.lower():
        courseAbbrev = "IUWNIOS"
        courseName = "Embedded Nios Processor & Platform Designer"
    else:
        courseAbbrev = courseName[0:5]
    # cleaning up user input
    cleaned_courseDate = courseDate.replace("/","")

    # Create a new folder in the To_be_processed folder of the SharePoint
    username = getpass.getuser()
    toBeProcessed_path = f'C:\\Users\\{username}\\OneDrive - Intel Corporation\\Documents - FPGA Academic Council\\Workshops\\Workshop invite list universities\\To_be_Processed'
    os.chdir(toBeProcessed_path)
    os.makedirs(courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName)
    newFolder = toBeProcessed_path + '\\' + courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName + '\\'
    certificateGeneratorFolder = f'C:\\Users\\{username}\\OneDrive - Intel Corporation\\Documents - FPGA Academic Council\\Workshops\\Workshop invite list universities\\Certificate_Generator\\'

    # create a copy of the certificate generator template and place it in this project's folder
    for file in os.listdir(certificateGeneratorFolder):
        if file[-5:] == '.xlsm':
            shutil.copy2(certificateGeneratorFolder + file, newFolder + courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName + '.xlsm')

    #surveyLink = "Paste the link to B8 in Certificate sheet"
    webbrowser.open('https://intel.az1.qualtrics.com/catalog/')
    surveyFrame.destroy()
    bypassFrame = tk.Frame(root, height = 800, width = 1000, bg='navy')
    bypassFrame.place(relx=0.5, rely=0.5, anchor="center")

    survLink_var = tk.StringVar()
    survLink = tk.Label(bypassFrame, text = 'Anonymous Survey Link:', bg='white', fg='navy', font=('calibre',10, 'bold'))
    survLink_entry = tk.Entry(bypassFrame, textvariable = survLink_var, font=('calibre',10,'normal'))
    survLink.grid(row = 2, column = 0)
    survLink_entry.grid(row = 2, column = 1)
    #surveyLink = survLink_var.get()

    bypass = tk.Button(bypassFrame, text="Populate the Excel file", command = lambda: populateExcel(newFolder, courseAbbrev, cleaned_courseDate, instructorName, survLink_var.get(), courseDate, courseName), bg = 'blue', fg='white', height = 2, width=35)
    bypass.grid(row = 4, column = 1)
    

# creates a survey using Chrome automation in Qualtrics site
def createSurvey(instructorName_var, courseName_var, courseDate_var, root, surveyFrame):
    courseName = courseName_var.get()
    courseDate = courseDate_var.get()
    instructorName = instructorName_var.get()

    # cleaning up user input
    if "intro" in courseName.lower():
        courseAbbrev = "IUWINTRO"
        courseName = "Introduction to FPGAs and the Intel Quartus Prime Software"
    elif "debug" in courseName.lower():
        courseAbbrev = "IUWSDBUG"
        courseName = "Introduction to Simulation and Debug of FPGAs"
    elif "nios" in courseName.lower():
        courseAbbrev = "IUWNIOS"
        courseName = "Embedded Nios Processor & Platform Designer"
    else:
        courseAbbrev = courseName[0:5]
    # cleaning up user input
    cleaned_courseDate = courseDate.replace("/","")

    # Create a new folder in the To_be_processed folder of the SharePoint
    username = getpass.getuser()
    toBeProcessed_path = f'C:\\Users\\{username}\\OneDrive - Intel Corporation\\Documents - FPGA Academic Council\\Workshops\\Workshop invite list universities\\To_be_Processed'
    os.chdir(toBeProcessed_path)
    os.makedirs(courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName)
    newFolder = toBeProcessed_path + '\\' + courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName + '\\'
    certificateGeneratorFolder = f'C:\\Users\\{username}\\OneDrive - Intel Corporation\\Documents - FPGA Academic Council\\Workshops\\Workshop invite list universities\\Certificate_Generator\\'

    # create a copy of the certificate generator template and place it in this project's folder
    for file in os.listdir(certificateGeneratorFolder):
        if file[-5:] == '.xlsm':
            shutil.copy2(certificateGeneratorFolder + file, newFolder + courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName + '.xlsm')


    # "Detach" the webdriver so that it does not automatically close after the program runs
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(chrome_options=options, executable_path='C:\webdrivers\chromedriver.exe')

    # Retrieve the Qualtrics catalog page and maximize the window
    driver.get('https://intel.az1.qualtrics.com/catalog/')
    driver.maximize_window()

    # Wait for the page to complete log-in and completely load the page
    driver.implicitly_wait(15)

    # WebDriverWait waits for the given element to appear and be clickable on the screen
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[2]/div[1]/div[2]/div[1]/div[2]/div')))
    # clicks the "Survey" button under the "Projects from scratch" label
    surveyFromScratch = driver.find_element_by_xpath('//*[@id="root"]/div[2]/div[1]/div[2]/div[1]/div[2]/div')
    surveyFromScratch.click()

    # clicks the "Get started" button in the menu that comes up
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[2]/div[2]/div/div[3]/div/button')))
    getStarted = driver.find_element_by_xpath('//*[@id="root"]/div[2]/div[2]/div/div[3]/div/button')
    getStarted.click()

    # populates the "Folder" input box
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, '_2NEGN')))
    folder = driver.find_element_by_class_name('_2ZuUb').find_element_by_class_name('_2NEGN')
    folder.send_keys(Keys.CONTROL + "a")
    folder.send_keys(Keys.DELETE)
    folder.send_keys("Surveys Workshop") # the name of the folder that all the survey projects are kept in
    #folder.send_keys(Keys.ENTER)
    folder.send_keys(Keys.TAB)

    # clicks the dropdown menu to open it and make the options visible
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, '_3IhHss')))
    blankSurveyDropdown = driver.find_element_by_class_name('_3IhHss')
    blankSurveyDropdown.click()

    # finds the third option in the dropdown menu - "Copy from existing project"
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CLASS_NAME, '_3Gtyq')))
    copyFromExistingProjectList = driver.find_elements_by_class_name('_2f4-X')
    copyFromExistingProject = copyFromExistingProjectList[2]
    copyFromExistingProject.click()

    # clicks the survey dropdown to choose the template
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, '_2x14cC')))
    selectASurvey = driver.find_element_by_class_name('_2x14cC')
    selectASurvey.click()

    # uses XPATHs
    # hovers the mouse over a folder 
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div/div[2]/div[2]')))
    workshopTemplateFolder = driver.find_element_by_xpath('/html/body/div[8]/div/div[2]/div[2]')
    hover = ActionChains(driver).move_to_element(workshopTemplateFolder)
    hover.perform()

    # click the correct template using XPATH
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[9]/div/div[2]/div[2]')))
    workshopTemplate = driver.find_element_by_xpath('/html/body/div[9]/div/div[2]/div[2]')
    workshopTemplate.click()

    # enter the name of the project using the user input
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.NAME, 'name')))
    projectName = driver.find_element_by_class_name('_1fwU_').find_element_by_name('name')
    projectName.send_keys(courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName)

    # submit these settings - click "Create New Project"
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, '_2uTqs')))
    createNewProject = driver.find_element_by_class_name('_2uTqs')
    createNewProject.click()

    # wait for page to load
    time.sleep(15)

    # sometimes the page doesn't load properly, so exception handling:
    try:
        # click the publish button if it works
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'publish-button')))
        publish = driver.find_element_by_id('publish-button')
        publish.click()
    except:
        # if the page didn't load, refresh the page and try again
        driver.refresh()
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'publish-button')))
        publish = driver.find_element_by_id('publish-button')
        publish.click()

    # click the publish confirmation button
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, 'confirm-button')))
    publishConfirmList = driver.find_elements_by_class_name('confirm-button')
    publishConfirm = publishConfirmList[0]
    publishConfirm.click()

    # retrieve the anonymous survey link
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, 'published-survey-anonymous-link')))
    surveyLink = driver.find_element_by_id('published-survey-anonymous-link').text
    print(surveyLink)

    # print the survey link on the GUI to let the user know the next step is ready
    surveyLabel = tk.Label(root, text="Here is the survey link: " + surveyLink, bg='LightBlue1', fg='white', font=('calibre',10, 'bold'))
    surveyLabel.place(anchor="s", relx = 0.5, rely = 0.9)
    surveyFrame.destroy()

    # button to populate the excel file with the basic information and attendee list
    popExcelButton = tk.Button(root, text="Populate the Excel file", command = lambda: populateExcel(newFolder, courseAbbrev, cleaned_courseDate, instructorName, surveyLink, courseDate, courseName), bg = 'blue', fg='white', height = 2, width=35)
    popExcelButton.place(relx = 0.5, rely=0.7, anchor="center")
    root.deiconify() # show the GUI

# accesses the generated excel workbook and fills in the info
def populateExcel(newFolder, courseAbbrev, cleaned_courseDate, instructorName, surveyLink, courseDate, courseName):
    #surveyLink = survLink_var.get()
    # open the excel workbook
    generatorFileName = newFolder + courseAbbrev + '_' + cleaned_courseDate + '_' + instructorName + '.xlsm'
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    certifGenerator = excel.Workbooks.Open(generatorFileName)

    # place the appropriate info in the right cells of the right sheet
    certificateSheet = certifGenerator.Worksheets("Certificate")
    studentSheet = certifGenerator.Worksheets("Students")
    certificateSheet.Cells(8,2).Value = surveyLink
    certificateSheet.Cells(7,2).Value = courseDate
    certificateSheet.Cells(9,2).Value = instructorName
    certificateSheet.Cells(12,2).Value = instructorName
    certificateSheet.Cells(10,2).Value = courseName

    # show the workbook
    excel.Visible = True

    # call the "CopyInfo" macro in the workbook to fill in the attendees
    certifGenerator.Application.Run("CopyInfo.CopyInfo")

    # save the file just in case
    certifGenerator.Save()

    # button to preview an email to ensure that all information appears correctly
    prevEmailButton = tk.Button(root, text="Preview Email", command = lambda: previewEmail(certifGenerator), bg = 'blue', fg='white', height = 2, width=35)
    prevEmailButton.place(relx = 0.5, rely=0.7, anchor="center")
    root.deiconify()


# preview an email before it is sent to make sure the information is right
def previewEmail(certifGenerator):
    # run the "PreviewEmail" macro in the workbook
    certifGenerator.Application.Run("PreviewEmail.PreviewEmail")

    # button to send the rest of the emails to the attendees
    sendEmailsButton = tk.Button(root, text="Send all emails", command = lambda: sendEmails(certifGenerator), bg = 'blue', fg='white', height = 2, width=35)
    sendEmailsButton.place(relx = 0.5, rely=0.7, anchor="center")
    root.deiconify()

# sends emails to the rest of the attendees without displaying them first
def sendEmails(certifGenerator):
    # calls the "CreateAndSend" macro in the workbook
    certifGenerator.Application.Run("CreateAndSend.CreateAndSend")

    

## Main code body ##

# Setting up the GUI using tkinter
root = tk.Tk()
root.title("Intel Workshops - Create Survey and Send Certificates")
root.configure(background="white")

# setting the size of the window
root.geometry("600x400")

# adding color to the window
canvas = tk.Canvas(root, height = 700, width = 700, bg = 'LightBlue1')
canvas.grid(row = 0, column = 0)

# adding an Intel FPGA image to the window
# placing the image is optional - requires you to have the image on your computer
logo = ImageTk.PhotoImage(Image.open("IntelFPGA.jpg"))
logo_label = tk.Label(image=logo)
logo_label.place(x=0, y=0)

# button to start the entire process
surveyStart = tk.Button(root, text="Start Survey Creation", command = lambda: startSurvey(root), font='Raleway', bg="navy", fg="white", height=2, width=25)
surveyStart.place(relx=0.5, rely=0.5, anchor="center")


# exit button
tk.Button(root, text="Exit", command=root.destroy, bg = "blue", fg="white").place(relx=0.9, rely=0.9)

# run the window
root.mainloop()









# -*- coding: utf-8 -*-
#Important modules that need to be installed
import xlrd
import xlsxwriter
import os
from unicodedata import normalize
from difflib import SequenceMatcher
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import *
from zipfile import ZipFile
import time
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(True)

#maps school name to school id
schoolIDs = {}
#maps school name to a list of al students in the school
schoolStudents = {}

link = {}

#The rubric for each of the categories
extemp = ["Speaker (student's first and last name)", "Code", "Topic selected", "Judge's Name (first and last)", "Round",
          "Room",
          "1. Introduction – is the introduction effective in gaining the audience attention and giving direction to the speech?",
          "2. Content – are the points valid? – does the speaker show knowledge of the subject? – is the source material, if used, well-integrated into the speech?",
          "3. Organization -- is the speech easy to follow? -- was there a logical progression of ideas? – does the speaker stick to the topic?",
          "4: Language - is the vocabulary suitable for informal speaking? -- are the sentences clear and to the point?",
          "5. Conclusion – does the conclusion summarize the speaker's main points? -- does the speaker bring the speech to an effective close? ",
          "6. Use of the voice - does the voice enhance the delivery (appropriate volume, varied pitch, controlled rate, enunciation, etc.)? – does the speaker use a natural, convincing tone?",
          "7. Eye contact – does the speaker maintain effective visual contact with the audience?",
          "8. Delivery – are the face and body gestures effective? - is the delivery style natural and effective?",
          "9. Overall effect – did the speaker effectively handle the topic?",
          "Additional comments (i.e. Great job! You used your time well to organize a strong presentation!)",
          "Total Score: Max 36, Superior 32-36, Good 27-31", "Rank in Room", "Over Time"]
persua = ["Speaker (student's first and last name)", "Code", "Selection (name of the piece)",
          "Judge's Name (first and last)", "Round", "Room",
          "Topic: Is the topic of timely importance? Are there sufficient arguments to develop the topic? Is the topic persuadable? ",
          "Originality: Is the speaker's approach to the topic original and interesting?",
          "Introduction: Did the introduction gain audience attention? Did the introduction establish rapport with the audience? Did the introduction make clear the speech's purpose?  Did the introduction provide a smooth transition into the body of the speech? ",
          "Organization: Was it easy to follow and understand the organizational pattern of the speech? Was there a sufficient variety of proof used to persuade? Were the main premises substantiated? ",
          "Proof: Was the type(s) of proof chosen appropriate for the topic? Was there a sufficient variety of proof used to persuade? Were the main premises substantiated? ",
          "Conclusion: Did the conclusion bring the speech to a satisfying end? Did the conclusion reinforce the persuasive purpose? ",
          "Language: Were the word selection and sentence arrangement effective? Did the speaker use effective transitions? ",
          "Delivery: Were the facial expressions and body gestures effective? Was the delivery style natural and believable? ",
          "Eye Contact - Does the speaker maintain effective visual contact with the audience? ",
          "Vocal Production: Is the speech articulate? Is the tone persuasive? Did the voice convey sincerity? ",
          "Overall Effect: Did the speaker accomplish his/her persuasive purpose? (Should reflect all the speech's components)",
          "Total Score: Maximum 44, Superior 39-44, Good 33-38", "Rank in Room ", "Over Time"]
inform = ["Speaker (student's first and last name)", "Code", "Selection (name of the piece)",
          "Judge's Name (first and last)", "Round", "Room",
          "Subject - Is the subject appropriate for informative speaking?  Is the subject appropriate for the designated audience?  Can the subject be handled adequately in the allotted time?",
          "Introduction - Does the introduction gain audience attention?  Does the introduction establish rapport with the audience?  Did the introduction make clear the speech's purpose?  Did the introduction provide a smooth transition into the body of the speech?  ",
          "Organization (Body) - Is the outline pattern easy to follow?  Is the outline pattern appropriate for the subject matter?",
          "Conclusion - Does the conclusion adequately sum up the main points?  Does the conclusion reinforce the speech's purpose?  Does the conclusion bring the speech to a satisfying end?  ",
          "Language - Are the word selections and structural arrangement effective?  Are the transitions effective?",
          "Originality - Is the speaker's approach to the subject original?  Is the speaker's approach to the subject creative?",
          "Vocal Delivery - Is the tone appropriate for the informative purpose?  Are appropriate vocal techniques used?  Is the speech articulate?  ",
          "Nonverbal Delivery - Does the speaker's style appear natural?  Are the facial and body gestures natural and effective?",
          "Eye Contact - Does the speaker make effective use of eye contact?",
          "Overall Effect - Did the speaker accomplish his/her informative purpose? ",
          "Additional Comments (i.e. Great job! You used a lot of elements to make this an impressive performance)",
          "Total Score: Maximum 40, Superior 36-40, Good 32-36", "Rank in Room ", "Over Time"]
original = ["Speaker (student's first and last name)", "Code", "Selection (name of the piece)",
            "Judge's name (first and last)", "Round (number)",
            "Vocal Projection and Articulation - Does the voice both suit and enhance the material? Is volume sufficient, and/or are variations in volume used effectively?",
            "Tempo and Phrasing - Does the pacing augment the mood and effect (drama, humor, etc.)? Does the work flow with clarity? Does the phrasing establish the writer's thought pattern clearly?   If pauses are used, are they deliberate and effective?",
            "Gestures and Facial Expressions -  Do the gestures enhance the interpretation of the material? Are the gestures suggestive and not sustained to the point of acting? Is any movement confined to the podium area?",
            "Eye Contact - Does the speaker maintain appropriate / sufficient eye contact with the audience? If there is more than one character, is there an effective shift of focus? If a focus is used, is it clearly and effectively maintained?",
            "Narration / Characterization - Is the work coherent and easy to follow? Is a specific mood established and maintained that is appropriate to the genre? Does the speaker utilize the language of the piece effectively? Does the personality of the character / speaker come through clearly? If there are multiple characters, is each distinct?",
            "Into which genre does the student's OW piece fall?      (You will be directed to the appropriate section's genre-specific criteria.)",
            "Drama Criteria -- Is there adequate conflict and tension for dramatic effect? Is the conflict clear, and does it build? Is the selection presented as an interpretation rather than an acting performance? Is the reader able to handle the maturity of the selection? Is the reader believable?",
            "Poetry Criteria -- Do the rate and phrasing reflect the natural rhythm of the poem(s)? Do voice inflections, tone, and pitch effectively contribute to and communicate meaning? Are sense / image words emphasized to suggest mood and poetic effect? Does the speaker avoid being excessively dramatic?",
            "Prose Criteria -- Is narration a significant part of the piece? If dialogue is present, is there a smooth transition between narration and dialogue? Does the speaker's visual contact communicate the narrative point of view of the material? Does the speaker effectively convey the images / style of the piece? Does the speaker avoid being excessively dramatic?",
            "Humor Criteria -- Does the pacing, timing, and phrasing suggest the meaning / humor of the piece? Does the characterization and/or narration create effective humor? Are there sufficient light elements in the piece to make it humorous? Is the interpretation original rather than imitative?",
            "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
            "Rank in Round:",
            "Total Score: Maximum 24, Superior 21-24, Good 18-21 * (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
            "Was this student's performance over time (i.e., longer than 11 minutes, not including the introduction)?"]
children = ["Speaker (student's first and last name)", "Code", "Selection (name of the piece)",
            "Judge's Name (first and last)", "Round (number)", "Room (letter)",
            "Vocal Projection and Articulation - Does the voice both suit and enhance the material? Does the speech flow with clarity? Is the volume used effectively?",
            "Body & Facial Function - Do the gestures enhance, but not distract from the material?  Do the (visible) facial expressions effectively reinforce the selection?",
            "Eye Contact - Does the reader maintain sufficient visual contact with the audience?  Does the reader maintain a distinct focus for each character?",
            "Characterization - Does the personality of the character(s) come through?  Is the reader's interpretation believable?  If there is more than one character, are they distinct?  Is the narrator distinguished from the other characters?",
            "Selection of Material - Is the piece selected appropriate for an audience of children who are within a range of kindergarten through sixth grade? (It is acceptable for a piece to be appropriate for a limited age group within the prescribed range.)",
            "Language - Is the diction clear and precise?  Is sufficient vocal emphasis used?  Are sense and image words emphasized?  Is the language used to its fullest potential? ",
            "Age Level / Interpretation - Is the material presented in such a way as to appeal to children? ",
            "Overall Effect",
            "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
            "Total Score: Maximum 32, Superior 32-28, Good 27-24          (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
            "Rank in Room ", "Over Time"]
drama = ["Speaker (student's first and last name)", "Code: ", "Selection (name of the piece)",
         "Judge's name (first and last)", "Round (number)", "Room (letter)",
         "Vocal Projection and Articulation - Does the voice both suit and enhance the material? Is volume used effectively? ",
         "Tempo and Phrasing - Does the pacing augment the mood and effect? Does the work flow with clarity? Does the phrasing establish the writer's thought pattern clearly?   Are dramatic pauses used effectively?",
         "Body and Facial Expression - Do the gestures enhance the interpretation of the material? Are the gestures suggestive and not sustained to the point of acting? Is the movement confined to the podium area?",
         "Eye Contact and Focus - Is a focus established? If two or more characters are present, is there an effective shift of focus? Is character focus effectively maintained throughout the interpretation?",
         "Sustaining of Mood and Character - Is there sufficient tension to communicate mood and conflict? Does the speaker utilize the language of the piece effectively?",
         "Characterization - Does the personality of the character(s) come through? If there are multiple characters, is each distinct?",
         "Conflict - Is there sufficient conflict for dramatic effect? Is the conflict clear? Does the conflict build?",
         "Total Effect - Is the selection appropriate to the category? Is the reader able to handle the maturity of the selection? Is the speaker's interpretation believable? Is the selection presented as an interpretation rather than an acting performance? ",
         "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
         "Total Score: Maximum 32, Superior 28-32, Good 24-27 *                                                                                                          (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
         "Rank in Room ", "Over Time"]
humor = ["Speaker (student's first and last name)", "Code", "Selection (name of the piece)",
         "Judge's Name (first and last)", "Round (number)", "Room (letter)",
         "Material Selection - Is the piece selected appropriate for humorous interpretation?  Are there sufficient light elements to the literature to make it humorous?",
         "Characterization/Narration - Is the characterization(s) and/or narration effective and humorous?  If multiple characters are evident, are they distinct?",
         "Tempo - Did the pacing, timing, and phrasing suggest the meaning and humor of the piece?",
         "Vocal Production - Did the voices in the characterization(s) and narration suit and enhance the humorous effect?  Did the speech flow with clarity?",
         "Delivery - Were the facial expressions (to the extent they could be seen) humorous?  Did the gestures enhance the humorous quality of the selection?  Is the interpretation original rather than imitative? ",
         "Eye Contact / Focus - Did the interpreter make effective use of focus?  Was effective visual contact maintained with the audience during the narrative elements of the selection? ",
         "Total Effect - Was the interpreter's presentation humorous?",
         "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
         "Total Score: Maximum 28, Superior 25-28, Good 21-24 (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
         "Rank in Room ", "Over Time"]
poetry = ["Speaker (student's first and last name)", "Code", "Selection (name of the piece)",
          "Judge's Name (first and last)", "Round (number)", "Room (letter)",
          "Projection of Mood and/or Character - Is the mood established and maintained? If characters are present, is each distinct?",
          "Body and Facial Expression - Do the facial expressions (to the extent they can be seen) contribute to the effective communication of the poem? Are the gestures suggestive and not sustained?",
          "Vocal Quality and Expression - Do voice inflections communicate the meaning of the poem? Is there suitable variation in the volume? Are tone and pitch suitable to the selection?",
          "Phrasing and Pacing - Does the rate suggest the mood and effect? Does the phrasing and rhythm establish the proper thought pattern of the poem? Do the rate and phrasing create a naturalness in the rhythm of the poem? Does the poem flow with clarity?",
          "Eye Contact/Focus - Does the speaker maintain effective visual contact with the audience? If focus was used, was it clearly and effectively maintained?",
          "Language - Is the diction clear and precise? Is sufficient vocal emphasis used? Are sense and image words emphasized? Is the language verbalized effectively?",
          "Total Effect - Did the interpretation enhance the poem's meaning? Does the speaker avoid being excessively dramatic? ",
          "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
          "Total Score: Maximum 28, Superior 25-28, Good 21-24 (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
          "Rank in Room ", "Over Time"]
prose = ["Speaker (student's first and last name)", "Code", "Selection (name of the piece)",
         "Judge's Name (first and last)", "Round (number)", "Room (letter)",
         "Vocal Projection and Articulation - Does the voice both suit and enhance the material? Does the speech flow with clarity? Is the volume used effectively?",
         "Tempo and Phrasing - Does the pacing suggest the mood and meaning? Do the phrasing and rhythm establish the proper thought pattern of the selection? If dialogue is present, is there a smooth transition between the narration and the dialogue?",
         "Body and Facial Expression - Do the gestures and facial expressions (to the extent they can be seen) enhance the interpretation of the material? ",
         "Eye Contact - Does the speaker maintain effective visual contact with the audience? Is the visual contact appropriate to the narrative point of view of the material? If characterization is present, is effective focus used?",
         "Narration/Characterization - Does the personality of the speaker/narrator and/or character(s) come through? If there is more than one character, is each distinct?",
         "Total Effect - Does the speaker communicate the understanding of the author's images and narrative style? Is the selection appropriate to the category? Is narration a significant part of the selection? Was the mood of the selection sustained throughout the interpretation? Does the speaker avoid being excessively dramatic? ",
         "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
         "Total Score: Maximum 24, Superior 21-24, Good 18-20 (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
         "Rank in Room ", "Over Time"]
ensemble = ["Code", "Selection (name of the piece)", "Judge's name (first and last)", "Room (letter)",
            "1. Material and introduction - Does the introduction appropriately set up the material? - Is the material appropriate for acting? - Is the material well cut? ",
            "2. Characterization - Are the characterizations appropriate to the selection? - Is the character of each reader clearly delineated? ",
            "3. Vocal production and articulation - Did the voices enhance the material? - Given the size of the room, did the voices enhance the characterization? - Was the energy level high? ",
            "4. Pacing - Was there a sense of movement in the scene? - Did the scene build to a climax? - Were the pauses effectively handled?",
            "5. Character interaction - Was the relationship between the characters clear? - Was the interaction between the characters believable? ",
            "6. Movement and staging - Did the staging create sufficient visual interest? - Did the movement enhance the text? - Did the movement appear natural to the characters? - Were entrances and exits handled effectively? ",
            "7. Total effect - Were the characters believable in the scene? - Did the scene have an impact on the audience?",
            "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
            "Total Score: Maximum 28, Superior 25-28, Good 21-24 (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
            "Rank in Room ", "Over Time"]
readers = ["Code", "Selection (name of the piece)", "Judge's Name (first and last)", "Room (letter)",
           "1. Material and introduction - does the introduction appropriately set up the material and establish rapport with the audience? - is the material effectively adapted for Readers Theater? - is the interpretation original? ",
           "2. Characterization - are the characterizations appropriate to the selection? - is the character of each reader clearly delineated? ",
           "3. Vocal production and articulation - did the voices enhance the material? - given the size of the room, did the voices project clear images? - did the text flow with clarity? - was the energy level high? ",
           "4. Nonverbal projection - did the gestures enhance the text? - were the gestures appropriate to the characterization? ",
           "5. Focus - did each performer (with the exception of the narrator) maintain appropriate offstage focus? - could one easily understand to whom each spoken line was addressed? ",
           "6. Staging - did the staging create sufficient visual interest? -- if movement was used, was it suggestive, rather than sustained to the point of acting? - were entrances and exits handled effectively? ",
           "7. Total effect - were effective Readers Theater techniques used to enhance the text (literature)? - was a unified impression created? -- did the performers make occasional reference to the text? - did the production communicate the message of the text? ",
           "Comments: Please provide a MINIMUM of TWO comments. At least ONE of them should contain specific, actionable feedback that the performer could use to improve their performance.                                                   (This is especially important if you give two performances the same number of points, to be transparent with our performers about why they received their particular rank.)",
           "Total Score: Maximum 28, Superior 25-28, Good 21-24 (Make sure to put this number, along with the performer's rank, on the paper summary sheet.)",
           "Rank in Room ", "Over Time"]

# compiling the rubric into a dictionary with key = full name, value = (rubric list, short name)

link = {"Extemporaneous Speaking": (extemp, "Extemp"), "Persuasive Ballot": (persua, "Persuasive"),
        "Original Works": (original, "OW"), "Informative Speaking": (inform, "Informative"),
        "Children_s Literature": (children, "Childrens"), "Dramatic Interpretation": (drama, "Drama"),
        "Humorous Interpretation": (humor, "Humor"), "Serious Poetry": (poetry, "Poetry"),
        "Serious Prose": (prose, "Prose"), "Ensemble Acting": (ensemble, "EA"), "Readers Theater": (readers, "RT")}
#If the user has made a rubric.txt, they can use their custom rubric instead
try:
    rubricData = open("rubric.txt", "r").readlines()
    copy_link = {}
    for item in rubricData:
        if (item != "\n"):
            l = item.split("~~")
            #Read the data from the file
            title = l[0].strip().rstrip("\"").lstrip("\"")
            shorttitle = l[1].strip().rstrip("\"").lstrip("\"")
            rubricList = eval(l[2])
            copy_link[title] = (rubricList, shorttitle)
    link = copy_link.copy()
except:
    pass
#Function that gives a number from 0 to 1 to judge how similar two strings are
def evaluateSimilarity(stringOne, stringTwo):
    stringOne = stringOne.lower()
    stringTwo = stringTwo.lower()
    return SequenceMatcher(None, stringOne, stringTwo).ratio()

#Uses the spreadsheet filename and finds the most probable category it corresponds to (using the evaluateSimilarity function)
def getCategory(filename):
    order = []
    #iterating through all the categories
    for item in link:
        if (item in filename):
            return link[item][1]
        order.append((evaluateSimilarity(item, filename), link[item][1]))
    #sorting in increasing similarity number
    order = sorted(order)
    #returns the value with the largest score
    return order[-1][1]

#Given the name of the tournament folder and the school code, returns the most probable school name
def getSchool(foldername, code):
    code = int(code)
    #Given the name of the folder, determine if it is qualifying tournament 1,2 or 3
    order = [(evaluateSimilarity(foldername, "QT1"), 0), (evaluateSimilarity(foldername, "QT2"), 1),
             (evaluateSimilarity(foldername, "QT3"), 2)]
    order = sorted(order)
    index = order[2][1]
    #Return the first school that matches the same code AND qualifying tournament
    for school in schoolIDs:
        if (schoolIDs[school][index] == code):
            return school
    #Return the first school that just matches the code (just in case the foldername prediction is incorrect).
    for school in schoolIDs:
        if (code in schoolIDs[school]):
            return school
    return ""

#Given the rubric and the question in the form, find what rubric column the question corresponds too
def getIndex(rubric, formtitle):
    order = []
    for i in range(len(rubric)):
        #Rubric item "Over Time" causes some issues
        if evaluateSimilarity(rubric[i], "Over Time") > .7:
            order.append((evaluateSimilarity("Was this student's performance over time (i.e., longer than 11 minutes)?", formtitle), i))
        else:
            order.append((evaluateSimilarity(rubric[i], formtitle), i))
    order = sorted(order)
    #find the best index position
    return order[-1]

#Take the ID/Code spreadsheet and create a dictionary with the IDs for each school (key = school, value = list of IDS)
def processIDs(filename):
    schoolIDs.clear()
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_index(0)
    # format = Name, Abrev, Round 1, Round 2, Round 3, Round...N
    rows = 0
    columns = 0
    #Finding the number of rows and columns in the spreadsheet
    while (columns != worksheet.ncols and worksheet.cell_value(0, columns) != ""):
        columns += 1
    columns -= 1

    while (rows != worksheet.nrows and worksheet.cell_value(rows, 0) != ""):
        rows += 1
    rows -= 1

    #Reading through the spreadsheet and grabbing all the data
    for i in range(1, rows + 1):
        info = []
        for j in range(2, columns + 1):
            #Some values have an "N". We ignore them by putting a code of -1
            if ("N" in str(worksheet.cell_value(i, j))):
                info.append(-1)
                continue
            info.append(int(worksheet.cell_value(i, j)))
        schoolIDs[worksheet.cell_value(i, 0)] = info
errors = open("errors.txt", "w+")
errors.close()
#Go through all the spreadsheets in the folder and add the student data into a dictionary with all the schools
#Format of dictionary: Key = school name, value = list of student data (judges comments for that student's performance"
def processDataFolder(foldername):
    categories = {}
    errors = open("errors.txt", "a")
    for file in os.listdir(foldername):
        trimmed = file.strip("Copy of ").strip(".xlsx")
        filename = os.path.join(foldername, file)
        #Ignore all files that are not spreadsheets
        if (not filename.endswith(".xlsx")):
            continue
        workbook = xlrd.open_workbook(filename)
        worksheet = workbook.sheet_by_index(0)
        rubric = []
        #Read the rubric for each sheet
        for i in range(worksheet.ncols):
            rubric.append(normalize('NFKD', worksheet.cell_value(0, i)))

        for i in range(1, worksheet.nrows):
            person = []
            #Read the student data for each row in the spreadsheet
            for j in range(worksheet.ncols):
                person.append((normalize('NFKD', str(worksheet.cell_value(i, j))), rubric[j]))

            #Add that student data to the correct school in the dictionary "schoolStudents"
            for j in range(len(person)):
                if (evaluateSimilarity(rubric[j], "code") > .75):
                    try:
                        school = getSchool(foldername, float(person[j][0]))
                        if (school in schoolStudents):
                            schoolStudents[school].append((trimmed, person))
                        else:
                            schoolStudents[school] = [(trimmed, person)]
                    except:
                        #If there is an error, inform the user
                        for item in person:
                            errors.write(item[0])
                            errors.write(" , ")
                        errors.write("\n")
                        pass

        categories[trimmed] = rubric
    errors.close()
#Take the dictionary "schoolStudents" and create spreadsheets for each school

def createSpreadsheets():
    statusInfo.set("Status: Compiling Information")
    root.update()
    files = []
    timenumbers = [15]
    seen = 0
    for school in schoolIDs:

        if (school not in schoolStudents):
            seen += 1
            continue
        #Creating a new XLSX file for each school
        start = time.time()
        prev = start
        workbook = xlsxwriter.Workbook(school + '.xlsx')
        files.append(school + '.xlsx')

        prediction = round((sum(timenumbers) / len(timenumbers)) * (len(schoolIDs) - seen))
        for item in link:
            currentCategory = link[item][1]
            row = 1
            #Creating a new sheet for each category
            worksheet = workbook.add_worksheet(currentCategory)
            for i in range(len(link[item][0])):
                worksheet.write(0, i, link[item][0][i])
            #Write all the student data to the sheet
            for student in schoolStudents[school]:
                if (getCategory(student[0]) == currentCategory):
                    data = []
                    done = []
                    for info in student[1]:
                        #Finding the best index to place the student data in
                        index = getIndex(link[item][0], info[1])
                        if (index == -1):
                            continue
                        data.append((index[0], index[1], info[0]))
                        #Check if 1 second has passed
                        if (time.time() - prev > 1):
                            prev += 1
                            prediction -= 1
                            #Update time
                            timeInfo.set("Estimated Time: " + str(prediction) + " seconds")
                            root.update()
                    data = sorted(data)[::-1]
                    for best in data:
                        if best[1] not in done:
                            worksheet.write(row, best[1], best[2])
                            done.append(best[1])
                    row += 1
                #Check if one second has passed
                if (time.time() - prev > 1):
                    #Update time
                    prev += 1
                    prediction -= 1
                    timeInfo.set("Estimated Time: " + str(prediction) + " seconds")
                    root.update()

        seen += 1
        end = time.time()
        timenumbers.append(end-start)
        #Use the RRAM technique to calculate the average time estimate to run one spreadsheet and multiply that by the number of sheets remaining
        prediction = round((sum(timenumbers) / len(timenumbers)) * (len(schoolIDs) - seen))
        pb['value'] = round((seen / len(schoolIDs)) * 100)
        timeInfo.set("Estimated Time: " + str(prediction) + " seconds")
        root.update()
        workbook.close()
    #Zip all the created files into output.zip
    with ZipFile('output.zip', 'w') as zipFile:
        for item in files:
            zipFile.write(item)
    #Delete all the files from directory (only output.zip remains)
    for item in files:
        os.remove(item)
    pb['value'] = 0
    statusInfo.set("Status: Finished Program")
    timeInfo.set("Estimated Time: ")
    root.update()




#Asks user for tournament folder and runs process data folder on it
def open_folder():
    if (len(schoolIDs) == 0):
        statusInfo.set("Error: Select Code Sheet First")
        timeInfo.set("Estimated Time:")
        return
    statusInfo.set("Status: Processing Qualifying Tournament Folder")
    timeInfo.set("Estimated Time: 4 seconds")
    root.update()
    folder = filedialog.askdirectory()
    processDataFolder(folder)
    try:
        processDataFolder(folder)
    except:
        statusInfo.set("Error: Select Folder")
        return
    timeInfo.set("Estimated Time: 0 seconds")
    statusInfo.set("Status: Upload Qualifying Tournament Folder (again) or Press Create Sheet button")

#Asks for the code sheet and processes the the IDs in the code sheet
def open_file():
    statusInfo.set("Status: Processing Code Sheet")
    timeInfo.set("Estimated Time: 2 seconds")
    root.update()
    file = filedialog.askopenfilename()
    if (".xlsx" not in file):
        statusInfo.set("Error: Must be a .xlsx File")
        timeInfo.set("Estimated Time: ")
        root.update()
    else:
        try:
            processIDs(file)
        except:
            statusInfo.set("Error: Must be the Code Sheet")
            timeInfo.set("Estimated Time: ")
            root.update()
            return
        timeInfo.set("Estimated Time: 0 seconds")
        statusInfo.set("Status: Upload Qualifying Tournament Folder")


#Create & Configure root
root = Tk()
root.title('Invenire Product')

width = root.winfo_screenwidth()//2
height = root.winfo_screenheight()//2
root.geometry(str(width) + "x" + str(height))
Grid.rowconfigure(root, 0, weight=1)
Grid.columnconfigure(root, 0, weight=1)

#Set up the moveable frame
frame=Frame(root, highlightbackground="black", highlightthickness=5)
frame.grid(row=0, column=0, sticky=N+S+E+W)
frame.configure(background='white')


#Create a 6x9 window of grid items that will stretch according to resolution
for row_index in range(6):
    Grid.rowconfigure(frame, row_index, weight=1)
    for col_index in range(9):
        Grid.columnconfigure(frame, col_index, weight=1)
        text = Label(frame, text=" ")
        text.grid(row=row_index, column=col_index)
        text.config(bg="white")

#Set up the text elements of the GUI
title = Label(frame, text="Invenire Application", font = ("Helvetica 18 bold"))
title.grid(row=0, column=0, columnspan=9)
title.config(bg="white")
statusInfo = StringVar()
statusInfo.set("Status: Upload Code Sheet")
status = Label(frame,textvariable = statusInfo, font = ("Helvetica 13 bold"),  wraplength=600)
status.grid(row=1, column=0, columnspan = 9)
status.config(bg="white")
timeInfo = StringVar()
timeInfo.set("Estimated Time:")
etime = Label(frame,textvariable = timeInfo, font = ("Helvetica 13 bold"))
etime.grid(row=2, column=0, columnspan = 9)
etime.config(bg="white")
pb = ttk.Progressbar(frame, orient='horizontal',mode='determinate',length=width//2)
pb['value'] = 0
pb.grid( column=0, row=3, columnspan=9)

#Set up the buttons in the GUI
school_codes = Button(frame, text='Select School Codes', font = ("Helvetica 10 bold"),  wraplength=150,command=open_file, pady=40, highlightthickness=2, highlightbackground="black")
school_codes.grid(column=1, row=4, sticky='')

tournament_folder = Button(frame, text='Select Qualifying Tournament Folder',  wraplength=200, font = ("Helvetica 10 bold"),command=open_folder, pady=40, highlightthickness=2, highlightbackground="black")
tournament_folder.grid(column=0, row=4, sticky='', columnspan=9)

spreadsheets = Button(frame, text='Create Spreadsheets', font = ("Helvetica 10 bold"),  wraplength=150,command=createSpreadsheets, pady=40, highlightthickness=2, highlightbackground="black")
spreadsheets.grid(column=7, row=4, sticky='')

#Open the GUI and constantly update it
root.mainloop()

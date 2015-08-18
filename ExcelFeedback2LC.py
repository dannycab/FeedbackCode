#! /usr/bin/env python

import sys
import os.path
import pandas as pd


def Parse2DF(filename):
    """Takes an input file and reads it into a Pandas dataframe. 
    If the file doesn't exist, returns an error."""

    if os.path.isfile(filename):

        DF = pd.read_excel(filename)

        return DF

    else:

        raise NameError(filename)


def Break2Strs(RowFromDF):
    """Takes a row from the dataframe it is passed and formats it into a string
    that LON-CAPA can parse using its group messaging feature. Includes testing for
    prior week's feedback scores."""

    # Assumes the number of columns in a single week (not historical feedback)
    # is 12; change if number of columns changes
    SingleWeekCol = 12

    DisclaimerString = "<i>To see this feedback formatted properly, log in to LON-CAPA.</i><br /><br />"

    AddressString = "\"" + RowFromDF['MSUNet_ID'] + ":" +\
        RowFromDF['domain'] + ":"

    AddressString += DisclaimerString

    FeedbackString = "<b>Student Name:</b> " + RowFromDF['Student_Name'] + "<br />"\
        " <b>Instructor:</b>" + RowFromDF['Instructor'] + "<br /><br />"\
        " <b>Feedback:</b><br /> " + RowFromDF['Feedback'] + "<br /><br/>"

    ScoreString = "<b>Weekly Group Work Score (out of 4)</b>: " + str(RowFromDF["Weekly_Group_Work_Score"].iat[0]) +\
    	"<br/><i>Highest score below counts 50%, second highest counts 33.3%, lowest counts 16.7%</i><br/>" +\
        "<ul><li>Group Understanding (out of 4): " +\
        str(RowFromDF["Group_Understanding_Score"].iat[0]) + "</li><li>Group Focus (out of 4): " +\
        str(RowFromDF["Group_Focus_Score"].iat[0]) + "</li><li>Individual Understanding (out of 4): " +\
        str(RowFromDF["Individual_Understanding_Score"].iat[0]) + "</li></ul>"

    # Tests if there's other feedback in the Excel file using the known length of the feedback file
    # and parses it using ProcessOldFeedback to generate historical description of the feedback
    # Quick and dirty and could probably be streamlined so it does not pass
    # the whole row from the dataframe

    if(RowFromDF.shape[1] > SingleWeekCol):

        # Checks how many extra columns there are
        xCols = RowFromDF.shape[1] - SingleWeekCol
        # Processes feedback based on those extra cols (ADD TEST FOR
        # DIVSIBILITY)
        OldFeedbackString = ProcessOldFeedback(RowFromDF, xCols)
        return AddressString + FeedbackString + ScoreString + OldFeedbackString

    else:

        return AddressString + FeedbackString + ScoreString


def ProcessOldFeedback(RowFromDF, xCols):
    """Processes the old feedback for additional weeks. After testing for the number of
    past weeks, it passes the Row and the week number to ProcessPriorWeek for string
    construction."""

    OldFeedbackString = ""  # blank string

    FeedbackScoreCategories = 4  # Assumes 3 score categories + overall score

    # Processes for the number of weeks of feedback (ADD DIVISIBILITY TEST TO
    # RAISE ERROR))
    for i in range(0, xCols / FeedbackScoreCategories):

        WeekString = ProcessPriorWeek(RowFromDF, i)
        OldFeedbackString += WeekString

    return OldFeedbackString


def ProcessPriorWeek(RowFromDF, i):
    """Processes a single set of feedback scores for a prior week
    Only includes the numeric scores and assumes a particular structure to the Excel 
    file headers where each col is numbered so that the feedback can be traced to 
    numbers of weeks ago."""

    gusCat = "Group_Understanding_Score_Past_" + str(i + 1)
    gfsCat = "Group_Focus_Score_Past_" + str(i + 1)
    iusCat = "Individual_Understanding_Score_Past_" + str(i + 1)
    weekly = "Weekly_Group_Work_Score_Past_" + str(i + 1)

    WeekString = "<b>"+str(i+1)+" Week Ago, Group Score:</b> "+str(RowFromDF[weekly].iat[0])+" <br/>"+\
    	"<ul><li>Group Understanding (" + str(i + 1) + " week ago): " + str(RowFromDF[gusCat].iat[0]) + "</li>"\
        "<li>Group Focus (" + str(i + 1) + " week ago): " + str(RowFromDF[gfsCat].iat[0]) + "</li>"\
        "<li>Individual Understanding (" + str(i + 1) + " week ago): " + str(RowFromDF[iusCat].iat[0]) + "</li></ul>"

    return WeekString


def DF2LCfile(DataFrame, OutPutFile):
    """Takes the dataframe that it is passed and converts it to a csv file
    (OutPutFile) that can be uploaded to LON-CAPA to make use of its group
    messaging feature."""

    for i in range(0, DataFrame.shape[0]):

        # Appends to the open file, so file will continue to grow if the code
        # is run multiple times
        with open(OutPutFile, "a") as LCFile:

            Break2Strs(DataFrame[
                       i:i + 1]).to_csv(LCFile, sep='\t', encoding='utf-8', header=False, index=False)

# Change to filename and path of file as needed.
File2Parse = "F2015_Spring_PHY183_Sec3_feedback_WeekZ_witholdscores.xlsx"

UploadFile = "upload.txt"

ExcelFeedback = Parse2DF(File2Parse)
DF2LCfile(ExcelFeedback, UploadFile)

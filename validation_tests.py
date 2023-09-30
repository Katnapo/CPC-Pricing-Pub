import unittest
import constants
import pytest as pytest
from constants import Constants
from classes import WorkbookManager, DataFormatter

def generateSampleIndexes(listLen):

    import random
    sampleIndexes = random.sample(range(0, listLen), listLen)
    return sampleIndexes
def sampleDataGenerator(sampleIndexes, outputDict):

    keys = list(outputDict.keys())
    refSamplesDict = {}

    for x in range(len(sampleIndexes)):

        identifier = keys[sampleIndexes[x]]
        refSamplesDict[identifier] = outputDict[identifier]

    return refSamplesDict

class TestPriceCalc():

    def setUpInputOutput(self, callback_function):

        # Loads the output workbook chosen by the user, then the input workbook which was given.
        # Note any workbook loading must be put down the input workbook function as the output workbook function isn't
        # picked up by the getInputData function

        self.workbookManager = WorkbookManager()
        callback_function(20, "Using general function to load output workbook given: " + Constants.outputBookName)
        self.workbookManager.loadInputWorkbook(Constants.outputBookName)
        self.formatter = DataFormatter()

        callback_function(20, "Getting and formatting output workbook data for: " + Constants.outputBookName)
        outputDict = self.workbookManager.getInputData(Constants.outputColumnNames, Constants.outputGeneralIdentifier, False)
        outputDict = self.formatter.removeNoneRows(outputDict)

        callback_function(10, "Generating sample sizes and indexes for testing")
        import math
        refSamplesLen = math.ceil(len(outputDict) * Constants.exchangeTestCoverage)

        if refSamplesLen > len(outputDict):
            refSamplesLen = len(outputDict)

        indexes = generateSampleIndexes(refSamplesLen)
        self.outputDict = outputDict
        self.outputSamples = sampleDataGenerator(indexes, outputDict)
        self.workbookManager.closeInputWorkbook()

        callback_function(20, "Using general function to load input workbook given: " + Constants.originalBookName)
        self.workbookManager.loadInputWorkbook(Constants.originalBookLocation)
        callback_function(20, "Getting and formatting input workbook data for: " + Constants.originalBookName)
        regularInputData = self.workbookManager.getInputData(Constants.inputColumnIdentifiers,
                                                             Constants.inputIdentifierColName, True, True)
        regularInputData = self.formatter.removeNoneRows(regularInputData)
        regularInputData = self.formatter.cellFormatterDict(regularInputData)
        self.regularInputData = regularInputData

        callback_function(10, "Generating samples for input workboook and finalizing setup")
        self.inputSamples = sampleDataGenerator(indexes, regularInputData)

    def cleanSetup(self):

        self.workbookManager = None
        self.formatter = None
        self.outputSamples = None
        self.inputSamples = None
        self.outputDict = None
        self.regularInputData = None
    def validateRandomConversionCalcs(self, callback_function):


        self.setUpInputOutput(callback_function)
        callback_function(0, "KILL", True)

        resultArray = []
        increment = len(self.inputSamples) / 200

        for sampleIdentifier in self.inputSamples:
            inputPrice = self.inputSamples[sampleIdentifier][Constants.inputPriceColName]
            callback_function(increment, "Matching input and output price for " + sampleIdentifier + "")

            for country in Constants.countryCodeExchange:
                newOutputPrice = self.formatter.convertCurrency(inputPrice, country[0])
                oldOutputPrice = self.outputSamples[sampleIdentifier][Constants.customerNumberPrefix + country[0]]
                resultArray.append([sampleIdentifier, country[0], newOutputPrice, oldOutputPrice])

        errorList = []
        for result in resultArray:

            callback_function(increment, "Validating " + result[0] + " " + result[1] + " " + str(result[2]) + " " + str(result[3]))
            print(str(result[2]) + " " + str(result[3]) + " " + str(result[0]))
            if result[2] != result[3]:
                errorList.append([result[0], result[1]])

        if len(errorList) > 0:
            import csv

            with open("error.csv", "w", newline='') as f:
                writer = csv.writer(f)
                writer.writerows(errorList)

            return False

        callback_function(100, "Finished.", True)
        return True

    def checkInputAndOutputSizeMatch(self, callback_function):

        # Find output and input lengths, divide the output dict by 13 to get its appropriate starting length then check.
        self.cleanSetup()
        self.setUpInputOutput(callback_function)

        callback_function(0, "KILL", True)
        callback_function(50, "Checking if input and output sizes match")

        outputDictCount = 0

        for key in self.outputDict:
            for country in Constants.countryCodeExchange:
                if self.outputDict[key][Constants.customerNumberPrefix + country[0]] != None:
                    outputDictCount += 1

        outputDictLen = int(outputDictCount / len(Constants.countryCodeExchange))
        inputDictLen = len(self.regularInputData)

        callback_function(50, "Found lengths: " + str(outputDictLen) + " for output and " + str(inputDictLen)
                          + " for input")

        if outputDictLen != inputDictLen:
            return False

        else:
            return True







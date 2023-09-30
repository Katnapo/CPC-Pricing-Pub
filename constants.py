
class Constants(object):

    # Version constants
    version = "1.5.1"

    # Country and cost conversion data
    countryCodeExchange = [["UK", 1], ["USA", 1.22], ["UAE",4.48],
                           ["Saudi", 4.57], ["Japan", 182.21], ["Hong Kong", 9.55],
                           ["South Korea", 1650], ["Australia", 1.88], ["Canada", 1.65],
                           ["Europe", 1.15], ["Qatar", 4.44], ["ROW", 1.15], ["Mexico", 21.25]]
    countryCodeRef = {}

    customerNumberPrefix = "Shopify "

    #TODO: Phase out identifier Pos here
    outputColumnNames = ["Item Reference", "Customer No.", "Starting Date", "Ending Date", "Price Type", "Unit Price"]
    itemReferenceIdentifierPos = 0
    outputGeneralIdentifier = "Item Reference"
    shopifyCodeIdentifierPos = 1
    outputShopifyCodeName = "Customer No."
    priceIdentifierPos = 5
    outputPriceUsed = "Unit Price"


    possiblePriceTypes = ["Full Price", "Markdown Price", "Promotional Price"]
    priceType = "Full Price"

    # Date constants. Start date customization 1: use today's date. 2: use custom date. 3: use date in input sheet
    startDateCustomization = 1
    startDateConstants = ["Use Today's Date", "Use Custom Date", "Use Date in Input Sheet"]
    customEndDateBool = False
    customStartDate = "07/08/2023"
    customEndDate = "06/08/2023"

    # Workbook constants
    inputBookName = "input.xlsx"
    originalBookLocation = "C:/Demo/"
    originalBookName = "original.xlsx"
    outputBookName = "output.xlsx"
    outputBookLocation = "C:/Demo/"
    errorFileName = "error.txt"

    # Method 2 for iterating through input sheet
    inputColumnIdentifiers = ["Price", "Item ID"]

    noneCountTolerance = 10
    searchRadius = 1000

    outputBookSuffix = " export.xlsx"

    # Constants for multiple uses
    inputIdentifierColName = "Item ID"
    inputPriceColName = "Price"
    inputStartDateColName = "Date"

    # API Usage Constants
    BC_KEYS = "REDACTED"
    useAPI = False
    closeOff = False

    # Testing
    exchangeTestCoverage = 0.1

    randomConversionCalcHelp = """
     This test randomly selects a number of items from the output sheet,
     and compares the calculated price from the input sheet with the price in the output sheet.
     If the calculated price is different from the output price, the test fails.
     This test is designed to test the currency conversion function.
     The number of items selected is determined by the 'coverage' text box. """

    inputOutputSizeMatchHelp = """
    Takes the input and output sheets, and compares the number of rows in the output to the input. 
    If the number of rows in (output / amount of currency conversions ) is not equal to the number of rows in the input,
    the test fails. Please note the test may fail if duplicate SKUs are used in sheets. Leave this to your own
    discretion and manually check the sheets if you are unsure. """





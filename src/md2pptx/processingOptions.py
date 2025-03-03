"""

processingOptions

"""

# Note: Options stored with lower case keys
class ProcessingOptions:
    def __init__(self):
        self.defaultOptions = {}
        self.presentationOptions = {}
        self.currentOptions = {}
        self.dynamicallyChangedOptions = {}
        self.hideMetadataStyle = False

    def getDefaultOption(self, optionName):
        return self.defaultOptions[optionName.lower()]

    def setDefaultOption(self, optionName, value):
        self.defaultOptions[optionName.lower()] = value

    def getPresentationOption(self, optionName):
        return self.presentationOptions[optionName.lower()]

    def setPresentationOption(self, optionName, value):
        self.presentationOptions[optionName.lower()] = value

    def getCurrentOption(self, optionName):
        return self.currentOptions[optionName.lower()][-1]

    # Note: Can't pop to an empty stack. Will always have default available
    def popCurrentOption(self, optionName):
        key = optionName.lower()
        if len(self.currentOptions[key]) > 1:
            # Have a non-default value to use
            self.currentOptions[key].pop()

    def setCurrentOption(self, optionName, value):
        key = optionName.lower()

        if key in self.currentOptions:
            # Add new value to existing stack
            self.currentOptions[key].append(value)
        else:
            # Start a new stack for this option
            self.currentOptions[key] = [value]

    def setOptionValues(self, optionName, value):
        key = optionName.lower()
        if key not in self.defaultOptions:
            self.setDefaultOption(optionName, value)

        self.setPresentationOption(optionName, value)

        self.setCurrentOption(optionName, value)

    def setOptionValuesArray(self, optionArray):
        for keyValuePair in optionArray:
            self.setOptionValues(keyValuePair[0], keyValuePair[1])

    def dynamicallySetOption(self, optionName, optionValue, conversion):
        lowerName = optionName.lower()
        if optionValue == "default":
            self.setCurrentOption(lowerName, self.getDefaultOption(lowerName))

        elif optionValue == "pres":
            self.setCurrentOption(lowerName, self.getPresentationOption(lowerName))

        elif optionValue in ["pop", "prev"]:
            self.popCurrentOption(lowerName)

        elif conversion == "":
            self.setCurrentOption(lowerName, optionValue)

        elif conversion == "float":
            self.setCurrentOption(lowerName, float(optionValue))

        elif conversion == "sortednumericlist":
            self.setCurrentOption(lowerName, sortedNumericList(optionValue))

        elif conversion == "int":
            self.setCurrentOption(lowerName, int(optionValue))

        self.dynamicallyChangedOptions[lowerName] = True

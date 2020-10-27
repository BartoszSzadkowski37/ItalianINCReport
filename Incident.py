import re
class Incident():
    def __init__(self, snowNumber, shortDesc, desc, resolvedDate, resolInfo, location):
        self.snowNumber = snowNumber
        self.shortDesc = shortDesc
        self.desc = desc
        self.resolvedDate = resolvedDate
        self.resolInfo = resolInfo
        self.location = location
        self.REF = '' 
        self.hriNumber = ''
        self.italianResolution = ''
        self.englishResolution = ''
        self.unknownResolution = ''
        
    def getHRIandREF(self):
        # create regexes
        hriRegex = re.compile(r'INC\d\d\d\d\d\d\d\D')
        refRegex = re.compile(r'Ref:MSGHSM\d\d\d\d\d\d\d\d')
        # searching in short description field
        hriMatch = hriRegex.search(self.shortDesc)
        refMatch = refRegex.search(self.shortDesc)
        # if cannot find in short description look in the description
        if hriMatch == None:
            hriMatch = hriRegex.search(self.desc)
        # the same for REF
        if refMatch == None:
            refMatch = refRegex.search(self.desc)
        # if still cannot find fill the hriNumber and REF fields for current incident by blank string
        if hriMatch == None or refMatch == None:
            self.hriNumber = ''
            self.REF = ''
        # if numbers have been found fill current incident fields
        else:
            self.hriNumber = hriMatch.group()
            self.hriNumber = self.hriNumber[:-1]

            self.REF = refMatch.group()
# thing about solving resolution getting
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
        self.resolutionTemplate = ''
        
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
            self.REF = self.REF[4:len(self.REF)]

    def createTheTemplate(self):
        self.resolutionTemplate = '''

        Hello,
        please be informed that the incident ''' + self.hriNumber + ''' has been resolved.
        Resolution information: [RESOLUTION INFO].

        The ticket has been resolved and will be closed within 5 days.
        If you do not agree with the Solution or the issue persists, please send your feedback and the request/ticket will be Reopened.
        After 5 days you have to raise a new request/ticket.
  
        Best Regards 
        Bridge Hitachi Team
        ----------------------------------------
        Buongiorno,
        l'incidente nr. ''' + self.hriNumber + ''' è stato lavorato

        Informazioni sulla risoluzione: [RESOLUTION INFO].
 
        Il ticket è stato risolto e verrà chiuso entro 5 giorni.
        Se non si è d'accordo con la Soluzione o il problema persiste, si prega di inviare il proprio feedback e la richiesta/ticket verrà riaperto.
        Dopo 5 giorni si deve presentare una nuova richiesta/ticket.
 
        Distinti saluti
        Hitachi Bridge Team
  
        REF: ''' + self.REF 

# thing about solving resolution getting


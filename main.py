import GoogleAPI
import CSVreader
import EmailList
import invoiceCreator


def getUsedInvoiceNumbers():
    usedInvoiceNumbers = GoogleAPI.readOldInvoiceNumbers()
    return usedInvoiceNumbers


def readFarnellCsvFile():
    allOrdersInInvoice, POnumber = CSVreader.readFarnellCSVfile()
    InvoicesAlreadySentForThisOrder = GoogleAPI.checkPOnumbers(POnumber)
    return allOrdersInInvoice, InvoicesAlreadySentForThisOrder, POnumber


def getEmailAddresses(allOrdersInCSVfile):
    currentEmailAddresses = GoogleAPI.readCurrentEmailAddresses()
    AllEmailAdressesExist, emailAddresses = EmailList.checkEmailAddresses(allOrdersInCSVfile,currentEmailAddresses)
    return AllEmailAdressesExist, emailAddresses


def createInvoices(allOrdersInCSVfile,oldInvoiceNumbers):
    invoiceNumbers = []
    for order in allOrdersInCSVfile:
        invoiceNumber = invoiceCreator.createPDF(order,oldInvoiceNumbers)
        invoiceNumbers.append(invoiceNumber)
    return invoiceNumbers


def addInvoicesToSheet(invoiceNumbers,allOrdersInCSVfile):
    for index, order in enumerate(allOrdersInCSVfile):
        entry = GoogleAPI.createSingleEntryForSheet(order,invoiceNumbers[index])
        GoogleAPI.addInvoiceToSheet(entry)


def sendInvoices(emailAddresses,invoiceNumbers):
    for index, address in enumerate(emailAddresses):
        GoogleAPI.sendInvoice(address[0],address[1],invoiceNumbers[index])

    
def main():
    allOrdersInCSVfile, InvoicesAlreadySentForThisOrder, POnumber = readFarnellCsvFile()      # Läs csv-filen från Farnell
    if InvoicesAlreadySentForThisOrder:
        return 0
    else:
        AllEmailAdressesExist, emailAddresses = getEmailAddresses(allOrdersInCSVfile)   # Kolla upp eller lägg till mailadresser
        if AllEmailAdressesExist:
            oldInvoiceNumbers = getUsedInvoiceNumbers()                                 # Läs gamla fakturanummer i Google Sheet #AUTH
            invoiceNumbers = createInvoices(allOrdersInCSVfile,oldInvoiceNumbers)       # Skapa fakturor+fakturanummer 
            addInvoicesToSheet(invoiceNumbers,allOrdersInCSVfile)                       # Lägg till nya fakturor i Google Sheet #AUTH
            sendInvoices(emailAddresses,invoiceNumbers)                                 # Skicka mail med fakturor #AUTH
            GoogleAPI.addPOnumber(POnumber)
        else:
            print('Script ended, no invoices have been created nor sent.')


if __name__ == '__main__':
    main()
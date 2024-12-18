from reportlab.pdfgen import canvas 
from reportlab.pdfbase.ttfonts import TTFont 
from reportlab.pdfbase import pdfmetrics 
from reportlab.lib import colors
from reportlab.platypus.tables import Table
from reportlab.platypus.tables import TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from datetime import date
from dateutil.relativedelta import relativedelta


def createPDF(totalOrder,usedInvoiceNumbers):
    name = totalOrder[0]
    orders = totalOrder[1]
    orderTotal = totalOrder[2]
    orderDate = orders[0][0]
    index = str(totalOrder[3]+1)
    invoiceNumber = createInvoiceNumbers(index,usedInvoiceNumbers)                                                              # +1 because the first invoice should be nr.1 and not nr.0
    documentTitle = 'ETA Faktura' + ' ' + str(name[0]) + ' ' + str(name[1]) + ' ' + invoiceNumber
    fileName = documentTitle + '.pdf'
    indent = 40
    today = str(date.today())

    # Create PDF
    pdf = canvas.Canvas(fileName)
    pdf.setTitle(documentTitle)
    pdfmetrics.registerFont(TTFont('abc', 'Times New Roman.ttf')) 
  
    # Create title
    title = 'Faktura' 
    pdf.setFont('Helvetica-Bold', 36)
    pdf.setFillColorRGB(0.2,0.4,0.6)
    pdf.drawString(indent, 770, title)

    # Create subtitle
    subTitle = 'Farnellförmedling'
    pdf.setFillColor(colors.black) 
    pdf.setFont("Helvetica", 24) 
    pdf.drawString(indent, 740, subTitle)

    # ETA logo
    image = 'ETA_logga.jpg'
    pdf.drawInlineImage(image, 430, 740, width=130, height=60)

    # Address
    address = [
        'ETA Chalmers',
        'Rännvägen 4',
        '412 96 Göteborg'
    ]
    text = pdf.beginText(indent,700)
    text.setFont('Helvetica',12)
    for line in address:
        text.textLine(line)
    pdf.drawText(text)

    # Invoice date
    # Title
    text = pdf.beginText(indent,630)
    text.setFont('Helvetica',12)
    text.textLine('Fakturadatum:')
    pdf.drawText(text)
    # Value
    text = pdf.beginText(indent,615)
    text.setFont('Helvetica-Bold',12)
    text.textLine(today)
    pdf.drawText(text)

    # Due date
    # Title
    text = pdf.beginText(indent+95,630)
    text.setFont('Helvetica',12)
    text.textLine('Förfallodatum:')
    pdf.drawText(text)
    # Value
    text = pdf.beginText(indent+95,615)
    text.setFont('Helvetica-Bold',12)
    due_date = date.today() + relativedelta(months=+3)
    text.textLine(str(due_date))
    pdf.drawText(text)

    # Invoice number
    # Title
    text = pdf.beginText(indent+190,630)
    text.setFont('Helvetica',12)
    text.textLine('Fakturanummer:')
    pdf.drawText(text)
    # Value
    text = pdf.beginText(indent+190,615)
    text.setFont('Helvetica-Bold',12)
    text.textLine(invoiceNumber)
    pdf.drawText(text)

    # Recipient
    # Title
    text = pdf.beginText(indent+350,630)
    text.setFont('Helvetica',12)
    text.textLine('Mottagare:')
    pdf.drawText(text)
    # Value
    text = pdf.beginText(indent+350,615)
    text.setFont('Helvetica-Bold',12)
    text.textLine(str(name[0]) + ' ' + str(name[1]))
    pdf.drawText(text)

    # How to pay and bankgiro number
    text = pdf.beginText(indent,580)
    text.setFont('Helvetica-Bold',12)
    orderDate = orderDate[:4] + '-' + orderDate[5:7] + '-' + orderDate[8:10]
    textContent = 'Vänligen ange fakturanummer vid betalning. ETAs bankgironummer: 5930-5680'
    text.textLine(textContent)
    pdf.drawText(text)

    # Date when the order was placed by ETA
    text = pdf.beginText(indent,565)
    text.setFont('Helvetica',9)
    orderDate = orderDate[:4] + '-' + orderDate[5:7] + '-' + orderDate[8:10]
    textContent = 'Varorna beställdes av ETA från Farnell vid följande datum: ' + orderDate
    text.textLine(textContent)
    pdf.drawText(text)
    
    #pdf.showPage()
    verticalOffset = 18*(len(orders)+2)

    # Table of all items purchased
    if len(orders) <= 22:
            data = [['Produktlänk','Antal','Á-pris','Totalpris']]
            for order in orders:
                data = makeTableEntries(order,data)
            data.append(['','','',''])
            data.append(['Belopp att betala:','','',str('%.2f' % orderTotal + ' SEK').replace('.',',')])
            t = Table(data,colWidths=[170,100,130,100])
            t.setStyle(TableStyle([('FONTNAME',(0,0),(4,0),'Helvetica-Bold'),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('SIZE',(0,0),(-1,0),12),('LINEBELOW',(0,0),(-1,0),0.5,colors.black),('LINEBELOW',(0,-1),(-1,-1),0.5,colors.black),('SIZE',(0,1),(-1,-1),12)]))
            t.wrapOn(pdf, 0,0)
            verticalOffset = 18*len(orders)
            t.drawOn(pdf,indent,500-verticalOffset)
    else:
        data = [['Produktlänk','Antal','Á-pris','Totalpris']]
        for order in orders[:22]:
            data = makeTableEntries(order,data)
        t = Table(data,colWidths=[170,100,130,100])
        t.setStyle(TableStyle([('FONTNAME',(0,0),(4,0),'Helvetica-Bold'),('SIZE',(0,0),(-1,0),12),('LINEBELOW',(0,0),(-1,0),0.5,colors.black),('LINEBELOW',(0,-1),(-1,-1),0.5,colors.black),('SIZE',(0,1),(-1,-1),12)]))
        t.wrapOn(pdf, 0,0)
        t.drawOn(pdf,indent,115)
        addInformationAtBottomOfPage(pdf)
        orders = orders[22:]

        while len(orders) > 37:
            pdf.showPage()
            data = [['Produktlänk','Antal','Á-pris','Totalpris']]
            for order in orders[:37]:
                data = makeTableEntries(order,data)
            t = Table(data,colWidths=[170,100,130,100])
            t.setStyle(TableStyle([('FONTNAME',(0,0),(4,0),'Helvetica-Bold'),('SIZE',(0,0),(-1,0),12),('LINEBELOW',(0,0),(-1,0),0.5,colors.black),('LINEBELOW',(0,-1),(-1,-1),0.5,colors.black),('SIZE',(0,1),(-1,-1),12)]))
            t.wrapOn(pdf, 0,0)
            t.drawOn(pdf,indent,120)
            addInformationAtBottomOfPage(pdf)
            orders = orders[37:]

        pdf.showPage()
        data = [['Produktlänk','Antal','Á-pris','Totalpris']]
        for order in orders:
            data = makeTableEntries(order,data)
        data.append(['','','','']) 
        data.append(['Belopp att betala:','','',str('%.2f' % orderTotal + ' SEK').replace('.',',')])
        t = Table(data,colWidths=[170,100,130,100])
        t.setStyle(TableStyle([('FONTNAME',(0,0),(4,0),'Helvetica-Bold'),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('SIZE',(0,0),(-1,0),12),('LINEBELOW',(0,0),(-1,0),0.5,colors.black),('LINEBELOW',(0,-1),(-1,-1),0.5,colors.black),('SIZE',(0,1),(-1,-1),12)]))
        t.wrapOn(pdf, 0,0)
        verticalOffset = 18*len(orders)
        t.drawOn(pdf,indent,765-verticalOffset)

    # General ETA information at bottom of page
    addInformationAtBottomOfPage(pdf)

    # Create the document
    pdf.save()
    return invoiceNumber


def addInformationAtBottomOfPage(pdf):
    my_Style=ParagraphStyle('My Para style', fontName='Helvetica', fontSize=11, textColor=colors.blue)
    eta_website = Paragraph('<a href=https://eta.chalmers.se/> eta.chalmers.se </a>',my_Style)
    data = [
        ['ETA Chalmers', 'Kontaktuppgifter', 'Betalningsinformtion'],
        ['Rännvägen 4', 'Telefon: 031-20 78 60', 'Bank: SEB'],
        ['412 96 Göteborg', 'Mejl: eta@eta.chalmers.se', 'BIC: ESSESESSXXX'],
        ['Org.nr: 802413-7963', eta_website, 'IBAN: SE8150000000050261003836']
    ]
    t = Table(data,colWidths=[185,185,185])
    t.setStyle(TableStyle([('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('LINEABOVE',(0,0),(-1,0),0.5,colors.black),('SIZE',(0,0),(-1,-1),11),]))
    t.wrapOn(pdf,0,0)
    t.drawOn(pdf,18,20)

    # Bottom of page blue field
    t = Table([''],colWidths=[596]) 
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(0,0),colors.Color(red=0.2,green=0.4,blue=0.6))]))
    t.wrapOn(pdf,0,0)
    t.drawOn(pdf,0,0)    


def makeTableEntries(order,data):
    orderLink = '<a href=' + 'https://se.farnell.com/' + order[1] + '>' + order[1] + '</a>'
    my_Style=ParagraphStyle('My Para style', fontName='Helvetica', fontSize=12, textColor=colors.blue)  # Paragraph style for URL links via HTML
    orderCode =  Paragraph(orderLink, my_Style)
    orderQuantity = int(float(order[2]))
    orderUnitPrice = (str('%.2f' % float(order[3][3:])) + ' SEK').replace('.',',')
    orderTotalPrice = (str('%.2f' % float(order[4][3:])) + ' SEK').replace('.',',')
    tableEntry = [orderCode,orderQuantity,orderUnitPrice,orderTotalPrice]
    data.append(tableEntry)
    return data


def createInvoiceNumbers(index,usedInvoiceNumbers):
    index = int(float(index))
    today = str(date.today())
    invoiceDateNoDashes = today[:4] + today[5:7] + today [8:10]
    for oldInvoiceNumber in usedInvoiceNumbers:
        if len(oldInvoiceNumber) == 0:
            pass
        elif str(oldInvoiceNumber[0])[:8] == invoiceDateNoDashes:
            index += 1
        else:
            pass
    for _ in range(4-len(str(index))):
        index = '0' + str(index)
    invoiceNumber = invoiceDateNoDashes + index
    return invoiceNumber


def main():
    print('This is the invoice PDF creator script')


if __name__ == '__main__':
    main()
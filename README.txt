TableFactory is a very simple interface for creating report tables from data
sets you provide. It uses other projects for most of the heavy lifting:
ReportLab makes PDFS and xlwt makes spreadsheets. It's especially well
suited to adding reporting capabilities to your Pyramid, TurboGears, Pylons,
or Django projects.

# Motivation

I maintain a website that provides many custom reports to its users, and
most of them need to be available in several different output formats.
Almost all of those reports follow the same pattern:

1. Run a database query,
2. Reformat the data slightly as needed, and
3. Return a (usually) simple grid of a few columns from each of those rows.

TableFactory addresses #3. Some customers are content to view their data in
their web browser, while others want print-ready PDFs and still others want
to edit it in a spreadsheet program. I needed an easy-to-use wrapper around
the other reporting backends so that I could configure a report one time and
then publish it in any desired format. ReportLab is astoundingly powerful,
but these simple little tables only use a fraction of its power. The same is
true for xlwt: it can do many amazing things that I never need it to do.

This is where TableFactory comes in. It's not as flexible as either of those
projects, but it does all the tedious, repetitive, and fragile work of
getting the data ready for output and building the layout of the finished
tables.

# Example

Suppose I want to build a table listing the customer name and total amount
from some invoices in our database.  My company uses SQLAlchemy and I have
an "Invoice" class that maps to our invoice table. This fetches the first
ten rows from that table:

    invoices = session.query(Invoice).limit(10)

I'm finished with step #1 above. This is a simple report and I'll skip the
second step and go straight to generating the output.

First, I'll build a "row specification" object that contains information
about the columns in the report. Each "column specification" lists the name
of a column from an invoice table row and its human-readable name. The
customer wants the invoice amounts to stand out, so I'll make them bold:

    rowmaker = RowSpec(ColumnSpec('customer', 'Customer'),
                       ColumnSpec('invamt', 'Invoice Amount', bold=True))

TableFactory classes work on "table row" objects. The RowSpec instance I
just made can convert those SQLAlchemy results into TableRows:
		       
    lines = rowmaker.makeall(invoices)
    
Behind the scenes, it loops across all of the objects in "invoices" and
converts the "customer" and "invamt" columns into table cells. Next, I'll
create the table builder:

    pdfmaker = PDFTable('Invoice amounts by customer', headers=rowmaker)
    
This will give us a PDF titled "Invoice amounts by customer" with columns
titled "Customer" and "Invoice Amount" (pulled from the RowSpec I made a
couple of steps ago!). Finally, to assemble the PDF and write it to a file:
    
    open('invoicetable.pdf', 'wb').write(pdfmaker.render(lines))
    
Our PDFTable's "render" method accepts the TableRows I made earlier and
turns them into a PDF. That's it! I'm done and ready to go home for the day.

But wait! The customer's accounting department needs an Excel spreadsheet
they can import into their own database. Conveniently, I've already done all
the "hard" work and only need to make a new spreadsheet generator:

    sheet = SpreadsheetTable('Invoice amounts by customer', headers=rowmaker)
    open('invoicetable.xls', 'wb').write(sheet.render(lines))

And with that, it's time to go on a break.

# License

TableFactory is available under the permissive MIT License.

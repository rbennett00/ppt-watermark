# ppt-watermark
PowerPoint VBA add-in to generate unique watermarks for each recipient and export to PDF

When run, the user is prompted to enter the template text for the watermark - the default is "Prepared for {n} on {d}", where {n} will be replaced with each individual's name, and {d} is the current date.

Then the user is prompted to select an Excel workbook containing the list of names. They can either select a Table and column within the workbook or manually select a range that contains the names.

Then the user is prompted to select the output directory for the generated PDFs.

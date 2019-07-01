import tabula

# Read pdf into DataFrame
df = tabula.read_pdf("Sound-Platform-Attendants Complete List_March 2019-1.pdf", spreadsheet=True, multiple_tables=True, lattice=True, stream=False)
# , area=(83.655,47.52,286.605,738.54)

# Read remote pdf into DataFrame
# df2 = tabula.read_pdf("https://github.com/tabulapdf/tabula-java/raw/master/src/test/resources/technology/tabula/arabic.pdf")

# convert PDF into CSV
tabula.convert_into("Sound-Platform-Attendants Complete List_March 2019-1.pdf", "output.csv", output_format="csv")

# convert all PDFs in a directory
# tabula.convert_into_by_batch("input_directory", output_format='csv')

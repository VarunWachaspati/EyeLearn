import glob, csv, xlwt, os

csv_loc = "/home/varunwachaspati/LVPEI/csv/"
def main():
    os.chdir(csv_loc)
    wb = xlwt.Workbook()
    for filename in glob.glob(csv_loc+"/*.csv"):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)
        ws = wb.add_sheet(f_short_name)
        spamReader = csv.reader(open(filename, 'rb'))
        for rowx, row in enumerate(spamReader):
            for colx, value in enumerate(row):
                ws.write(rowx, colx, value)
    wb.save("/home/varunwachaspati/LVPEI/compiled.xls")

    if __name__ == '__main__':
        main()
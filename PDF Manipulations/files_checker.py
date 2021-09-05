import os
from pathlib import Path
import csv

root_path = "/mnt/c/Users/Yash/Desktop/Desktop/Aero Telcom Intern Work/Task 18(PDF manipulations)/PDF Test"
# print(root_path)
# print(os.getcwd())
contents = Path(root_path).glob("*")
# print(contents)
for content in contents:
    content = str(content)
    if(content.find(".")!=-1):
        continue
    else:
        pdfs_folder = content + "/CsvsPDFDownloaded"
        img_dir = content + "/Images/Images"

        contents1 = Path(content).glob("*")
        for c1 in contents1:
            c1 = str(c1)  
            if(c1.find('~')!=-1):
                continue
            if(c1.endswith("_PRODUCTS.xlsx")):
                excel_name = c1
            if(c1.endswith("ImageDetails.csv")):
                img_csv = c1
            if(c1.endswith("downloaded_pdf.csv")):
                pdf_csv = c1
        imgdetails={}
        pdfdetails={}
        with open(img_csv) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            lc=0
            for row in csv_reader:
                if(lc==0):
                    lc+=1
                    continue
                else:
                    if(row[0]):
                        imgdetails[row[0]]=row[1]
                lc+=1
            csv_file.close()
        with open(pdf_csv) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            lc=0
            for row in csv_reader:
                if(lc==0):
                    lc+=1
                    continue
                else:
                    if(row[0]):
                        pdfdetails[row[0]]=row[1]
                lc+=1
            csv_file.close()
        for k in imgdetails.keys():
            img_path  = img_dir+"/"+imgdetails[k]
            if(os.path.isfile(img_path)):
                continue
            else:
                print(f"Mfr. No. {k}, image name {imgdetails[k]} does not exist in Images folder\n")
        for k in pdfdetails.keys():
            pdf_path  = pdfs_folder+"/"+pdfdetails[k]
            if(os.path.isfile(pdf_path)):
                continue
            else:
                print(f"Mfr. No. {k}, PDF name {pdfdetails[k]} does not exist in Downloaded pdfs folder\n")
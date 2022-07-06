import glob
import win32com.client
import os
import logging

def log():
    logging.basicConfig(filename = 'log.log',level=logging.DEBUG, format='\n%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
    logging.error("Exception occurred", exc_info=True)

word = win32com.client.Dispatch("Word.Application")
word.visible = 0

path_input = r"C:\Users\akhil\Downloads\new"
pdfs_path = path_input

os.chdir(path_input)

print(pdfs_path)

#for i, doc in enumerate(glob.iglob(pdfs_path+"*.pdf")):
for pdf_file in os.listdir(pdfs_path):
    
    #print("inside loop")
    if pdf_file.endswith(".pdf"):
            try:
              filename = pdf_file.split('\\')[-1]
              in_file = os.path.abspath(pdf_file)
              print(in_file)
              wb = word.Documents.Open(in_file)
              log()
              out_file = os.path.abspath(reqs_path +filename[0:-4]+ ".docx".format(i))
              print("outfile\n",out_file)
              wb.SaveAs2(out_file, FileFormat=16) # file format for docx
              print("success...")
              wb.Close()
              word.Quit()
            except:
                 word.Quit()
                 print('Error: Please check PDF files for corruption and convert manually.')
                 log()  



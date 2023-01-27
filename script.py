import xltohtml
import argparse

parser = argparse.ArgumentParser()

parser.add_argument("-i","--input_file",help="Ex :- input file = ExcelFilename.xlsx")
parser.add_argument("-o","--output_file",help="Ex :- output file =output/ ")
parser.add_argument("-n","--name",help="Ex :- Prefix for filename = webpagePrefix")

args = parser.parse_args()

if args.input_file and args.output_file and args.name:
    xltohtml.input_file = args.input_file   
    xltohtml.output_file= args.output_file  
    xltohtml.name= args.name                              
    xltohtml.run()
    print("PAGES WERE CREATED!!")
    print("To check the pages open the given output folder.")
else:
  print("You can use the below to learn the usage")
  print("script.exe -i User:\FolderName\inout\Excel_filename.xlsx -o User:\FloderName\output\ -n webpagePrefix")

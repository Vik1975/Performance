import os, sys
import argparse		# Get arguments from the command line
from pathlib import Path	# handles path
import shutil		# needed to copy and unzip files
import json			# Needed to import and use JSON files
import csv			# work with the CSVs
from openpyxl import Workbook, load_workbook # work with the .xlsx

TEMPLATES_FOLDER = Path(sys.path[0]).joinpath('Templates')
TRACE_FOLDER = "~/.pnp/traces/"
DEBUG = False

#Command to run PNP (Should be "pnp" most of the time or "python pnp.par" for hcl)
#PNP_COMMAND = "py -3.10 C:\ProgramData\chocolatey\lib\fb-pnp-windows\bin\pnp.par"
PNP_COMMAND = "pnp"

# List of known hmds linking external and internal codenames
KNOWN_HMDS = {
	"eureka":"Stinson",
	"seacliff":"Arcata",
	"hollywood":"Miramar",
	"panther":"Ventura"
	}
# name of the CSV file that needs to be added to a summary.xlsx
CSV_FILENAME = "cpu_credit.csv"

# For storing all the uploaded trace names, to list at the end. 
UPLOADED_TRACES = []

def load_json_data(json_file):
	# Get values from JSON returns a dictionary
	
	with open(json_file, 'r') as metadata_file:
		dictionary = json.load(metadata_file)
	return dictionary


def find_file(directory, filename):
	# searches a directory for a filename and returns the path to the first instance of the file. 
	# Returns an absolute path if the file is found and False if it's not.
	
	files =  directory.rglob(filename)
	
	for file in files:
		if  file.name == filename:
			return file.absolute()
	return False

def find_template(hmd):
	# find the right summary_xlsx template for the given HMD and return the filename

	# convert the hmd names
	if hmd in KNOWN_HMDS:
		name = KNOWN_HMDS[hmd]
	
	for file in TEMPLATES_FOLDER.iterdir():
		if name in file.name:
			return file
	assert False, f'Could not find {hmd} in {TEMPLATES_FOLDER}'	


def parse_metadata(folder):
	# Gets the metadata from a folder. If the given folder is not a trace folder returns False.
	#Trace folders are named <HMD>_<Build_number>_<build_flavor>_<test_date>_<test_time>
	
	if not folder.is_dir():
		return False
	
	json_file = find_file(folder, "metadata.json")
	if json_file:
		run_info = {}
		metadata = load_json_data(json_file)
		if 'device_type' in metadata:
			run_info['device'] = metadata['device_type']
		if 'scenario' in metadata:
			run_info['scenario'] = metadata['scenario']
		if 'build_branch' in metadata:
			run_info['branch'] = metadata['build_branch'].split("-")[-1]
		if 'build_version' in metadata:
			run_info['build'] = metadata['build_version']
		
		# Check that the folder name matches the metadata, only trace folders should match this. 
		if f"{run_info['device']}_{run_info['build']}_{metadata['build_flavor']}" in folder.stem:
			#print(f'Got this metadata: {run_info}')
			return run_info
		else:
            #TODO: Make this error nicely and maybe ask if you want to use it anyway.
			return False
		
	else:
		#print(f"WARNING!! Could not find metadata.json in {folder}")
		return False

def find_or_make_folder(folder):
	#Takes a folder name and tries to make it, if it already exists it returns the existing folder.
	try:
		folder.mkdir(parents=True)
		return folder
	except FileExistsError:
		return folder

def make_run_folder(scenario_folder):
	#Takes a scenario_folder counts how many run folders are in it then makes a run_folder n+1
	#returns the run_folder and run_num
	run_num = 1
	for run in scenario_folder.iterdir():
		if run.is_dir() and run.stem.split(" ")[0] == "Run":
			run_num += 1
		elif run.stem.split(" ")[0] == "Run" and run.is_file() and run.suffix  == ".zip":
			run_num += 1
	
	run_folder_name = f"Run {run_num}"
	run_folder = scenario_folder.joinpath(run_folder_name)
	try:
		run_folder.mkdir()
		return run_folder, run_num
	except FileExistsError:
		assert False, f'FATAL ERROR!!! {run_folder} already exists for some reason!!'

def get_csv_data(input_csv):
	# reads the input CSV shoves it into an array
	csv_rows = []
	with open(input_csv) as csv_file:
		csv_reader = csv.reader(csv_file, delimiter=',')
		for row in csv_reader:	# iterates through all the rows
			csv_rows.append(row)
	return csv_rows

def find_summary_xlsx(folder, summary_name):
    # looks for a summary_name file and returns it. Returns false if none is found
    for file in folder.iterdir():
        if file.is_file() and file.name == summary_name:
            return file
    return False
	
def add_csv_to_summary_xlsx(summary_xlsx, csv_file, run_num):
	# adds copies the information from a csv file to the run_num worksheet in the summary_xlsx
	
	# load data from csv_file
	csv_rows = get_csv_data(csv_file)
	
	# load the summary_xlsx
	wb = load_workbook(summary_xlsx)
		
	# select correct worksheet
	sheet_name = f"Run#{run_num}"
	if sheet_name in wb.sheetnames:
		ws = wb[sheet_name]
		#print(f'Found Worksheet name: {sheet_name}')
	else:
		#TODO: Add a new sheet of the appropiate name. 
		#TODO: Add an extra data column in the "Summary" and "Summary for avg" sheets
		#TODO: May need to change the values for the average all runs column
		assert False, f'Could not find {sheet_name} in {summary_xlsx}'
	
	#print(f'Adding data to {sheet_name} in {summary_xlsx.name}')
	
	# add csv data to the worksheet on summary_xlsx iterating through rows and columns
	x=1
	for row in csv_rows:
		y=1
		for cell in row:
			# if the cell only contains numbers cast it as a number when importing so excel doesn't treat it as text.
			try:
				ws.cell(row=x,column=y,value=float(cell)).number_format
			except ValueError:
				# if it's a percentage (and therefor doesn't automatically convert above. 
				if not len(cell) == 0:
					if cell[-1] == '%':
						ws.cell(row=x,column=y,value=(float(cell[:-1])/100)).number_format = '0.00%'
					else:
						# add the CSV data as text if it's not already handled.
						ws.cell(row=x,column=y,value=cell)
			y+=1
		x+=1
	
	# https://openpyxl.readthedocs.io/en/stable/styles.html
	# https://stackoverflow.com/questions/8440284/setting-styles-in-openpyxl
	# https://stackoverflow.com/questions/12387212/openpyxl-setting-number-format
	
	
	# https://stackoverflow.com/questions/12976378/openpyxl-convert-csv-to-excel
	# https://www.blog.pythonlibrary.org/2021/09/25/converting-csv-to-excel-with-python/
	
	# save the workbook
	wb.save(summary_xlsx)

def delete_perfetto(folder):
	# deletes any .perfetto files in a folder
	files =  folder.rglob("*.perfetto")
	
	for file in files:
		print(f'Deleting {file}')
		file.unlink()
	return False
	
def recursive_unzip(zipfile):
	# takes a path object for a folder or a .zip file and unzipps the file and any .zip subfiles returns the location of the unzipped folder
	if DEBUG:
		print(f'unzip input file: {zipfile}')
	if zipfile.is_file() and zipfile.suffix  == ".zip":
		if DEBUG:
			print(f'This is a .zip : {zipfile}')
		#Make a folder with the same name as the zipfile
		unzip_folder = zipfile.parent.joinpath(Path(zipfile.stem))
		Path.mkdir(unzip_folder)
	
		# unzip the zip to the folder
		print(f'unzipping {zipfile.name} to {unzip_folder}')
		shutil.unpack_archive(zipfile, unzip_folder)
	
		#delete the original zipfile
		zipfile.unlink()
		
	elif zipfile.is_dir():
		if DEBUG:
			print(f'This is a dir : {zipfile}')
		unzip_folder = zipfile
	else:
		return False
	
	# iterate through the unzipped folder and unzip anything else found.
	for filename in unzip_folder.iterdir():
		if filename.suffix == ".zip" or filename.is_dir():
			recursive_unzip(filename)
	
	return unzip_folder

def zip_in_place(folder):
	# zips a given folder in place removing the original folder. 
	# Makes a folder.zip in the same location as folder was. 
	print(f'Zipping: {folder}')
	shutil.make_archive(folder,'zip',folder)
	shutil.rmtree(folder)
	
def is_trace_folder(folder):
	# Returns true if the given folder is a trace folder. 
	# Determines that this is a trace folder if there is a metadata.json file and a artifacts folder in it. 
	
	if folder.is_dir():
		metadata = artifacts = False
		for file in folder.iterdir():
			if file.name == "metadata.json":
				metadata = True
			if file.is_dir() and file.stem == "artifacts":
				artifacts = True
		return (metadata and artifacts)
			
	else: 
		return False

def store_uploaded_trace(scenario,run_num,folder):
	# adds the Scenario, Run number and .perfetto filename into the UPLOADED_TRACES list
	
	files =  folder.rglob("*.perfetto")
	perfetto = next(files).name

	UPLOADED_TRACES.append([scenario,run_num,perfetto])


def iterate_and_process_trace_folders(parent, args):
	# recursively iterates top down through a folder structure and moves the trace folders to dest
	if DEBUG:
		print(f'DEBUG: pull_trace folders given parent: {parent} and args: {args}')
	
	if parent.is_dir():
		if is_trace_folder(parent):
			# if this is a trace folder process it. 
			if DEBUG:
				print(f'DEBUG: {parent} is a trace folder. Starting Processing')
			process_trace_folder(parent, args)
		else:
			# if this currently isn't a trace folder iterate through all subfolders searching for them
			if DEBUG:
				print(f'DEBUG: {parent} is not a trace folder.')
			for child in parent.iterdir():
				if child.is_dir():
					iterate_and_process_trace_folders(child,args)
	
def process_trace_folder(folder, args):
	# Do Everything you need to process a single trace folder. 
	
	# First get the run_info from the metadata
	run_info = parse_metadata(folder)
	
	# False if the folder isn't a trace folder
	if run_info:
		
		# Upload the trace if needed. 
		if args.upload:
			print(f'Uploading all trace from: {folder}')
			os.system(f'{PNP_COMMAND} trace_manager --upload-local --upload "{folder}"')
			print("Finished uploading trace.")
		
		
		# delete perfetto files if needed
		if args.delete_perfetto:
			# Warning message and hold to verify upload
			print("\n\n=================================================ATTENTION=====================================================")
			input("Please verify that data uploaded correctly. Press ENTER to organize traces and delete perfetto file.")
			delete_perfetto(folder)
		
		# Find or make a nested folder structure for cwd / <HMD><branch> /Scenario <number>
		hmd_folder = f"{KNOWN_HMDS[run_info['device']]} {run_info['branch']}" # HMD Branch#
		scenario_folder = f"Scenario {run_info['scenario']}"	# Scenario #    TODO?? Change Scenario to BUC ??
		scenario_folder = Path.cwd().joinpath(hmd_folder,scenario_folder) # path/HMD Branch#/Scenario #
		scenario_folder = find_or_make_folder(scenario_folder)
		
		# Figure out how many run folders exist, and make a new folder with the right run number.
		run_folder, run_num = make_run_folder(scenario_folder)
		
		# Move the folder to the right run folder
		shutil.move(folder,run_folder)
		
		# Find or Create Summary.xlsx 
		
		summary_name = f"{KNOWN_HMDS[run_info['device']]} - Scenario {run_info['scenario']}.xlsx"
		summary_xlsx = find_summary_xlsx(scenario_folder, summary_name)
		
		if summary_xlsx == False:
			# find the right template
			summary_template = find_template(run_info['device'])
			summary_xlsx = scenario_folder.joinpath(summary_name)
			# Copy template to the scenario folder
			print(f"Copying summary template for {run_info['device']} to scenario {run_info['scenario']}") 
			shutil.copy(summary_template, summary_xlsx)
		
		# Import CSV data to summary_xlsx
		csv_file = find_file(run_folder, CSV_FILENAME)
		print(f"Adding CSV data for Scenario {run_info['scenario']} run {run_num} to {summary_xlsx.name}")
		add_csv_to_summary_xlsx(summary_xlsx, csv_file, run_num)
		
		store_uploaded_trace(run_info['scenario'],run_num,run_folder)
		
		# zip the run if needed
		if args.zip_runs:
			zip_in_place(run_folder)
	
	
	else:
		print(f'WARNING!! {folder} is not a trace folder. Ignoring')


	
def main():
	# gets arguments from the command line
	parser = argparse.ArgumentParser(description="Organizes pnp folders into scenario and run folders.")
	# Required argument of a filepath to work on
	parser.add_argument("-t", "--traces", dest='input_path', default=TRACE_FOLDER ,help="Filename of the traces folder for PNP, defaults to ~/.pnp/traces/")
	parser.add_argument("-u", "--upload", dest="upload", action='store_true', default=False, help="Uploads all the traces before organizing.")
	parser.add_argument("-d", "--delete_perfetto", dest="delete_perfetto", action='store_true', default=False, help="Deletes the large .perfetto files for easy uploading.")
	parser.add_argument("-z", "--zip", dest="zip_runs", action='store_true', default=False, help="Zips the runs after organizing.")
	args = parser.parse_args()
	# gets the path of the file that was passed
	if DEBUG:
		print(f'Given input_path: {args.input_path}')
	
	#dumb shit to fix powershell only sending the trailing " character in a path with spaces
	if str(args.input_path)[-1] == '"' and str(args.input_path)[0] != '"':
		print(f'Somehow got a path with an unmatched trailing " removing the extra "')
		input_path = Path(str(args.input_path)[:-1]).expanduser().resolve()
	else:
		input_path = Path(args.input_path).expanduser().resolve()
	
	if not input_path.exists():
		assert False, f'FATAL ERROR!!! {input_path} does not exist!! Did you give the wrong filename?'
	
	if input_path.is_dir() or input_path.suffix == ".zip":
		# makes sure everything is unzipped. 
		traces_folder = recursive_unzip(input_path)
	else:
		assert False, f'FATAL ERROR!!! {input_path} is not a directory or .zip file!!'	
	if DEBUG:
		print(f'traces_folder = {traces_folder}')
	
	# pull all the traces folders into the main traces folder
	iterate_and_process_trace_folders(traces_folder, args)
	if DEBUG:
		print('Finished pulling the trace folders.')
	
	print("Processed traces: ")
	for trace in UPLOADED_TRACES:
		print(f"Scenario {trace[0]} | Run {trace[1]} | {trace[2]}")
	
	if args.upload:
		print("Please verify that the uploaded traces appear in this list. ")
		print("Note: HCL uploads will not show below. ")
		os.system(f"{PNP_COMMAND} ls")
		
	print("Done organizing pnp folders and building summary spreadsheets.")
	print("Upload the scenario folders to google drive:  https://drive.google.com/drive/folders/18H_GK5mi5EcyUrqQue9RuUttV0KI0v09")
	
if __name__ == "__main__":
    main()

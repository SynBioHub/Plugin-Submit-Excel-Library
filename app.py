from flask import Flask, request, abort, send_file, jsonify
import os, shutil, glob, random, string
import pandas as pd
from Excel import read_library, quality_check, write_sbol
from sbol2 import *
#import all functions from .py files

app = Flask(__name__)

@app.route("/status")
def status():
    return("The Submit Excel Plugin Flask Server is up and running")



@app.route("/evaluate", methods=["POST"])
def evaluate():
    #uses MIME types
    #https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types
    
    eval_manifest = request.get_json(force=True)
    files = eval_manifest['manifest']['files']
    
    #temp
    cwd = os.getcwd()
    data = str(eval_manifest)
    # with open(os.path.join(cwd,"eval_manifest_recieved.txt"), 'w') as temp:
    #     temp.write(data) 
    
    eval_response_manifest = {"manifest":[]}
    
    for file in files:
        file_name = file['filename']
        file_type = file['type']
        file_url = file['url']
        
        ########## REPLACE THIS SECTION WITH OWN RUN CODE #################
        acceptable_types = {'application/vnd.ms-excel',
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}

        #could change what appears in the useful_types based on the file content
        useful_types = {}
        
        file_type_acceptable = file_type in acceptable_types
        file_type_useable = file_type in useful_types
        ################## END SECTION ####################################
        
        if file_type_acceptable:
            useableness = 2
        # elif file_type_useable:
        #     useableness = 1
        else:
            useableness = 0
        
        eval_response_manifest["manifest"].append({
            "filename": file_name,
            "requirement": useableness})
    # with open(os.path.join(cwd,"eval_manifest_response.txt"), 'w') as temp:
    #     temp.write(str(eval_response_manifest))   
    return jsonify(eval_response_manifest)


@app.route("/run", methods=["POST"])
def run():
    cwd = os.getcwd()
    
    #create a temporary directory
    temp_dir = tempfile.TemporaryDirectory()
    zip_in_dir_name = temp_dir.name
    
    #take in run manifest
    run_manifest = request.get_json(force=True)
    files = run_manifest['manifest']['files']
    
    #Read in template to compare to
    template_path = os.path.join(cwd, "templates", "darpa_template_blank.xlsx")
    
    #initiate response manifest
    run_response_manifest = {"results":[]}
    
    #Read in template to compare to
    file_path = os.path.join(cwd, "templates", "darpa_template_blank.xlsx")
    
    for a_file in files:
        try:
            file_name = a_file['filename']
            file_type = a_file['type']
            file_url = a_file['url']
            data = str(a_file)
           
            converted_file_name = f"{file_name}.converted"
            file_path_out = os.path.join(zip_in_dir_name, converted_file_name)
        
            ########## REPLACE THIS SECTION WITH OWN RUN CODE #################
            #Create own xml files using Excel.py etc.            
            start_row = 13
            nrows = 8
            description_row = 9
            

            filled_library, filled_library_metadata, filled_description = read_library(file_url, 
                        start_row = start_row, nrows = nrows, description_row = description_row)

            blank_library, blank_library_metadata, blank_description = read_library(file_path,  
                        start_row = start_row, nrows = nrows, description_row = description_row)

            quality_check(filled_library, blank_library, filled_library_metadata, 
                      blank_library_metadata, filled_description, blank_description,
                      nrows=nrows, description_row=description_row)

            ontology = pd.read_excel(file_path, header=None, sheet_name= "Ontology Terms", skiprows=3, index_col=0)
            ontology= ontology.to_dict("dict")[1]
            doc = write_sbol(filled_library, filled_library_metadata, filled_description, ontology)

            doc.write(file_path_out)
            ################## END SECTION ####################################
        
            # add name of converted file to manifest
            run_response_manifest["results"].append({"filename":converted_file_name,
                                    "sources":[file_name]})

            
        except Exception as e:
                print(e)
                abort(415)
            
    #create manifest file
    file_path_out = os.path.join(zip_in_dir_name, "manifest.json")
    with open(file_path_out, 'w') as manifest_file:
            manifest_file.write(str(run_response_manifest)) 
      
    
    with tempfile.NamedTemporaryFile() as temp_file:
        #create zip file of converted files and manifest
        shutil.make_archive(temp_file.name, 'zip', zip_in_dir_name)
        
        #delete zip in directory
        shutil.rmtree(zip_in_dir_name)
        
        #return zip file
        return send_file(f"{temp_file.name}.zip")

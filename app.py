from flask import Flask,request,render_template_string
from werkzeug.utils import secure_filename
import os
import shutil
from azure.storage.blob import BlobServiceClient
import fitz
import pandas
from datetime import datetime



app =Flask(__name__)
output_paths =[]

app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

AZURE_CONNECTION_STRING = os.environ["AZURE_CONNECTION_STRING"]

CONTAINER_NAME ="pdfstorage"

blob_service_client = BlobServiceClient.from_connection_string(AZURE_CONNECTION_STRING)
container_client = blob_service_client.get_container_client(CONTAINER_NAME)


HTML_TEMPLATE ='''<!DOCTYPE html>
<html>
<head>
    <title>PDF and Excel File Selector</title>
</head>
<body>
    <h2>Select PDF Files:</h2>
    <form method="POST" enctype="multipart/form-data" action="/process">
        <input type="file" name="pdfs" multiple required><br><br>

        <h2>Select Excel File:</h2>
        <input type="file" name="excel" required><br><br>

        <input type="submit" value="Process Files" style="background-color: black; color: white; padding: 10px;">
    </form>
</body>
</html>
'''


def upload_file_to_blob(local_path):
    blob_name = os.path.basename(local_path)
    with open(local_path, "rb") as data:
        container_client.upload_blob(name=blob_name, data=data, overwrite=True)
    return blob_name


def download_blob_to_local(blob_name, local_path):
    with open(local_path, "wb") as f:
        data = container_client.download_blob(blob_name)
        f.write(data.readall())


def download_processed_pdfs_from_blob(container_client, prefix="filled_form_", download_dir=os.path.join("Downloads", "pdfs"), limit=4):
    """
    Downloads the latest `limit` number of processed PDF files from blob storage to local folder.

    :param container_client: An instance of azure.storage.blob.ContainerClient
    :param prefix: Prefix of files to filter processed PDFs
    :param download_dir: Local directory to store downloaded files
    :param limit: Number of latest files to download
    """
    download_dir = os.path.abspath(download_dir)

    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # List all blobs with the given prefix
    blobs = container_client.list_blobs(name_starts_with=prefix)
    blobs = sorted(blobs, key=lambda x: x.last_modified, reverse=True)

    for blob in blobs[:limit]:
        blob_name = blob.name
        local_path = os.path.join(download_dir, os.path.basename(blob_name))
        print(f"Downloading: {blob_name} to {local_path}")

        with open(local_path, "wb") as f:
            data = container_client.download_blob(blob_name)
            f.write(data.readall())


def process_files(pdf_files, excel_file):
    if not pdf_files or not excel_file:
        return {"error": "Missing PDF or Excel file."}

    os.makedirs("temp", exist_ok=True)
    os.makedirs("downloads", exist_ok=True)

    # Save and upload original files
    uploaded_pdf_paths = []
    for file in pdf_files:
        filename = secure_filename(file.filename)
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(save_path)
        upload_file_to_blob(save_path)
        uploaded_pdf_paths.append(save_path)

    excel_filename = secure_filename(excel_file.filename)
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    excel_file.save(excel_path)
    upload_file_to_blob(excel_path)

    local_pdf_path = os.path.join("temp", os.path.basename(uploaded_pdf_paths[0]))
    local_excel_path = os.path.join("temp", os.path.basename(excel_path))
    download_blob_to_local(os.path.basename(uploaded_pdf_paths[0]), local_pdf_path)
    download_blob_to_local(os.path.basename(excel_path), local_excel_path)

    # Process and upload generated PDFs
    output_paths = auto_fill_pdf(local_pdf_path, local_excel_path)
    for path in output_paths:
        upload_file_to_blob(path)

    shutil.rmtree("temp", ignore_errors=True)
    shutil.rmtree("temp_outputs",ignore_errors=True)
    print("delete temp")
    download_processed_pdfs_from_blob(container_client)

    return {"message": f"{len(output_paths)} PDFs generated, uploaded and Downloaded successfully."}


def auto_fill_pdf(pdf_path,excel_path):
     doc =fitz.open(pdf_path)
     pd = pandas.read_csv(excel_path)

     patterns=["Please enter your name:","Option 1","Option 2","Option 3","Name of Dependent","Age of Dependent"]
     dict1={}
     page =doc[0]
     for block in page.get_text("dict")["blocks"]:
          for line in block["lines"]:
               for span in line['spans']:
                         print(span['text'],span['bbox'])
                         if(span['text'].strip() in patterns):
                              dict1[span['text'].strip()]=span['bbox']
          print(dict1)

     def get_new_cordinates_below(x0,y0,x1,y1):
          y1,y0 = x1-7,(y1-x1)+x1
          x1,x0 = y0,(y1-x1)+y0
          return x0,y0,x1,y1
     
     def enter_optionClick(x0,y0,x1,y1):
          check_box_size = 10

     # Calculate top-left corner of the checkbox (just to the left of the label)
          fin_cord_x = x0 - check_box_size - 3
          fin_cord_y = ((y0 + y1) / 2) - (check_box_size / 2)  # vertically centered

     # Define a rectangle (x0, y0, x1, y1)
          checkbox_rect = fitz.Rect(fin_cord_x, fin_cord_y,
                                   fin_cord_x + check_box_size,
                                   fin_cord_y + check_box_size)

     # Draw the filled black square (checkbox)
          page.draw_rect(checkbox_rect, fill=(0, 0, 0), color=(0, 0, 0))
     
     def name_enter(x0,y0,x1,y1,text):
          x0,y0,x1,y1 = get_new_cordinates_below(x0,y0,x1,y1)
          fin_cord_x,fin_cord_y = x1+45,y1+2
          page.insert_text((fin_cord_x,fin_cord_y),text,fontsize =12,fontname="helv",color=(0, 0, 0))

     def text_fileds_enter(x0,x1,y0,y1,text):
          fin_cord_x,fin_cord_y = (x1-x0)//2+40,y1-2
          print("insert.....")
          page.insert_text((fin_cord_x,fin_cord_y),text,fontsize =12)

     for ind,row in pd.iterrows():
          print(f"Row Index: {ind}")
          doc =fitz.open(pdf_path)
          page=doc[0]
          for col_name,value in row.items():
               print(col_name,value)
               if(col_name=="Name"):
                    name_enter(*dict1["Please enter your name:"],text=value)
               elif(col_name=="Options"):
                    lis=value.split(",")
                    for ele in lis:
                         enter_optionClick(*dict1[f"{ele}"])
               else:
                    print(dict1[f"{col_name}"])
                    text_fileds_enter(*dict1[f"{col_name}"],text=str(value))
               print("about to save....")
          
          output_dir = "temp_outputs"  # or any temporary folder
          os.makedirs(output_dir, exist_ok=True)

          timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
          output_filename = f"filled_form_{ind}_{timestamp}.pdf"
          output_path = os.path.join(output_dir, output_filename)
          doc.save(output_path)
          output_paths.append(output_path)
          doc.close()
          print(output_paths)
     return output_paths

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/process', methods=['POST'])
def handle_upload():
    pdf_files = request.files.getlist("pdfs")
    excel_file = request.files.get("excel")

    result = process_files(pdf_files, excel_file)
    if "error" in result:
        return f"<h3>Error: {result['error']}</h3><a href='/'>Back</a>"
    else:
        return f"<h3>{result['message']}</h3><a href='/'>Back</a>"


if __name__=="__main__":
    app.run(host="0.0.0.0", port=10000)

import os
from datetime import datetime
import shutil
from pathlib import Path
import yagmail
from docxtpl import DocxTemplate
from flask import Flask, request, send_file
from io import BytesIO
import zipfile


def clear_outputs_folder():
    outputs_path = Path.cwd() / "outputs"
    if outputs_path.exists():
        try:
            # Use shutil.rmtree to delete the entire folder and its contents
            shutil.rmtree(outputs_path)
            print("Folder deleted successfully.")
        except PermissionError:
            print("Permission denied. Please check if the folder is in use or if you have necessary permissions.")
        except FileNotFoundError:
            print("Folder not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


def make_documents(values):

    # set home path and download path
    home_path = Path.cwd()
    # Find all text files inside the templates directory
    path = Path(home_path / "templates")
    files = list(path.glob("*.docx"))

    # create text lists of sellers and buyers for POE
    sellers_dict = values["SELLERS_DICT"]
    buyers_dict = values["BUYERS_DICT"]
    values["SELLERS_LIST"] = make_parties_text_lists(sellers_dict)
    values["BUYERS_LIST"] = make_parties_text_lists(buyers_dict)

    # Format the contract date
    server_date = values["CONTRACT_DATE"]
    # First convert to a datetime object
    date_object = datetime.strptime(server_date, '%Y-%m-%d')
    # Format to the desired format: DD/MM/YYYY
    formatted_date = date_object.strftime('%d/%m/%Y')
    values["CONTRACT_DATE"] = formatted_date


    # Set the output directory
    outputs_path = home_path / "outputs"
    outputs_path.mkdir(parents=True, exist_ok=True)

    # Get the user's Downloads folder
    for file in files:
        doc = DocxTemplate(file)
        doc.render(values)
        doc.save(outputs_path / f"{file.name} {values['ADDRESS']}.docx")

    if len(sellers_dict) > 2 or len(buyers_dict) > 2:
        sellers_chunks = chunk_list_except_first(sellers_dict, 2)
        for more_sellers in sellers_chunks:
            index = 1
            doc = DocxTemplate(home_path / "templates" / "הצהרת נכונות פרטים.docx")
            values["SELLERS_DICT"] = more_sellers
            doc.render(values)
            doc.save(outputs_path / f"הצהרת נכונות פרטים - מוכרים נוספים{index}.docx")
            index += 1

        buyers_chunks = chunk_list_except_first(buyers_dict, 2)
        for more_buyers in buyers_chunks:
            index = 1
            doc = DocxTemplate(home_path / "templates" / "הצהרת נכונות פרטים.docx")
            values["BUYERS_DICT"] = more_buyers
            doc.render(values)
            doc.save(outputs_path / f"הצהרת נכונות פרטים - רוכשים נוספים{index}.docx")
            index += 1


def send_email(email):
    outputs_path = Path.cwd() / "outputs"
    documents = list(outputs_path.glob("*.docx"))
    try:
        # Initialize the yagmail.SMTP object
        yag = yagmail.SMTP('adavidov.deals.app@gmail.com', 'fjps gamx tkmm ycnu')

        # Subject and body of the email
        subject = "מסמכים אוטומטיים לעסקאות נדל״ן"
        body = "להודעה זו מצורפים המסמכים שהוכנו עבורך במערכת לייצור מסמכים אוטומטיים מבית א. דוידוב ושות׳, עורכי דין"
        yag.send(
            to=email,
            subject=subject,
            contents=body,
            attachments=documents  # List of file paths
        )
        # mongodb.update_docs_to_mail(deal_id, email)
        return {
            "result": "success",
            "result_code": 200,
            "email": email
        }
    except Exception as e:
        return {
            "result": f"Failed to send email: {str(e)}",
            "result_code": 300,
        }


def download_zip():
    outputs_path = Path.cwd() / "outputs"
    try:
        # Get list of all files in the "outputs" directory
        files = [f for f in os.listdir(outputs_path) if os.path.isfile(os.path.join(outputs_path, f))]

        if not files:
            return {'error': 'No files to zip in the outputs folder'}, 400

        # Create a ZIP file in memory
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file_name in files:
                file_path = os.path.join(outputs_path, file_name)
                zip_file.write(file_path, file_name)  # Add file to ZIP

        # Seek to the beginning of the buffer before sending
        zip_buffer.seek(0)

        # Send the ZIP file as an attachment
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='documents.zip'
        )
    except Exception as e:
        return {'error': str(e)}, 500


def chunk_list_except_first(lst, n):
    """Return all chunks of size n from lst, except the first chunk."""
    chunks = []
    # Create all chunks of size `n`
    for i in range(0, len(lst), n):
        chunks.append(lst[i:i + n])

    # Return all chunks except the first one
    return chunks[1:]  # Exclude the first chunk (index 0)


def make_parties_text_lists(data):
    # takes the sellers and the party data dictionary, and concatinate the text so it returns a full text list with all the sellers anf another text list with all the buyers
    parties_list = ""
    for i in range(len(data)):
        party = data[i]
        party_full_name = f'{party["LAST_NAME"]} {party["FIRST_NAME"]} {party["ID_KIND"]} {party["ID"]}'
        parties_list = f'{parties_list}{party_full_name} '
        add = "ו-" if i != len(data) - 1 else ""
        parties_list = f'{parties_list}{add}'
    return parties_list


app = Flask(__name__)


@app.route("/create", methods=["POST", "GET"])
def create():
    request_data = request.get_json()
    print(request_data)
    values = request_data["values"]
    output_method = request_data["output_method"]
    email_address = request_data["email_address"]

    make_documents(values=values)

    if output_method == "download":
        return download_zip(), 201
    if output_method == "mail":
        send_email(email_address)
        return "Docs sent to email : " + email_address, 201
    return "No method", 500


if __name__ == "__main__":
    app.run(debug=True)

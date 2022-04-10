import docx2pdf
from flask import Flask,render_template,request,send_from_directory
import os
from win32com.client import pythoncom

wdFormatPDF = 17

from werkzeug.utils import secure_filename, send_file


app=Flask(__name__)

@app.route("/",methods=["GET"])
def homePage():
    return render_template('index.html')

@app.route("/submitForm",methods=["POST"])
def wordDocRecieved():
    if request.method=="POST":
        try:
            for fileItem in os.listdir("./static/"):
                try:
                    if ".pdf" in fileItem:
                        os.remove(f"static/{fileItem}")
                    else:
                        pass
                except:
                    print("No file Present")
            
            listdire="./upload/"
            f=request.files["file1"]
            f.save("upload/"+secure_filename(f.filename))
            x=f.filename
            if " " in f.filename:
                x=f.filename.replace(" ","_")
            else:
                pass   
            
            ourFile=x.split(".")[0]
            ourFile=ourFile+".pdf"
            pythoncom.CoInitialize()      
            for files in os.listdir(listdire):
                print(files)
                x=files
                if " " in files:
                    x=files.replace(" ","_")
                fName=x.split(".")[0]
                docx2pdf.convert(f"upload/{x}",f"./static/{fName}.pdf")
            for files in os.listdir(listdire):
                os.remove(f"upload/{files}")   

            workingdir = os.path.abspath(os.getcwd())
            filepath = workingdir + "/static/"        
            return send_from_directory(filepath,f"{ourFile}")

        except Exception as e:
            print(e)

            return render_template("index.html")          


    

if __name__=="__main__":
    app.run(debug=True)



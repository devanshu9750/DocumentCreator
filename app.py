from __future__ import print_function
from flask import Flask, send_file, render_template, request
import pickle
import os.path
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import io
from datetime import date

SCOPES = ['https://www.googleapis.com/auth/drive']
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('docs', 'v1', credentials=creds)
drive_service = build('drive', 'v2', credentials=creds)
form1doc = "1JY2M9sWVuVTazdSFOwhj2DKhmppJeen8jKx_Kw4j8rA"
form2doc = "1mK1U0Zmp0Wt0FDd6BhARJJ79YDpsO9mF2CmGLNNI34U"
form3doc = "1rXj91XoFQ-lVpRZoGj7ODqBvhbAN022MBvfr96P8-3s"
form4doc = "1KwpUTGvFRMZwNz6tO67NFbJ7O18J0z1wpRdJXsHd6cY"
form5doc = "1bqYTVRCKF6jpRyj3tFNYudkmIw6TkWpT3kpaIZzXfTc"
form6doc = "1Zn9A1WbmHiREbB_VtKc-1BHUF9jAi_-8TCHBaCmyFWM"

app = Flask(__name__, template_folder='templates',)


def generateFile(main_id, requests):
    copy_title = 'Main Document (Copy)'
    body = {
        'name': copy_title
    }
    drive_response = drive_service.files().copy(
        fileId=main_id, body=body).execute()
    document_copy_id = drive_response.get('id')

    result = service.documents().batchUpdate(
        documentId=document_copy_id, body={'requests': requests}).execute()

    if os.path.exists("document.docx"):
        os.remove("document.docx")

    file_id = document_copy_id
    request = drive_service.files().export_media(fileId=file_id,
                                                 mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    fh = io.FileIO('document.docx', mode='wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))

    drive_service.files().delete(fileId=file_id).execute()


@app.route('/generateForm6', methods=["POST"])
def generateForm6():
    requests = [
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{employeer}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['employeer'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{addresse}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['addresse'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{owner}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['owner'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{addressp}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['addressp'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{city}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['city'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{mobile}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['mobile'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{director}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['director'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{company}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['company'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{din}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['din'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{fullname}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['fullname'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{fathername}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['fathername'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{address}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['address'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{email}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['email'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{para1}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['para1'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{para2}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['para2'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{dob}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['dob'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{occupation}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['occupation'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{nationality}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['nationality'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{pan}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['pan'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{date}}",
                    'matchCase':  'true'
                },
                'replaceText': str(date.today().day) + "/"+str(date.today().month)+"/"+str(date.today().year),
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{place}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['place'],
            }},
    ]
    generateFile(form6doc, requests)
    return send_file('document.docx')


@app.route('/generateForm5', methods=["POST"])
def generateForm5():
    requests = [
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{mobile}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['mobile'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{director}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['director'],
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{company}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['company'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{din}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['din'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{fullname}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['fullname'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{fathername}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['fathername'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{address}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['address'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{email}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['email'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{para1}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['para1'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{para2}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['para2'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{dob}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['dob'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{occupation}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['occupation'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{nationality}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['nationality'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{pan}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['pan'],
            }}, {
            'replaceAllText': {
                'containsText': {
                    'text': "{{date}}",
                    'matchCase':  'true'
                },
                'replaceText': str(date.today().day) + "/"+str(date.today().month)+"/"+str(date.today().year),
            }},
        {
            'replaceAllText': {
                'containsText': {
                    'text': "{{place}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['place'],
            }},
    ]
    generateFile(form5doc, requests)
    return send_file("document.docx")


@app.route('/generateForm4', methods=["POST"])
def generateForm4():
    requests = [{
        'replaceAllText': {
                'containsText': {
                    'text': "{{director}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['director'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{place}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['place'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{company}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['company'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{date}}",
                    'matchCase':  'true'
                },
                'replaceText': str(date.today().day) + "/"+str(date.today().month)+"/"+str(date.today().year),
                }}, ]
    generateFile(form4doc, requests)
    return send_file('document.docx')


@app.route('/generateForm3', methods=["POST"])
def generateForm3():
    requests = [{
        'replaceAllText': {
                'containsText': {
                    'text': "{{owner}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['owner'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{addressp}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['addressp'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{director}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['director'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{company}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['company'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{city}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['city'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{date}}",
                    'matchCase':  'true'
                },
                'replaceText': str(date.today().day) + "/"+str(date.today().month)+"/"+str(date.today().year),
                }}, ]
    generateFile(form3doc, requests)
    return send_file('document.docx')


@app.route('/generateForm2', methods=["POST"])
def generateForm2():
    requests = [{
        'replaceAllText': {
                'containsText': {
                    'text': "{{director}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['director'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{company}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['company'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{addresse}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['addresse'],
                }}, {
        'replaceAllText': {
                'containsText': {
                    'text': "{{employeer}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['employeer'],
                }}, ]
    generateFile(form2doc, requests)
    return send_file('document.docx')


@app.route('/generateForm1', methods=["POST"])
def generateForm1():
    requests = [{
        'replaceAllText': {
                'containsText': {
                    'text': "{{mobile}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['mobile'],
                }},
                {
        'replaceAllText': {
                'containsText': {
                    'text': "{{director}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['director'],
                }},
                {
        'replaceAllText': {
                'containsText': {
                    'text': "{{company}}",
                    'matchCase':  'true'
                },
                'replaceText': request.form['company'],
                }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{din}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['din'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{fullname}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['fullname'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{fathername}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['fathername'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{address}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['address'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{email}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['email'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{para1}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['para1'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{para2}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['para2'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{dob}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['dob'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{occupation}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['occupation'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{nationality}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['nationality'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{pan}}",
                'matchCase':  'true'
            },
            'replaceText': request.form['pan'],
        }}, {
        'replaceAllText': {
            'containsText': {
                'text': "{{date}}",
                'matchCase':  'true'
            },
            'replaceText': str(date.today().day) + "/"+str(date.today().month)+"/"+str(date.today().year),
        }},

    ]
    generateFile(form1doc, requests)
    return send_file('document.docx')


@app.route('/form1')
def form1():
    return render_template('form1.html')


@app.route('/form2')
def form2():
    return render_template("form2.html")


@app.route('/form3')
def form3():
    return render_template("form3.html")


@app.route('/form4')
def form4():
    return render_template("form4.html")


@app.route('/form5')
def form5():
    return render_template("form5.html")

@app.route('/form6')
def form6():
    return render_template("form6.html")


@app.route('/')
def selector():
    return render_template('selector.html')


if __name__ == "__main__":
    app.run()

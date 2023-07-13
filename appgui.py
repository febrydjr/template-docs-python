import datetime
from pathlib import Path

import PySimpleGUI as sg
from docxtpl import DocxTemplate

document_path = Path.cwd() / "lakerig.docx"
# document_path = Path(__file__).parent / "lakerig.docx"
doc = DocxTemplate(document_path)

today = datetime.datetime.today()
today_in_one_week = today + datetime.timedelta(days=7)

layout = [
    [sg.Text("Tanggal Sekarang:"), sg.Input(key="tanggal", do_not_clear=False)],
    [sg.Text("Nama Perusahaan:"), sg.Input(key="namaper", do_not_clear=False)],
    [sg.Text("Alamat Perusahaan"), sg.Input(key="alamatper", do_not_clear=False)],
    [sg.Text("Asal Iklan:"), sg.Input(key="ig", do_not_clear=False)],
    [sg.Text("Tanggal Iklan:"), sg.Input(key="tanggaliklan", do_not_clear=False)],
    [sg.Text("Posisi:"), sg.Input(key="bagian", do_not_clear=False)],
    [sg.Button("Create"), sg.Exit()],
]
window = sg.Window("Lamaran Generator(Instagram)", layout, element_justification="right")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Create":
        doc.render(values)
        output_path =  Path.cwd() / f"Lamaran Kerja - {values ['namaper']} - Febry Dharmawan Junior.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")

window.close()
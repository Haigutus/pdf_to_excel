#!/usr/bin/env python
# -*- coding: utf-8 -*-

import datetime
import base64
import uuid
from flask import send_file

import dash
from dash.dependencies import Input, Output, State
import dash_core_components as dcc
import dash_html_components as html
import dash_table_experiments as dt

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

#app = dash.Dash(__name__)

app.scripts.config.serve_locally = True

app.layout = html.Div([
    dcc.Upload(
        id       = "upload-files",
        accept   = "application/pdf",
        children = html.Div([u'Tiri siia või ', html.A('vali fail(id)')]),
        style={
            'width': '100%',
            'height': '120px',
            'lineHeight': '120px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        # Allow multiple files to be uploaded
        multiple=True
    ),
    html.Div(id='uploaded-files-list'),
    html.A("Hindamise Plank", id="template-link", href= "/template")
])


def parse_contents(contents, filename, date):

    file_save_name  = "incoming/{}.pdf".format(uuid.uuid4())
    date_string     = datetime.datetime.fromtimestamp(date).isoformat()

    with open(file_save_name, "wb") as downloaded_pdf:
        file_header, pdf_file = contents.split(",")
        downloaded_pdf.write(base64.b64decode(pdf_file))

    return {u"Sisend faili nimi":filename, u"Sisend fail loodud":date_string, u"Salvestatud faili nimi":file_save_name}


@app.callback(Output('uploaded-files-list', 'children'),
              [Input('upload-files', 'contents')],
              [State('upload-files', 'filename'),
               State('upload-files', 'last_modified')])
def update_output(list_of_contents, list_of_names, list_of_dates):
    if list_of_contents is not None:
        children = [parse_contents(c, n, d) for c, n, d in zip(list_of_contents, list_of_names, list_of_dates)]
        table = dt.DataTable(rows=children, filterable=False, row_selectable=False)
        return table

@app.server.route("/template")
def download_template():

    return send_file("test.pdf", mimetype="application/pdf", as_attachment=True, attachment_filename="hindamise_plank.pdf" )

if __name__ == '__main__':
    app.run_server(debug=True)
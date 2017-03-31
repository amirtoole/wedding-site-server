import logging

import pygsheets
from flask import Flask, jsonify, request

from settings import service_file_name, spreadsheet_key

app = Flask(__name__)

gc = pygsheets.authorize(service_file=service_file_name)
spreadsheet = gc.open_by_key(spreadsheet_key)


@app.route('/')
def root():
    return app.send_static_file('index.html')


@app.route('/<path:path>')
def static_proxy(path):
    # send_static_file will guess the correct MIME type
    return app.send_static_file(path)


@app.route('/api/code/<string:code>')
def getPeople(code):
    worksheet = spreadsheet.worksheet_by_title('website-form')
    # cell_list = worksheet.get_col(1, returnas='cell')
    matches = worksheet.find(code)
    response = []
    for m in matches:
        if m.col == 1:
            row = worksheet.get_row(m.row)
            if len(row) < 5:
                continue
            response.append({
                'row': m.row,
                'name': row[1],
                'child': row[2],
                'attending': row[3],
                'dinner': row[4],
                'comments': row[5]
            })

    return jsonify(response)


@app.route('/api/attendance/<string:code>', methods=['POST'])
def submitAttendance(code):
    data = request.json
    if not data or len(data) == 0:
        return server_error('Invalid data')

    worksheet = spreadsheet.worksheet_by_title('website-form')
    for person in data:
        row = worksheet.get_row(person['row'])
        if (row[0] == code):
            # match
            row[3] = person['attending']
            row[4] = person['dinner'] if person['attending'] == 'y' else ''
            row[5] = person['comments'].strip()
            worksheet.update_row(person['row'], row)
        else:
            return server_error("Something malicious is happening. Mismatched code")

    return 'Success'


@app.errorhandler(500)
def server_error(e):
    logging.exception('An error occurred during a request.')
    return """
    An internal error occurred: <pre>{}</pre>
    See logs for full stacktrace.
    """.format(e), 500


if __name__ == '__main__':
    # This is used when running locally. Gunicorn is used to run the
    # application on Google App Engine. See entrypoint in app.yaml.
    app.run(host='127.0.0.1', port=9000, debug=True)
# [END app]

# -*- coding: utf-8 -*-
##############################################################################
#
#    Copyright (C) 2012 Agile Business Group sagl (<http://www.agilebg.com>)
#    Copyright (C) 2012 Domsense srl (<http://www.domsense.com>)
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as published
#    by the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################
try:
    import json
except ImportError:
    import simplejson as json

import web.http as openerpweb

from web.controllers.main import ExcelExport

class ExcelExportView(ExcelExport):
    _cp_path = '/web/export/xls_view'

    def __getattribute__(self, name):
        if name == 'fmt':
            raise AttributeError()
        return super(ExcelExportView, self).__getattribute__(name)

    @openerpweb.httprequest
    def index(self, req, data, token):
        data = json.loads(data)
        model = data.get('model', [])
        domain = data.get('domain') or []
        rows = data.get('rows', None)
        column_headers = data.get('headers', [])
        column_headers_string = data.get('headers_string', [])
        column_headers_translated = dict(zip(column_headers, column_headers_string and column_headers_string or column_headers))
        
        # rows not supplied, meaning the user wanted to get all in domain
        if rows == None:
            # get ids in domain, data for ids
            ids = req.session.model(model).search(domain)
            record_data = req.session.model(model).read(ids, column_headers)
            
            # extract data into row list
            rows = []
            for record in record_data:
                record_values = []
                for column in column_headers:
                    val = record[column]
                    if isinstance(val, tuple):
                        val = val[1]
                    record_values.append(val)
                rows.append(record_values)

        return req.make_response(
            self.from_data([column_headers_translated[column] for column in column_headers], rows),
            headers=[
                ('Content-Disposition', 'attachment; filename="%s"'
                    % self.filename(model)),
                ('Content-Type', self.content_type)
            ],
            cookies={'fileToken': token}
        )

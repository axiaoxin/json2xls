#!/usr/bin/env python
#-*- coding:utf-8 -*-
import json
import requests
import os
import click
import xlwt
from xlwt import Workbook


class Json2Xls(object):

    def __init__(self, filename, url_or_json, method='get',
                 params=None, data=None, headers=None, form_encoded=False,
                 sheet_name='sheet0', title_style=None):
        self.sheet_name = sheet_name
        self.filename = filename
        self.url_or_json = url_or_json
        self.method = method
        self.params = params
        self.data = data
        self.headers = headers
        self.form_encoded = form_encoded

        self.__check_file_suffix()

        self.book = Workbook(encoding='utf-8')
        self.sheet = self.book.add_sheet(self.sheet_name)

        self.start_row = 0

        self.title_style = xlwt.easyxf(title_style or
                                       'font: name Arial, bold on;'
                                       'align: vert centre, horiz center;'
                                       'borders: top 1, bottom 1, left 1, right 1;'
                                       'pattern: pattern solid, fore_colour lime;'
                                       )

    def __check_file_suffix(self):
        suffix = self.filename.split('.')[-1]
        if '.' not in self.filename:
            self.filename += '.xls'
        elif suffix not in ['xls', 'xlsx']:
            raise Exception('filename format must be .xls/.xlsx')

    def __get_json(self):
        data = None
        try:
            data = json.loads(self.url_or_json)
        except:
            if os.path.isfile(self.url_or_json):
                with open(self.url_or_json, 'r') as source:
                    data = [json.loads(line) for line in source]
            else:
                try:
                    if self.method.lower() == 'get':
                        resp = requests.get(self.url_or_json,
                                            params=self.params,
                                            headers=self.headers)
                        data = resp.json()
                    else:
                        if os.path.isfile(self.data):
                            with open(self.data, 'r') as source:
                                self.data = [json.loads(line) for line in source]
                        if not self.form_encoded:
                            self.data = json.dumps(self.data)
                        resp = requests.post(self.url_or_json,
                                             data=self.data, headers=self.headers)
                        data = resp.json()
                except Exception as e:
                    print e
        return data

    def __fill_title(self, data):
        for index, key in enumerate(data.keys()):
            self.sheet.col(index).width = (len(key) + 1) * 256
            self.sheet.row(self.start_row).write(index,
                                                       key, self.title_style)
        self.start_row += 1

    def __fill_data(self, data):
        for index, value in enumerate(data.values()):
            if isinstance(value, basestring):
                value = value.encode('utf-8')
            else:
                value = str(value)
            width = self.sheet.col(index).width
            new_width = (len(value) + 1) * 256
            self.sheet.col(index).width = width if width > new_width else new_width
            self.sheet.row(self.start_row).write(index, str(value))

        self.start_row += 1

    def make(self, title_callback=None, body_callback=None):
        data = self.__get_json()
        if not isinstance(data, (dict, list)):
            raise Exception('bad json format')
        if isinstance(data, dict):
            data = [data]

        if title_callback != None:
            title_callback(self, data[0])
        else:
            self.__fill_title(data[0])

        if body_callback != None:
            for d in data:
                body_callback(self, d)
        else:
            for d in data:
                self.__fill_data(d)

        self.book.save(self.filename)


@click.command()
@click.argument('filename')
@click.argument('source')
@click.option('--method', '-m', default='get')
@click.option('--params', '-p', default=None)
@click.option('--data', '-d', default=None)
@click.option('--headers', '-h', default=None)
@click.option('--sheet', '-s', default='sheet0')
@click.option('--style', '-S', default=None)
@click.option('--form_encoded', '-f', is_flag=True)
def make(filename, source, method, params, data, headers, sheet, style, form_encoded):
    if isinstance(headers, basestring):
        headers = eval(headers)
    Json2Xls(filename, source, method=method, params=params,
             data=data, headers=headers, form_encoded=form_encoded, sheet_name=sheet,
             title_style=style).make()

if __name__ == '__main__':
    make()

# !/usr/bin/env python
# -*- coding:utf-8 -*-
import collections
import json
import os
from collections import OrderedDict
from functools import partial

import click
import requests
import xlwt
from xlwt import Workbook


class Json2Xls(object):
    """Json2Xls API 接口

    :param string xls_filename: 指定需要生成的的excel的文件名

    :param string json_data: 指定json数据来源，
       可以是一个返回json的url，
       也可以是一行json字符串，
       也可以是一个包含每行一个json的文本文件

    :param string method: 当数据来源为url时，请求url的方法，
       默认为get请求

    :param dict params: get请求参数，默认为 :py:class:`None`

    :param dict post_data: post请求参数，默认为 :py:class:`None`

    :param dict headers: 请求url时的HTTP头信息 (字典、json或文件)

    :param bool form_encoded: post请求时是否作为表单请求，默认为 :py:class:`False`

    :param string sheet_name: Excel的sheet名称，默认为 sheet0

    :param string title_style: Excel的表头样式，默认为 :py:class:`None`

    :param function json_dumps: 带ensure_ascii参数的json.dumps()，
                                默认参数值为 :py:class:`False`

    :param function json_loads: 带object_pairs_hook参数的json.loads()，默认为保持原始排序

    :param bool dumps: 生成excel时对表格内容执行json_dumps，默认为 :py:class:`False`
    """

    def __init__(self,
                 xls_filename,
                 json_data,
                 method='get',
                 params=None,
                 post_data=None,
                 headers=None,
                 form_encoded=False,
                 dumps=False,
                 sheet_name='sheet0',
                 title_style=None):
        self.json_dumps = partial(json.dumps, ensure_ascii=False)
        self.json_loads = partial(json.loads, object_pairs_hook=OrderedDict)

        self.sheet_name = sheet_name
        self.xls_filename = xls_filename
        self.json_data = json_data
        self.method = method
        self.params = params
        self.post_data = post_data
        self.headers = headers
        self.form_encoded = form_encoded
        self.dumps = dumps

        self.__check_file_suffix()

        self.book = Workbook(encoding='utf-8', style_compression=2)
        self.sheet = self.book.add_sheet(self.sheet_name)

        self.start_row = 0

        self.title_style = xlwt.easyxf(
            title_style or 'font: name Arial, bold on;'
            'align: vert centre, horiz center;'
            'borders: top 1, bottom 1, left 1, right 1;'
            'pattern: pattern solid, fore_colour lime;')

    def __check_file_suffix(self):
        suffix = self.xls_filename.split('.')[-1]
        if '.' not in self.xls_filename:
            self.xls_filename += '.xls'
        elif suffix != 'xls':
            raise Exception('filename suffix must be .xls')

    def __get_json(self):
        data = None
        try:
            data = self.json_loads(self.json_data)
        except Exception:
            if os.path.isfile(self.json_data):
                with open(self.json_data, 'r') as source:
                    try:
                        data = self.json_loads(source.read().replace('\n', ''))
                    except Exception:
                        source.seek(0)
                        data = [self.json_loads(line) for line in source]
            else:
                if self.headers and os.path.isfile(self.headers):
                    with open(self.headers) as headers_txt:
                        self.headers = self.json_loads(
                            headers_txt.read().decode('utf-8').replace(
                                '\n', ''))
                elif isinstance(self.headers, ("".__class__, u"".__class__)):
                    self.headers = self.json_loads(self.headers)
                try:
                    if self.method.lower() == 'get':
                        resp = requests.get(
                            self.json_data,
                            params=self.params,
                            headers=self.headers)
                        data = resp.json()
                    else:
                        if isinstance(
                            self.post_data,
                            ("".__class__, u"".__class__)) and os.path.isfile(
                                self.post_data):
                            with open(self.post_data, 'r') as source:
                                self.post_data = self.json_loads(
                                    source.read().replace(
                                        '\n', ''))
                        if not self.form_encoded:
                            self.post_data = self.json_dumps(self.post_data)
                        resp = requests.post(
                            self.json_data,
                            data=self.post_data,
                            headers=self.headers)
                        data = resp.json()
                except Exception as e:
                    print(e)
        return data

    def __fill_title(self, data):
        '''生成默认title'''
        data = self.flatten(data)
        for index, key in enumerate(data.keys()):
            if self.dumps:
                key = self.json_dumps(key)
            try:
                self.sheet.col(index).width = (len(key) + 1) * 256
            except Exception:
                pass
            self.sheet.row(self.start_row).write(index, key.decode('utf-8'),
                                                 self.title_style)
        self.start_row += 1

    def __fill_data(self, data):
        '''生成默认sheet'''
        data = self.flatten(data)
        for index, value in enumerate(data.values()):
            if self.dumps:
                value = self.json_dumps(value)
            self.auto_width(self.start_row, index, value)
            self.sheet.row(self.start_row).write(index, value)

        self.start_row += 1

    def auto_width(self, row, col, value):
        '''单元格宽度自动伸缩

        :param int row: 单元格所在行下标

        :param int col: 单元格所在列下标

        :param int value: 单元格中的内容
        '''

        try:
            self.sheet.row(row).height_mismatch = True
            # self.sheet.row(row).height = 0
            width = self.sheet.col(col).width
            new_width = min((len(value) + 1) * 256, 256 * 50)
            self.sheet.col(col).width = width \
                if width > new_width else new_width
        except Exception:
            pass

    def flatten(self, data_dict, parent_key='', sep='.'):
        '''对套嵌的dict进行flatten处理为单层dict

        :param dict data_dict: 需要处理的dict数据。

        :param str parent_key: 上层字典的key，默认为空字符串。

        :param str sep: 套嵌key flatten后的分割符， 默认为“.” 。
        '''

        out = {}

        def _flatten(x, parent_key, sep):
            if isinstance(x, collections.MutableMapping):
                for a in x:
                    _flatten(x[a], parent_key + a + sep, sep)
            elif isinstance(x, collections.MutableSequence):
                i = 0
                for a in x:
                    _flatten(a, parent_key + str(i) + sep, sep)
                    i += 1
            else:
                if not isinstance(x, ("".__class__, u"".__class__)):
                    x = str(x)
                out[parent_key[:-1].encode('utf-8')] = x

        _flatten(data_dict, parent_key, sep)
        return OrderedDict(out)

    def make(self, title_callback=None, body_callback=None):
        '''生成Excel。

        :param func title_callback: 自定义生成Execl表头的回调函数。
           默认为 :py:class:`None`，即采用默认方法生成

        :param func body_callback: 自定义生成Execl内容的回调函数。
           默认为 :py:class:`None`，即采用默认方法生成
        '''

        data = self.__get_json()
        if not isinstance(data, (dict, list)):
            raise Exception('The %s is not a valid json format' % type(data))
        if isinstance(data, dict):
            data = [data]

        if title_callback is not None:
            title_callback(self, data[0])
        else:
            self.__fill_title(data[0])

        if body_callback is not None:
            for d in data:
                body_callback(self, d)
        else:
            for d in data:
                self.__fill_data(d)

        self.book.save(self.xls_filename)


@click.command()
@click.argument('xls_filename')
@click.argument('json_data')
@click.option('--method', '-m', default='get')
@click.option('--params', '-p', default=None)
@click.option('--post_data', '-d', default=None)
@click.option('--headers', '-h', default=None)
@click.option('--sheet', '-s', default='sheet0')
@click.option('--style', '-S', default=None)
@click.option('--form_encoded', '-f', is_flag=True)
@click.option('--dumps', '-D', is_flag=True)
def make(xls_filename, json_data, method, params, post_data, headers, sheet,
         style, form_encoded, dumps):
    Json2Xls(
        xls_filename,
        json_data,
        method=method,
        params=params,
        post_data=post_data,
        headers=headers,
        form_encoded=form_encoded,
        dumps=dumps,
        sheet_name=sheet,
        title_style=style).make()


if __name__ == '__main__':
    make()

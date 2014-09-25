#!/usr/bin/env python
#-*- coding:utf-8 -*-

from json2xls.json2xls import Json2Xls

json_data = u'''[
    {"姓名": "John", "年龄": 30, "性别": "男"},
    {"姓名": "Alice", "年龄": 18, "性别": "女"}
]'''
obj = Json2Xls('tests/json_strlist_test.xls', json_data)
obj.make()


params = {
    'location': u'上海',
    'output': 'json',
    'ak': '5slgyqGDENN7Sy7pw29IUvrZ'
}
Json2Xls('tests/url_get_test.xls', "http://httpbin.org/get", params=params).make()


obj = Json2Xls('tests/json_list_test.xls', json_data='tests/list_data.json')
obj.make()


obj = Json2Xls('tests/json_line_test.xls', json_data='tests/line_data.json')
obj.make()


post_data = {
    'location': u'上海',
    'output': 'json',
    'ak': '5slgyqGDENN7Sy7pw29IUvrZ'
}
Json2Xls('tests/url_post_test1.xls', "http://httpbin.org/post", method='post', post_data=post_data, form_encoded=True).make()


post_data = 'tests/post_data.json'
Json2Xls('tests/url_post_test2.xls', "http://httpbin.org/post", method='post', post_data=post_data, form_encoded=True).make()


def title_callback(self, data):
    '''use one of data record to generate excel title'''
    self.sheet.write_merge(0, 0, 0, 3, 'title', self.title_style)
    self.sheet.write_merge(1, 2, 0, 0, 'tag', self.title_style)
    self.sheet.write_merge(1, 2, 1, 1, 'ner', self.title_style)
    self.sheet.write_merge(1, 1, 2, 3, 'comment', self.title_style)
    self.sheet.row(2).write(2, 'x', self.title_style)
    self.sheet.row(2).write(3, 'y', self.title_style)

    self.sheet.write_merge(0, 0, 4, 7, 'body', self.title_style)
    self.sheet.write_merge(1, 2, 4, 4, 'tag', self.title_style)
    self.sheet.write_merge(1, 2, 5, 5, 'ner', self.title_style)
    self.sheet.write_merge(1, 1, 6, 7, 'comment', self.title_style)
    self.sheet.row(2).write(6, 'x', self.title_style)
    self.sheet.row(2).write(7, 'y', self.title_style)

    self.start_row += 3

def body_callback(self, data):

    key1 = ['title', 'body']
    key2 = ['tag', 'ner', 'comment']

    col = 0
    for ii, i in enumerate(key1):
        for ij, j in enumerate(key2):
            if j != 'comment':
                value = ', '.join(data[ii][i][j])
                self.sheet.row(self.start_row).write(col, value)
                col += 1
            else:
                for x in data[ii][i][j].values():
                    self.sheet.row(self.start_row).write(col, x)
                    col += 1
    self.start_row += 1

data = 'tests/callback_data.json'
j = Json2Xls('tests/callback.xls', data)
j.make(title_callback=title_callback, body_callback=body_callback)


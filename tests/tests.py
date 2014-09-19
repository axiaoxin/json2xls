#!/usr/bin/env python
#-*- coding:utf-8 -*-

from json2xls import Json2Xls

url_or_json = '''[
    {"name": "John", "age": 30, "sex": "male"},
    {"name": "Alice", "age": 18, "sex": "female"}
]'''
obj = Json2Xls('test.xls', url_or_json)
obj.make()

params = {
    'location': u'上海',
    'output': 'json',
    'ak': '5slgyqGDENN7Sy7pw29IUvrZ'
}
Json2Xls('test2.xls', "http://api.map.baidu.com/telematics/v3/weather", params=params).make()




def title_callback(self, data):
    '''use one data record to generate excel title'''
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
                self.sheet.col(col).width = (len(value) + 1) * 256
                self.sheet.row(self.start_row).write(col, value)
                col += 1
            else:
                for x in data[ii][i][j].values():
                    width = self.sheet.col(col).width
                    new_width = (len(x) + 1) * 256
                    self.sheet.col(col).width = width if width > new_width else new_width
                    self.sheet.row(self.start_row).write(col, x)
                    col += 1
    self.start_row += 1


data = '''[
            [
                {
                    "title":
                        {
                            "tag": ["title_tag1", "title_tag2"],
                            "ner": ["title_ner1", "title_ner2"],
                            "comment": { "good": "100", "bad": "20"}
                        }
                },
                {
                    "body":
                        {
                            "tag": ["body_tag1", "body_tag2"],
                            "ner": ["body_ner1", "body_ner2"],
                            "comment": { "good": "85", "bad": "60"}
                        }
                }
            ]
        ]'''

j = Json2Xls('title_callback.xls', data)
j.make(title_callback=title_callback, body_callback=body_callback)

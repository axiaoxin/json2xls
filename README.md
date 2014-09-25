json2xls:Generate Excel by JSON data
====================================

[![](https://badge.fury.io/py/json2xls.png)](http://badge.fury.io/py/json2xls)
[![](https://pypip.in/d/json2xls/badge.png)](https://pypi.python.org/pypi/json2xls)

       _                 ____       _
      (_)___  ___  _ __ |___ \__  _| |___
      | / __|/ _ \| '_ \  __) \ \/ / / __|
      | \__ \ (_) | | | |/ __/ >  <| \__ \
     _/ |___/\___/|_| |_|_____/_/\_\_|___/
    |__/

generate excel by json string or json file or url which return a json

**docs** <http://json2xls.readthedocs.org/en/latest/>

**install**

    pip install json2xls

or

    python setup.py install


## command usage:

#### gen xls from json string

    json2xls tests/cmd_str_test.xls '{"a":"a", "b":"b"}'
    json2xls tests/cmd_str_test1.xls '[{"a":"a", "b":"b"},{"a":1, "b":2}]'

#### gen xls from file: whole file is a complete json data

    json2xls tests/cmd_list_test.xls "`cat tests/list_data.json`"

#### gen xls from file: each line is a json data

    json2xls tests/cmd_line_test.xls tests/line_data.json

#### gen xls from a url which respond a json

    json2xls tests/cmd_get_test.xls http://httpbin.org/get
    json2xls tests/cmd_post_test.xls http://httpbin.org/post -m post -d '"hello json2xls"' -h "{'X-Token': 'bolobolomi'}"

## coding usage:

#### gen xls from json string

    #!/usr/bin/env python
    #-*- coding:utf-8 -*-

    from json2xls.json2xls import Json2Xls

    json_data = u'''[
        {"姓名": "John", "年龄": 30, "性别": "男"},
        {"姓名": "Alice", "年龄": 18, "性别": "女"}
    ]'''
    obj = Json2Xls('tests/json_strlist_test.xls', json_data)
    obj.make()


#### gen xls from a url which respond a json by GET

    params = {
        'location': u'上海',
        'output': 'json',
        'ak': '5slgyqGDENN7Sy7pw29IUvrZ'
    }
    Json2Xls('tests/url_get_test.xls',
             "http://httpbin.org/get",
             params=params).make()


#### gen xls from a url which respond a json by POST

    post_data = {
        'location': u'上海',
        'output': 'json',
        'ak': '5slgyqGDENN7Sy7pw29IUvrZ'
    }
    Json2Xls('tests/url_post_test1.xls',
             "http://httpbin.org/post",
             method='post',
             post_data=post_data,
             form_encoded=True).make()

#### gen xls from a url which respond a json by POST with file post data

    post_data = 'tests/post_data.json'
    Json2Xls('tests/url_post_test2.xls',
             "http://httpbin.org/post",
             method='post',
             post_data=post_data,
             form_encoded=True).make()


#### gen xls from file: whole file is a complete json data （从文件内容为一个json列表的文件生成excel）

    Json2Xls('tests/json_list_test.xls', json_data='tests/list_data.json').make()

#### gen xls from file: each line is a json data （从文件内容为每行一个的json字符串的文件生成excel）

    obj = Json2Xls('tests/json_line_test.xls', json_data='tests/line_data.json')
    obj.make()

#### gen custom excel by define your title and body callback function

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


Default request method is `get`, request argument pass by `params`.
and the `post` method's request argument pass by `data`, you can use `-d` to pass request data in command line, the data should be json or file

Default only support one layer json to generate the excel, the nested json will be flattened. if you want custom it,
you can write the `title_callback` function and `body_callback` function, the pass them in the `make` function.
for the `body_callback`, you just need to care one line data's write way, json2xls default think the data are all the same.

The test demo data file is in tests dir.

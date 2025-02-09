# Changelog

## 0.6.5-rcc

- insert_multiple_lines_with_formatting 支持传入 texts, 控制插入同一段落有不同样式的文本。
- paragraph.rb 新增 set_multi_texts 方法。

```ruby
base_doc.bookmarks['notes'].insert_multiple_lines_with_formatting([
  {text: '第一句'},
  {text: '第二句', formatting: formatting},
  {text: '第三句'},
  {
    texts: [
      {text: '第一部分', formatting: {italic: true}},
      {text: '第二部分', formatting: {color: 'FF0000'},
      {text: '第三部分', formatting: {underline: true}}
    ]
  }
])
```

## 0.6.4-rcc

- insert_multiple_texts_with_formatting
- insert_multiple_lines_with_formatting

```ruby
# 文本的样式
formatting = {
  italic: false,
  underline: true,
  bold: true,
  font: 'Times New Roman',
  font_size: 20,
  color: 'FF0000'
}

# 段落的样式 -- xml中缩进的格式 <w:ind w:left="420" w:leftChars="0" w:firstLine="411" w:firstLineChars="196" />
paragraph_formatting = { ind: { left: 420, leftChars: 0, firstLine: 411 } }


base_doc.bookmarks['context'].insert_multiple_texts_with_formatting([
  {text: '工程信息订阅了'},
  {text: '2', formatting: formatting},
  {text: '套，价格是'},
  {text: '1299', formatting: formatting},
  {text: '元。'}
])

base_doc.bookmarks['notes'].insert_multiple_lines_with_formatting([
  {text: '第一句'},
  {text: '第二句', formatting: formatting.merge(paragraph_formatting)},
  {text: '第三句'}
])
```

## 0.6.3-rcc

### add feature

支持传入样式：
- insert_text_before(text, formatting = {})
- insert_text_after(text, formatting = {})
- insert_multiple_lines(text, formatting = {})

也可以针对单个 `textrun` 进行设置：
- textrun.set_text(content, formatting = {})

formatting 支持的参数为：
- :italic => boolean
- :bold => boolean
- :underline => boolean
- :font => 'font_name'
- :font_size => font_size
- :color => color as hex string e.g. 'FF0000' for red


## v0.6.2

### Bug fixes

- Fix `Docx::Document#to_s` fails when given file has `document22.xml.rels` [#112](https://github.com/ruby-docx/docx/pull/112), [#106](https://github.com/ruby-docx/docx/pull/106)

## v0.6.1

### Bug fixes

- Use `Zip::File#glob` to match any `document.xml` [#104](https://github.com/ruby-docx/docx/pull/104)

### Chores

- Enable Coverall's coverage report [#102](https://github.com/ruby-docx/docx/pull/102)
- Add table write example to README.md [#99](https://github.com/ruby-docx/docx/pull/99)
- Replace Travis CI build with GitHub Action [#98](https://github.com/ruby-docx/docx/pull/98)
- Add ruby 3.0 to versions for testing on Travis CI [#97](https://github.com/ruby-docx/docx/pull/97)

## v0.6.0

### Enhancements

- Added support for hyperlinks (implemented [#70](https://github.com/ruby-docx/docx/pull/70) again) by ollieh-m and gopeter [#92](https://github.com/ruby-docx/docx/pull/92)

### Chores

- Drop ruby 2.4 from supporeted versions by satoryu [#93](https://github.com/ruby-docx/docx/pull/93)
- Refactoring `spec_helper` by satoryu [#90](https://github.com/ruby-docx/docx/pull/90)
- Starts measuring code coverage with coveralls by satoryu [#88](https://github.com/ruby-docx/docx/pull/88)

## v0.5.0

### Enhancements

- Added opening streams and outputting to a stream [#66](https://github.com/ruby-docx/docx/pull/66)
- Added supports for Office 365 files [#85](https://github.com/ruby-docx/docx/pull/85)

### Bug fixes

- `Docx::Document` handles a docx file without styles.xml [#81](https://github.com/ruby-docx/docx/pull/81)
- Fixes insert text before after were switched [#84](https://github.com/ruby-docx/docx/pull/84)

## v0.4.0

### Enhancements

- Implement substitute method on TextRun class. [#75](https://github.com/ruby-docx/docx/pull/75)

### Improvements

- Updates dependencies. [#72](https://github.com/ruby-docx/docx/pull/72), [#77](https://github.com/ruby-docx/docx/pull/77)
- Fix: #paragraphs grabs paragraphs in tables. [#76](https://github.com/ruby-docx/docx/pull/76)
- Updates supported ruby versions. [#78](https://github.com/ruby-docx/docx/pull/78)

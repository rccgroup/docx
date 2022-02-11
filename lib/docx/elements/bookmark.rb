require 'docx/elements/element'

module Docx
  module Elements
    class Bookmark
      include Element
      attr_accessor :name

      def self.tag
        'bookmarkStart'
      end

      def initialize(node)
        @node = node
        @name = @node['w:name']
      end

      # Insert text before bookmarkStart node
      def insert_text_before(text, formatting = {})
        text_run = get_run_before
        text_run.set_text("#{text}#{text_run.text}", formatting)
      end

      # Insert text after bookmarkStart node
      def insert_text_after(text, formatting = {})
        text_run = get_run_after
        text_run.set_text("#{text_run.text}#{text}", formatting)
      end

      # 【新增方法】一整批插入到书签后面，可以为各部分文字定自己的样式。 Oliver.chen 2022-02-11
      #
      # base_doc.bookmarks['context'].insert_multiple_texts_after([
      #   { text: '工程信息订阅了' },
      #   { text: '2', formatting: formatting },
      #   { text: '套，价格是' },
      #   { text: '1299', formatting: formatting },
      #   { text: '元。' }
      # ])
      #
      # Note: 注意一个 bookmark 应该只插入一次。
      #
      # 具体的逻辑:
      # - 找到书签的位置 A
      # - 将插入数组的第一个元素E1，插入到书签 A 之后.(A <- E1)
      # - 剩余的元素以此追加,(A <- E1 <- E2 <- E3 <- E4...)
      def insert_multiple_texts_with_formatting(text_array = [])
        len = text_array.length
        return if len == 0

        # 第一个元素
        new_r = Containers::TextRun.create_with(self)
        new_r.insert_after(self)

        new_r.set_text(text_array[0][:text], text_array[0][:formatting])
        for i in 1..(len - 1)
          another_new_r = Containers::TextRun.create_with(self)
          another_new_r.set_text(text_array[i][:text], text_array[i][:formatting])
          another_new_r.insert_after(new_r)

          new_r = another_new_r
        end
      end

      # 【新增方法】 一整批插入到书签后面，可以为各个段落定自己的样式。 Oliver.chen 2022-02-11
      #
      # formatting = {
      #   italic: false,
      #   underline: true,
      #   bold: true,
      #   font: 'Times New Roman',
      #   font_size: 20,
      #   color: 'FF0000',
      #   ind: { left: 0, leftChars: 0 } --> 设置断落缩进的
      # }
      #
      # base_doc.bookmarks['notes'].insert_multiple_lines_with_formatting([
      #   {text: '第一句'},
      #   {text: '第二句', formatting: formatting},
      #   {text: '第三句'},
      #   {
      #     texts: [
      #       {text: '第一部分', formatting: {italic: true}},
      #       {text: '第二部分', formatting: {color: 'FF0000'},
      #       {text: '第三部分', formatting: {underline: true}}
      #     ]
      #   }
      # ])
      #
      # Note:
      # 1. 注意一个 bookmark 应该只插入一次。
      # 2. 段落有时需要控制缩进的问题
      def insert_multiple_lines_with_formatting(text_array = [])
        # Hold paragraphs to be inserted into, corresponding to the index of the strings in the text array
        paragraphs = []
        paragraph = self.parent_paragraph
        # Remove text from paragraph
        paragraph.blank!
        paragraphs << paragraph
        for i in 0...(text_array.size - 1)
          # Copy previous paragraph
          new_p = paragraphs[i].copy
          # Insert as sibling of previous paragraph
          new_p.insert_after(paragraphs[i])
          paragraphs << new_p
        end

        # Insert text into corresponding newly created paragraphs
        paragraphs.each_index do |index|
          # 如果有 texts，则以这个为主，忽略 text + formatting 属性。
          # 因为在 texts，会包含更加详细文字与样式。
          # 这样就可以支持同一个段落，不同文字带有不同的样式。
          # {
          #   texts: [
          #     { text: 'A', formatting: formatting },
          #     { text: 'B', formatting: formatting },
          #     { text: 'C', formatting: formatting },
          #   ]
          # }
          texts = text_array[index][:texts]
          if texts
            paragraphs[index].set_multi_texts(texts)
          else
            # 一整段使用相同的样式。
            paragraphs[index].apply_formatting(text_array[index][:formatting])
            paragraphs[index].set_text(text_array[index][:text], text_array[index][:formatting])
          end
        end
      end

      # 支持写入样式
      # insert multiple lines starting with paragraph containing bookmark node.
      def insert_multiple_lines(text_array, formatting = {})
        # Hold paragraphs to be inserted into, corresponding to the index of the strings in the text array
        paragraphs = []
        paragraph = self.parent_paragraph
        # Remove text from paragraph
        paragraph.blank!
        paragraphs << paragraph
        for i in 0...(text_array.size - 1)
          # Copy previous paragraph
          new_p = paragraphs[i].copy
          # Insert as sibling of previous paragraph
          new_p.insert_after(paragraphs[i])
          paragraphs << new_p
        end

        # Insert text into corresponding newly created paragraphs
        paragraphs.each_index do |index|
          paragraphs[index].set_text(text_array[index], formatting)
        end
      end

      # Get text run immediately prior to bookmark node
      def get_run_before
        # at_xpath returns the first match found and preceding-sibling returns siblings in the
        # order they appear in the document not the order as they appear when moving out from
        # the starting node
        if not (r_nodes = @node.xpath("./preceding-sibling::w:r")).empty?
          r_node = r_nodes.last
          Containers::TextRun.new(r_node)
        else
          new_r = Containers::TextRun.create_with(self)
          new_r.insert_before(self)
          new_r
        end
      end

      # Get text run immediately after bookmark node
      def get_run_after
        if (r_node = @node.at_xpath("./following-sibling::w:r"))
          Containers::TextRun.new(r_node)
        else
          new_r = Containers::TextRun.create_with(self)
          new_r.insert_after(self)
          new_r
        end
      end
    end
  end
end
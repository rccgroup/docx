require 'docx/containers/text_run'
require 'docx/containers/container'
require 'docx/formatting/paragraph_formatting'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element
        include ParagraphFormatting

        attr_reader :properties_tag

        def self.tag
          'p'
        end

        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {})
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        #【新增方法】使用 set_text 支持 formatting。Oliver.chen 2022-02-11
        def set_text(content, formatting = {})
          self.text = content
          text_runs.each { |tr| tr.apply_formatting(formatting) }
        end

        #【新增方法】设置不同样式的文字。Oliver.chen 2022-02-11
        #
        # texts = [
        #   { text: 'A', formatting: formatting },
        #   { text: 'B', formatting: formatting },
        #   { text: 'C', formatting: formatting },
        # ]
        # 注意：该方法会清掉原来段落的内容。不是追加内容，而是重写。
        def set_multi_texts(texts = [])
          # 先清掉原有的文字
          text_runs.each {|r| r.node.remove }

          len = texts.size
          return if len == 0

          new_r = TextRun.create_within(self)
          new_r.set_text(texts[0][:text], texts[0][:formatting])

          for i in 1..(len - 1)
            another_new_r = Containers::TextRun.create_with(self)
            another_new_r.set_text(texts[i][:text], texts[i][:formatting])
            another_new_r.insert_after(new_r)

            new_r = another_new_r
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end
        alias_method :text, :to_s

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = ''
          text_runs.each do |text_run|
            html << text_run.to_html
          end
          styles = { 'font-size' => "#{font_size}pt" }
          styles['text-align'] = alignment if alignment
          html_tag(:p, content: html, styles: styles)
        end


        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink').map { |r_node| Containers::TextRun.new(r_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def font_size
          size_tag = @node.xpath('w:pPr//w:sz').first
          size_tag ? size_tag.attributes['val'].value.to_i / 2 : @font_size
        end

        private

        # Returns the alignment if any, or nil if left
        def alignment
          # alignment_tag = @node.xpath('.//w:jc').first
          # alignment_tag ? alignment_tag.attributes['val'].value : nil
          formatting[:alignment]
        end

      end
    end
  end
end

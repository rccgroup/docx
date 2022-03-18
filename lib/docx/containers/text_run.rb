require 'docx/containers/container'
require 'docx/formatting/text_run_formatting'

module Docx
  module Elements
    module Containers
      class TextRun
        include Container
        include Elements::Element
        include TextRunFormatting

        DEFAULT_FORMATTING = {
          italic:    false,
          bold:      false,
          underline: false
        }

        attr_reader :text
        attr_reader :document_properties
        attr_reader :properties_tag

        def self.tag
          'r'
        end

        def initialize(node, document_properties = {})
          @node = node
          @document_properties = document_properties
          @text_nodes = @node.xpath('w:t').map {|t_node| Elements::Text.new(t_node) }
          @text_nodes = @node.xpath('w:t|w:r/w:t').map {|t_node| Elements::Text.new(t_node) }

          @properties_tag = 'rPr'
          @text       = parse_text || ''
          @font_size = @document_properties[:font_size]
        end

        # Set the text of text run with formatting
        def set_text(content, formatting = {})
          self.text = content
          apply_formatting(formatting)
        end

        #【新增】支持插入时，保留文字前后的空格。Oliver.chen 2022-03-18
        # Set text of text run
        def text=(content)
          if @text_nodes.size == 1
            # 保留文字的前后空格，而不是自动去掉
            @text_nodes.first.node["xml:space"] = "preserve"
            @text_nodes.first.content = content
          elsif @text_nodes.empty?
            new_t = Elements::Text.create_within(self)

            # 保留文字的前后空格，而不是自动去掉
            new_t.node["xml:space"] = "preserve"
            new_t.content = content
          end
        end

        # Returns text contained within text run
        def parse_text
          @text_nodes.map(&:content).join('')
        end

        # Substitute text in text @text_nodes
        def substitute(match, replacement)
          @text_nodes.each do |text_node|
            text_node.content = text_node.content.gsub(match, replacement)
          end
        end

        def to_s
          @text
        end

        # Return text as a HTML fragment with formatting based on properties.
        def to_html
          html = @text
          html = html_tag(:em, content: html) if italicized?
          html = html_tag(:strong, content: html) if bolded?

          styles = {}
          styles['text-decoration'] = 'underline' if underlined?
          # No need to be granular with font size down to the span level if it doesn't vary.
          # styles['font-size'] = "#{font_size}pt" if font_size != @font_size
          styles['font-size'] = "#{font_size}pt" if font_size != @font_size
          styles['font-family'] = %Q["#{formatting[:font]}"] if formatting[:font]
          styles['color'] = "##{formatting[:color]}" if formatting[:color]

          html = html_tag(:span, content: html, styles: styles) unless styles.empty?
          html = html_tag(:a, content: html, attributes: {href: href, target: "_blank"}) if hyperlink?
          return html
        end

        def italicized?
          formatting[:italic]
        end

        def bolded?
          formatting[:bold]
        end

        def underlined?
          formatting[:underline]
        end

        def hyperlink?
          @node.name == 'hyperlink'
        end

        def href
          @document_properties[:hyperlinks][hyperlink_id]
        end

        def hyperlink_id
          @node.attributes['id'].value
        end

        def font_size
          size_tag = @node.xpath('w:rPr//w:sz').first
          size_tag ? size_tag.attributes['val'].value.to_i / 2 : @font_size
        end
      end
    end
  end
end

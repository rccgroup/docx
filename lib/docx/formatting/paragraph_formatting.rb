require 'docx/formatting/formatting'

module Docx
  module ParagraphFormatting
    include Formatting

    def apply_formatting(formatting)
      formatting ||= {}

      # 设置段落属性 -- 缩进
      if (formatting[:ind])
        ind_node = add_property('ind')

        # ind_node["w:left"] = "0"
        # ind_node["w:leftChars"] = "0"
        # ind_node["w:firstLine"] = "0"
        # ind_node["w:firstLineChars"] = "0"
        # <w:ind w:left="420" w:leftChars="0" w:firstLine="411" w:firstLineChars="196" />
        formatting[:ind].each do |key, value|
          ind_node["w:#{key}"] = value
        end
      end

      if (formatting[:alignment])
        alignment_node = add_property('jc')
        alignment_node['w:val'] = formatting[:alignment]
      end
    end

    def parse_formatting
      formatting = {}

      alignment_node = node.at_xpath('.//w:jc')
      formatting[:alignment] = alignment_node ? alignment_node['w:val'] : nil

      formatting
    end
    alias_method :formatting, :parse_formatting

    def self.default_formatting
      {
        alignment: nil
      }
    end
  end
end

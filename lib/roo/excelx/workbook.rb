require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Workbook < Excelx::Extractor

      class Label
        attr_reader :sheet, :name, :row, :col
        attr_reader :coordinates

        def initialize(name, sheet, coordinates)
          @name = name
          @sheet = sheet.gsub('\'', '')
          col, row = coordinates.split('$')
          @row = row.to_i
          @col = ::Roo::Utils.letter_to_number(col)
          @coordinates = coordinates.gsub("$", '')
        end

        def key
          [row, col].compact
        end
      end

      def initialize(path)
        super
        fail ArgumentError, 'missing required workbook file' unless doc_exists?
      end

      def sheets
        doc.xpath('//sheet')
      end

      # aka labels
      def defined_names
        Hash[doc.xpath('//definedName').map do |defined_name|
          # "Sheet1!$C$5"
          sheet, coordinates = defined_name.text.split('!$', 2)          
          name = defined_name['name']
          [name, Label.new(name, sheet, coordinates)]
        end]
      end

      def base_date
        @base_date ||=
        begin
          # Default to 1900 (minus one day due to excel quirk) but use 1904 if
          # it's set in the Workbook's workbookPr
          # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
          result = Date.new(1899, 12, 30) # default
          doc.css('workbookPr[date1904]').each do |workbookPr|
            if workbookPr['date1904'] =~ /true|1/i
              result = Date.new(1904, 01, 01)
              break
            end
          end
          result
        end
      end
    end
  end
end

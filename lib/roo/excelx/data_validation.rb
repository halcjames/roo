module Roo
	class Excelx
		class DataValidation
			attr_accessor :allow_blank, :prompt, :type, :sqref, :source

			def self.load_from_node(data_validation_node)
				allow_blank = data_validation_node.attribute('allowBlank')
				prompt = data_validation_node.attribute('prompt')
				type = data_validation_node.attribute('type')
				sqref = data_validation_node.attribute('sqref')
				source = data_validation_node.at('formula1')

				self.new(allow_blank: allow_blank && allow_blank.value,
									prompt: prompt && prompt.value,
									type: type && type.value,
									sqref: sqref && (sqref.value || sqref.content),
									source: source && source.content)
			end

			def in_range?(row, col)
				return if sqref_range.empty?
				sqref_range.find do |key, row_range|
					column_found = key.is_a?(Range) ? key.include?(col) : key == col
					row_found = row_range.include?(row)

					column_found && row_found
				end
			end
					
			private
			def initialize(**attrs)
				attrs.each { |property, value| send("#{property}=", value)}
			end

 			def sqref_range
        @sqref_range ||= begin
          # "BH5:BH271 BI5:BI271"
          if sqref.nil?
            []
          else
            !sqref.include?(':') && !sqref.include?(' ') ? build_single_range(sqref) : build_multiple_range(sqref)
          end
        end
     end


      def build_multiple_range(sqref)
        sqref.split(' ').map do |splitted_by_space_sqref|
          # ["BH5:BH271, "BI5:BI271"]
          if splitted_by_space_sqref.is_a?(Array)
            splitted_by_space_sqref.map do |sqref|
              build_range(splitted_by_space_sqref)
            end
          else
            # "BH5:BH271"
            build_range(splitted_by_space_sqref)
          end
        end.to_h
      end

      def build_single_range(sqref)
        cell_letter = sqref.gsub(/[\d]/, '')
        index = sqref.gsub(/[^\d]/, '').to_i
        { ::Roo::Utils::letter_to_number(cell_letter) => (index..index)}
      end

      def build_range(sqref)
        splitted_sqref = sqref.gsub(/[\d]/, '')
        sqref_without_letters = sqref.gsub(/[^\d:]/, '')
        if sqref.include?(':')
          start_letter, end_letter = splitted_sqref.split(':')
          start_index, end_index = sqref_without_letters.split(':').map(&:to_i)
          [(::Roo::Utils::letter_to_number(start_letter)..::Roo::Utils::letter_to_number(end_letter)),(start_index..end_index)]
        else
          [(::Roo::Utils::letter_to_number(splitted_sqref)..::Roo::Utils::letter_to_number(splitted_sqref)),(sqref_without_letters.to_i..sqref_without_letters.to_i)]
        end
      end
		end
	end
end
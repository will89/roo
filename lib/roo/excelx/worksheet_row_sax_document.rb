# frozen_string_literal: true
module Roo
  class Excelx
    class WorksheetRowSaxDocument < Nokogiri::XML::SAX::Document

      def initialize(shared, row_block)
        @shared = shared
        @row_block = row_block
        @row_started = false
        @current_row = nil

        @column_started = false
        # @column_content = nil
        @column_coordinate = nil
        @column_type = nil
        @column_style = nil

        @v_started = false
        @column_value = nil

        @f_started = false
        @formula = nil

        @tag_characters = nil
      end

      def start_element(name, attrs = [])
        # puts "#{name} #{attrs.join(', ')} started!"
        case name
        when 'row'
          @current_row = []
        when 'c'
          @column_started = true
          attrs.each do |attr_name, attr_value|
            case attr_name
            when 'r'
              @column_coordinate = attr_value
            when 't'
              @column_type = attr_value
            when 's'
              @column_style = attr_value
            end
          end
        when 'v'
          @v_started = true
        when 'f'
          @f_started = true
        end
      end

      def end_element(name)
        # puts "#{name} including tag #{@tag_characters} ended"
        case name
        when 'row'
          @row_block.call(@current_row)
          @current_row = nil
        when 'c'
          @current_row << create_cell_from_value if @current_row && @column_value
          @column_started = false
          @column_coordinate = nil
          @column_value = nil
          @column_type = nil
          @column_style = nil
          # @column_content = nil
          @formula = nil
        when 'v'
          @v_started = false
          @column_value = @tag_characters
        when 'f'
          @f_started = false
          @formula = @tag_characters # MAYBE FORMULA?
        end

        @tag_characters = nil
      end

      def characters(string)
        @tag_characters = string
      end

      private

      # Take an xml row and return an array of Excelx::Cell objects
      # optionally pad array to header width(assumed 1st row).
      # takes option pad_cells (boolean) defaults false
      def cells_for_row_element(row_element, options = {})
        return [] unless row_element
        cell_col = 0
        cells = []
        @sheet.each_cell(row_element) do |cell|
          cells.concat(pad_cells(cell, cell_col)) if options[:pad_cells]
          cells << cell
          cell_col = cell.coordinate.column
        end
        cells
      end

      def cell_value_type(type, format)

      end

      # Internal: Creates a cell based on an XML clell..
      #
      # cell_xml - a Nokogiri::XML::Element. e.g.
      #             <c r="A5" s="2">
      #               <v>22606</v>
      #             </c>
      # hyperlink - a String for the hyperlink for the cell or nil when no
      #             hyperlink is present.
      #
      # Examples
      #
      #    cells_from_xml(<Nokogiri::XML::Element>, nil)
      #    # => <Excelx::Cell::String>
      #
      # Returns a type of <Excelx::Cell>.
      def cell_from_xml(cell_xml, hyperlink, coordinate = nil)
        coordinate ||= ::Roo::Utils.extract_coordinate(cell_xml[COMMON_STRINGS[:r]])
        cell_xml_children = cell_xml.children
        return Excelx::Cell::Empty.new(coordinate) if cell_xml_children.empty?

        # NOTE: This is error prone, to_i will silently turn a nil into a 0.
        #       This works by coincidence because Format[0] is General.
        style = cell_xml[COMMON_STRINGS[:s]].to_i
        formula = nil

        cell_xml_children.each do |cell|
          case cell.name
          when 'is'
            content = String.new
            cell.children.each do |inline_str|
              if inline_str.name == 't'
                content << inline_str.content
              end
            end
            unless content.empty?
              return Excelx::Cell.cell_class(:string).new(content, formula, style, hyperlink, coordinate)
            end
          when 'f'
            formula = cell.content
          when 'v'
            format = style_format(style)
            value_type = cell_value_type(cell_xml[COMMON_STRINGS[:t]], format)

            return create_cell_from_value(value_type, cell, formula, format, style, hyperlink, coordinate)
          end
        end

        Excelx::Cell::Empty.new(coordinate)
      end

      def create_cell_from_value# (value_type, cell, formula, format, style, hyperlink, coordinate)
        # NOTE: format.to_s can replace excelx_type as an argument for
        #       Cell::Time, Cell::DateTime, Cell::Date or Cell::Number, but
        #       it will break some brittle tests.
        # excelx_type = [:numeric_or_formula, format.to_s]
        coordinate = ::Roo::Utils.extract_coordinate(@column_coordinate)
        hyperlink = nil
        style = @column_style
        format = @shared.styles.style_format(style).to_s
        excelx_type = [format]
        value_type = case @column_type
                     when 's'
                       :shared
                     when 'b'
                       :boolean
                     when 'str'
                       :string
                     when 'inlineStr'
                       :inlinestr
                     else
                       Excelx::Format.to_type(format)
                     end

        # NOTE: There are only a few situations where value != cell.content
        #       1. when a sharedString is used. value = sharedString;
        #          cell.content = id of sharedString
        #       2. boolean cells: value = 'TRUE' | 'FALSE'; cell.content = '0' | '1';
        #          But a boolean cell should use TRUE|FALSE as the formatted value
        #          and use a Boolean for it's value. Using a Boolean value breaks
        #          Roo::Base#to_csv.
        #       3. formula
        case value_type
        when :shared
          cell_content = @column_value.to_i
          value = shared_strings.use_html?(cell_content) ? shared_strings.to_html[cell_content] : shared_strings[cell_content]
          Excelx::Cell.cell_class(:string).new(value, @formula, style, hyperlink, coordinate)
        when :boolean, :string
          value = @column_value
          Excelx::Cell.cell_class(value_type).new(value, @formula, style, hyperlink, coordinate)
        when :time, :datetime
          cell_content = @column_value.to_f
          # NOTE: A date will be a whole number. A time will have be > 1. And
          #      in general, a datetime will have decimals. But if the cell is
          #      using a custom format, it's possible to be interpreted incorrectly.
          #      cell_content.to_i == cell_content && standard_style?=> :date
          #
          #      Should check to see if the format is standard or not. If it's a
          #      standard format, than it's a date, otherwise, it is a datetime.
          #      @styles.standard_style?(style_id)
          #      STANDARD_STYLES.keys.include?(style_id.to_i)
          cell_type = if cell_content < 1.0
                        :time
                      elsif (cell_content - cell_content.floor).abs > 0.000001
                        :datetime
                      else
                        :date
                      end
          base_value = cell_type == :date ? base_date : base_timestamp
          Excelx::Cell.cell_class(cell_type).new(cell_content, @formula, excelx_type, style, hyperlink, base_value, coordinate)
        when :date
          Excelx::Cell.cell_class(:date).new(@column_value, @formula, excelx_type, style, hyperlink, base_date, coordinate)
        else
          # Excelx::Cell.cell_class(:number).new(cell.content, formula, excelx_type, style, hyperlink, coordinate)
          # begin
          Excelx::Cell.cell_class(:number).new(@column_value, @formula, excelx_type, style, hyperlink, coordinate)
          # rescue => e
          #   puts 'crap'
          # end
        end
      end

      def base_date
        @shared.base_date
      end

      def base_timestamp
        @shared.base_timestamp
      end

      def shared_strings
        @shared.shared_strings
      end
    end
  end
end
